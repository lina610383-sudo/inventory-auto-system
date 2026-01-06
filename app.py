import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import io
import re

# 頁面配置
st.set_page_config(page_title="領用單自動化生成系統", layout="wide")
st.title("🚀 領用單流程自動化系統")

def get_col_idx_by_id(ws, header_row_idx, target_id):
    """
    在模板標題列搜尋工號（掛帳人 ID），返回欄位索引 (1-based)
    """
    if not target_id: 
        return None
    target_id = str(target_id).strip().upper()
    # 遍歷所有欄位
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row_idx, column=col).value
        if val and str(val).strip().upper() == target_id:
            return col
    return None

def get_row_idx_by_pn(ws, pn_col_idx, target_pn):
    """
    在模板料號欄搜尋 PN，返回行索引 (1-based)
    """
    if not target_pn: 
        return None
    target_pn = str(target_pn).strip().upper()
    # 從模板第 1 行開始搜尋到最大行數
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=pn_col_idx).value
        if val and str(val).strip().upper() == target_pn:
            return row
    return None

def process_excel(file):
    try:
        # 1. 載入活頁簿
        wb = openpyxl.load_workbook(file)
        sheet_names = wb.sheetnames
        
        # 搜尋包含「領用明細_數字」且結尾為「(未開單)」的分頁
        pattern = r".*領用明細_(\d+).*\(未開單\)"
        matches = []
        for s in sheet_names:
            m = re.search(pattern, s)
            if m:
                matches.append((m.group(1), s))
        
        if not matches:
            st.error("找不到符合格式的分頁！請確認分頁名稱包含『領用明細_日期』且結尾為『(未開單)』")
            return None, None
        
        latest_date, target_sheet_name = sorted(matches, key=lambda x: x[0])[-1]
        st.info(f"📍 偵測到目標明細：{target_sheet_name}")
        
        # 2. 讀取明細資料與掛帳人清單
        # header=1 假設明細標題在第 2 列
        detail_df = pd.read_excel(file, sheet_name=target_sheet_name, header=1)
        
        if "掛帳人清單" not in sheet_names:
            st.error("找不到『掛帳人清單』分頁！")
            return None, None
            
        payer_df = pd.read_excel(file, sheet_name="掛帳人清單")
        payer_df.iloc[:, 0] = payer_df.iloc[:, 0].ffill() # 補全 IEC/ICC 類型
        
        # 建立領用人與掛帳資訊的字典
        payer_map = {}
        for _, row in payer_df.iterrows():
            name = str(row['領用人']).strip()
            unit_type = str(row.iloc[0]).strip().upper() 
            if name and name != 'nan':
                payer_map[name] = {
                    'type': "IEC" if "IEC" in unit_type else "ICC",
                    'id': str(row['掛帳人']).strip()
                }

        # 3. 預分析：判斷本次明細包含哪些類型 (IEC 或 ICC)
        valid_person_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        types_needed = set()
        
        for _, row in detail_df.iterrows():
            for person in valid_person_cols:
                qty = row[person]
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    info = payer_map[str(person).strip()]
                    types_needed.add(info['type'])

        # 4. 準備產出分頁 (僅針對有資料的類型建立)
        output_ws_dict = {}
        for t in types_needed:
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_領用單_{latest_date}"
                output_ws_dict[t] = new_ws
            else:
                st.warning(f"⚠️ 找不到模板：{tmpl_name}，無法產出該類型分頁。")

        # 5. 雙向對位回填資料
        filled_count = 0
        for _, row in detail_df.iterrows():
            item_pn = row.get('IEC PN')
            if pd.isna(item_pn): continue
            
            for person in valid_person_cols:
                qty = row[person]
                
                # 僅處理有領用數量的資料
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    person_name = str(person).strip()
                    info = payer_map[person_name]
                    target_type = info['type']
                    
                    if target_type in output_ws_dict:
                        ws = output_ws_dict[target_type]
                        
                        # 搜尋座標
                        target_row = get_row_idx_by_pn(ws, 5, item_pn)   # 預設料號在 E 欄 (5)
                        target_col = get_col_idx_by_id(ws, 5, info['id']) # 預設工號在第 5 列
                        
                        if target_row and target_col:
                            ws.cell(row=target_row, column=target_col, value=qty)
                            filled_count += 1
                        else:
                            # 輔助提示座標遺失問題
                            if not target_row:
                                st.warning(f"⚠️ 在 {target_type} 模板搜尋不到料號: {item_pn}")
                            if not target_col:
                                st.warning(f"⚠️ 在 {target_type} 模板搜尋不到工號: {info['id']} ({person_name})")

        # 6. 修改狀態並儲存
        if filled_count > 0:
            ws_orig = wb[target_sheet_name]
            ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
            
            output = io.BytesIO()
            wb.save(output)
            return output.getvalue(), latest_date
        else:
            st.error("雖然偵測到需求，但無法在模板中找到對應的座標填入資料，請檢查模板與明細的 PN/工號。")
            return None, None

    except Exception as e:
        st.error(f"執行發生錯誤：{str(e)}")
        return None, None

# --- Streamlit 使用者介面 ---
uploaded_file = st.file_uploader("📂 請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    if st.button("✨ 依照預設格式產生領用單"):
        with st.spinner("正在分析資料並精準對位填表..."):
            processed_data, date = process_excel(uploaded_file)
            if processed_data:
                st.success(f"處理完成！已依照實際領用類型產出分頁。")
                st.download_button(
                    label="📥 下載產出檔案",
                    data=processed_data,
                    file_name=f"領用單產出_{date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
