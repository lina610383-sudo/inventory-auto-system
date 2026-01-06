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
    # 從模板第 1 行開始搜尋以確保彈性，通常料號清單在標題下方
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
        
        # 彈性搜尋分頁名稱：包含「領用明細_數字」且結尾有「(未開單)」
        pattern = r".*領用明細_(\d+).*\(未開單\)"
        matches = []
        for s in sheet_names:
            m = re.search(pattern, s)
            if m:
                matches.append((m.group(1), s))
        
        if not matches:
            st.error("找不到符合格式的分頁！請確認分頁名稱包含『領用明細_日期』且結尾為『(未開單)』")
            return None, None
        
        # 取得最新日期的分頁
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
        
        # 建立人名與掛帳資訊的對應
        payer_map = {}
        for _, row in payer_df.iterrows():
            name = str(row['領用人']).strip()
            unit_type = str(row.iloc[0]).strip().upper() 
            if name and name != 'nan':
                payer_map[name] = {
                    'type': "IEC" if "IEC" in unit_type else "ICC",
                    'id': str(row['掛帳人']).strip()
                }

        # 3. 準備產出分頁 (直接複製您的格式模板)
        output_ws_dict = {}
        for t in ['IEC', 'ICC']:
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                # 重點：複製模板分頁，這會保留您所有的格式與樣式
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_領用單_{latest_date}"
                output_ws_dict[t] = new_ws
            else:
                st.warning(f"⚠️ 找不到模板：{tmpl_name}，將跳過此類型的生成。")

        # 4. 雙向對位回填資料
        # 取得明細中存在的領用人欄位
        valid_person_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        filled_count = 0

        for _, row in detail_df.iterrows():
            item_pn = row.get('IEC PN')
            if pd.isna(item_pn): continue
            
            for person in valid_person_cols:
                qty = row[person]
                
                # 只有當數量大於 0 時才處理
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    info = payer_map[str(person).strip()]
                    target_type = info['type']
                    
                    if target_type in output_ws_dict:
                        ws = output_ws_dict[target_type]
                        
                        # 縱向：根據 IEC PN 找行 (搜尋 E 欄，即第 5 欄)
                        target_row = get_row_idx_by_pn(ws, 5, item_pn)
                        
                        # 橫向：根據 掛帳人 ID 找欄 (搜尋 第 5 列)
                        target_col = get_col_idx_by_id(ws, 5, info['id'])
                        
                        if target_row and target_col:
                            # 填入數量，此操作會保留儲存格原有的格式
                            ws.cell(row=target_row, column=target_col, value=qty)
                            filled_count += 1
                        else:
                            # 顯示未匹配成功的警告
                            if not target_row:
                                st.warning(f"⚠️ {target_type} 模板找不到料號: {item_pn}")
                            if not target_col:
                                st.warning(f"⚠️ {target_type} 模板找不到工號: {info['id']} ({person})")

        # 5. 修改原始分頁名稱並輸出
        ws_orig = wb[target_sheet_name]
        ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
        
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue(), latest_date

    except Exception as e:
        st.error(f"執行出錯：{str(e)}")
        return None, None

# --- Streamlit 使用者介面 ---
uploaded_file = st.file_uploader("📂 請上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    if st.button("✨ 依照預設格式產生領用單"):
        with st.spinner("正在讀取模板並回填資料..."):
            processed_data, date = process_excel(uploaded_file)
            if processed_data:
                st.success(f"處理完成！已產出符合模板格式的分頁。")
                st.download_button(
                    label="📥 下載產出檔案",
                    data=processed_data,
                    file_name=f"領用單產出_{date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
