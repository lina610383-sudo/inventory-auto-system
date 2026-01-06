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
    在模板標題列搜尋工號，返回欄位索引 (1-based)
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
    # 從資料起始行 (第6行) 開始搜尋，這部分可依模板實際狀況調整
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=pn_col_idx).value
        if val and str(val).strip().upper() == target_pn:
            return row
    return None

def process_excel(file):
    try:
        # 1. 讀取 Excel 結構
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
            st.error("找不到符合格式的分頁！請確保分頁名稱包含『領用明細_日期』且結尾為『(未開單)』")
            return None, None
        
        # 排序以取得最新日期的分頁
        latest_date, target_sheet_name = sorted(matches, key=lambda x: x[0])[-1]
        st.info(f"📍 偵測到目標分頁：{target_sheet_name}")
        
        # 2. 讀取資料對照表
        # header=1 表示標題在 Excel 的第 2 列
        detail_df = pd.read_excel(file, sheet_name=target_sheet_name, header=1)
        payer_df = pd.read_excel(file, sheet_name="掛帳人清單")
        
        # 處理類別合併儲存格 (補全第一欄的 IEC/ICC)
        payer_df.iloc[:, 0] = payer_df.iloc[:, 0].ffill() 
        
        # 建立領用人對照地圖，嚴格區分類型
        payer_map = {}
        for _, row in payer_df.iterrows():
            name = str(row['領用人']).strip()
            # 取得該領用人屬於 IEC 還是 ICC
            unit_type = str(row.iloc[0]).strip().upper() 
            if name and name != 'nan':
                payer_map[name] = {
                    'type': "IEC" if "IEC" in unit_type else "ICC",
                    'id': str(row['掛帳人']).strip()
                }

        # 3. 根據明細中實際出現的單位類型準備產出模板
        output_ws_dict = {}
        for t in ['IEC', 'ICC']:
            # 這裡名稱必須與您的 Excel 分頁名稱完全一致
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_產出_{latest_date}"
                output_ws_dict[t] = new_ws
            else:
                st.warning(f"⚠️ 提示：在檔案中找不到模板分頁『{tmpl_name}』")

        # 4. 執行雙向對位填寫邏輯 (核心：按人名所屬類型分流)
        valid_person_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        filled_count = 0

        for _, row in detail_df.iterrows():
            # 獲取每一橫列的料號
            item_pn = row.get('IEC PN')
            if pd.isna(item_pn): 
                continue
            
            # 檢查每個人的領用數量
            for person in valid_person_cols:
                qty = row[person]
                
                # 只有數量大於 0 才進行填寫
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    person_name = str(person).strip()
                    info = payer_map[person_name]
                    target_type = info['type'] # 判斷為 IEC 或 ICC
                    
                    # 選取正確的模板分頁
                    if target_type in output_ws_dict:
                        ws = output_ws_dict[target_type]
                        
                        # 1. 縱向：根據料號 PN 尋找對應行 (搜尋模板的 E 欄/第 5 欄)
                        # 注意：如果您的料號在其他欄位，請修改下方的數字 5
                        target_row = get_row_idx_by_pn(ws, 5, item_pn)
                        
                        # 2. 橫向：根據工號 ID 尋找對應欄 (搜尋模板的第 5 列標題)
                        # 注意：如果您的工號標題在其他行，請修改下方的數字 5
                        target_col = get_col_idx_by_id(ws, 5, info['id'])
                        
                        if target_row and target_col:
                            # 精準回填交叉點
                            ws.cell(row=target_row, column=target_col, value=qty)
                            filled_count += 1
                        else:
                            # 若找不到坐標，則在介面顯示警告以利除錯
                            if not target_row:
                                st.warning(f"⚠️ 在 {target_type} 模板找不到料號: {item_pn}")
                            if not target_col:
                                st.warning(f"⚠️ 在 {target_type} 模板找不到工號: {info['id']} ({person_name})")

        if filled_count == 0:
            st.warning("比對完成，但沒有任何數據被填入模板，請檢查模板與明細的 PN/工號 是否完全一致。")

        # 5. 更新狀態並匯出
        ws_orig = wb[target_sheet_name]
        ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
        
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue(), latest_date

    except Exception as e:
        st.error(f"系統執行出錯：{str(e)}")
        return None, None

# --- Streamlit 介面 ---
uploaded_file = st.file_uploader("請上傳您的領用單 Excel 檔案", type=["xlsx"])

if uploaded_file:
    if st.button("✨ 執行自動對位填表"):
        with st.spinner("正在進行單位識別與雙向對位..."):
            processed_data, date = process_excel(uploaded_file)
            if processed_data:
                st.success(f"處理完成！已成功區分 IEC/ICC 並將資料填入對應模板。")
                st.download_button(
                    label="📥 下載自動產出檔案",
                    data=processed_data,
                    file_name=f"領用單產出結果_{date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
