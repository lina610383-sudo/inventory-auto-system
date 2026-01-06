import streamlit as st
import pandas as pd
import re
import io
from openpyxl import load_workbook

# 設置頁面標題與寬度
st.set_page_config(page_title="領用單自動化系統", layout="wide")

st.title("🚀 領用單流程自動化系統")
st.markdown("---")
st.info("請將您的 Excel 檔案（如：領用單流程優化.xlsx）上傳至下方。")

def get_col_idx_by_id(ws, row_idx, target_id):
    """
    在 Excel 指定列中搜尋工號的欄位索引 (1-based)
    """
    if not target_id:
        return None
    search_val = str(target_id).strip().upper()
    for col in range(1, ws.max_column + 1):
        cell_val = ws.cell(row=row_idx, column=col).value
        if cell_val and str(cell_val).strip().upper() == search_val:
            return col
    return None

def process_logic(uploaded_file):
    try:
        # 1. 載入原始活頁簿
        wb = load_workbook(uploaded_file)
        all_sheets = wb.sheetnames

        # 2. 定位最新日期的明細分頁 (支援 "(說明) 領用明細_XXXX" 格式)
        pattern = r"領用明細_(\d+)"
        matches = []
        for s in all_sheets:
            m = re.search(pattern, s)
            if m:
                matches.append((m.group(1), s))
        
        if not matches:
            st.error("❌ 找不到符合格式『領用明細_日期』的分頁！請確認分頁名稱。")
            return None, None
            
        latest_date, target_sheet = sorted(matches, key=lambda x: x[0])[-1]
        st.success(f"📍 已鎖定最新明細分頁：`{target_sheet}`")

        # 3. 處理「掛帳人清單」
        if "掛帳人清單" not in all_sheets:
            st.error("❌ 找不到『掛帳人清單』分頁！")
            return None, None
        
        df_payers = pd.read_excel(uploaded_file, sheet_name="掛帳人清單")
        # 處理 A 欄 (IEC/ICC 類別) 的合併儲存格
        df_payers.iloc[:, 0] = df_payers.iloc[:, 0].ffill()
        
        # 建立對照字典: { 姓名: {工號, 類型} }
        payer_map = {}
        for _, row in df_payers.iterrows():
            name = str(row['領用人']).strip()
            if name and name != 'nan':
                payer_map[name] = {
                    'id': str(row['掛帳人']).strip(),
                    'type': str(row.iloc[0]).strip().upper()
                }

        # 4. 讀取明細資料 (標題在第 2 列，故 header=1)
        df_detail = pd.read_excel(uploaded_file, sheet_name=target_sheet, header=1)
        
        # 5. 準備輸出模板 (複製範例格式並重新命名)
        output_ws = {}
        for fmt in ["IEC", "ICC"]:
            template_name = f"領用單格式範例 {fmt}"
            if template_name in all_sheets:
                ws = wb.copy_worksheet(wb[template_name])
                ws.title = f"{fmt}_領用單_{latest_date}"
                output_ws[fmt] = ws
            else:
                st.warning(f"⚠️ 提示：缺少模板分頁：『{template_name}』")

        # 6. 核心填寫邏輯
        # 找出明細表中屬於領用人姓名的欄位
        person_cols = [c for c in df_detail.columns if str(c).strip() in payer_map]
        
        # 定義資料填入的起始行 (根據您的範例從第 6 行開始)
        row_counters = {"IEC": 6, "ICC": 6}

        for _, row in df_detail.iterrows():
            desc = row.get('Description')
            pn = row.get('IEC PN')
            
            # 填入內容防呆
            final_desc = str(desc) if pd.notna(desc) else "【無描述】"
            final_pn = str(pn) if pd.notna(pn) else "【無料號】"

            # 遍歷這一列中所有領用人的領用量
            for person in person_cols:
                qty = row[person]
                # 只有數量大於 0 才進行處理
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    info = payer_map[str(person).strip()]
                    # 判斷是歸類在 IEC 還是 ICC
                    p_type = "IEC" if "IEC" in info['type'] else "ICC"
                    
                    if p_type in output_ws:
                        ws = output_ws[p_type]
                        curr_r = row_counters[p_type]
                        
                        # A. 填入品項基本資訊 (Column 1=Description, Column 5=Part No)
                        ws.cell(row=curr_r, column=1, value=final_desc)
                        ws.cell(row=curr_r, column=5, value=final_pn)
                        
                        # B. 根據「工號」尋找模板第 5 列標題中對應的欄位索引
                        target_col = get_col_idx_by_id(ws, 5, info['id'])
                        
                        if target_col:
                            # C. 填入領用數量
                            ws.cell(row=curr_r, column=target_col, value=qty)
                        else:
                            # 若模板沒這工號，自動新增至最後一欄
                            new_col = ws.max_column + 1
                            ws.cell(row=5, column=new_col, value=info['id'])
                            ws.cell(row=curr_r, column=new_col, value=qty)
                        
                        # 完成一行填寫，計數器遞增
                        row_counters[p_type] += 1

        # 7. 將原明細分頁狀態標記為 (已開單)
        if "(未開單)" in target_sheet:
            wb[target_sheet].title = target_sheet.replace("(未開單)", "(已開單)")

        # 將活頁簿儲存至記憶體
        out_bio = io.BytesIO()
        wb.save(out_bio)
        return out_bio.getvalue(), latest_date

    except Exception as e:
        st.error(f"❌ 執行發生錯誤: {e}")
        return None, None

# --- Streamlit 網頁介面 ---
file_input = st.file_uploader("📂 請上傳領用單 Excel 檔案", type=["xlsx"])

if file_input:
    if st.button("🚀 開始產出已開單領用單文件"):
        with st.spinner("系統正在進行自動比對與填表作業..."):
            res_data, date_val = process_logic(file_input)
            if res_data:
                st.success(f"✅ 處理完成！已成功分析日期 {date_val} 的數據。")
                st.download_button(
                    label="📥 下載處理結果 Excel",
                    data=res_data,
                    file_name=f"領用單結果_{date_val}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
