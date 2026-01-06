import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import io
import re

# 頁面配置
st.set_page_config(page_title="領用單自動化生成系統", layout="wide")
st.title("🚀 領用單流程自動化系統 (雙向對位版)")

def get_col_idx_by_id(ws, header_row_idx, target_id):
    """在模板標題列搜尋工號，返回欄位索引 (1-based)"""
    if not target_id: return None
    target_id = str(target_id).strip().upper()
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row_idx, column=col).value
        if val and str(val).strip().upper() == target_id:
            return col
    return None

def get_row_idx_by_pn(ws, pn_col_idx, target_pn):
    """在模板料號欄搜尋 PN，返回行索引 (1-based)"""
    if not target_pn: return None
    target_pn = str(target_pn).strip().upper()
    # 從資料起始行 (第6行) 開始搜尋
    for row in range(6, ws.max_row + 1):
        val = ws.cell(row=row, column=pn_col_idx).value
        if val and str(val).strip().upper() == target_pn:
            return row
    return None

def process_excel(file):
    try:
        # 1. 讀取 Excel 結構
        wb = openpyxl.load_workbook(file)
        sheet_names = wb.sheetnames
        
        # 尋找最新的 "未開單" 分頁
        pattern = r"\(說明\) 領用明細_(\d+) \(未開單\)"
        matches = [(re.search(pattern, s).group(1), s) for s in sheet_names if re.search(pattern, s)]
        
        if not matches:
            st.error("找不到符合格式的『(說明) 領用明細_日期 (未開單)』分頁！")
            return None, None
        
        latest_date, target_sheet_name = sorted(matches, key=lambda x: x[0])[-1]
        st.info(f"正在處理明細分頁：{target_sheet_name}")
        
        # 2. 讀取資料對照表
        detail_df = pd.read_excel(file, sheet_name=target_sheet_name, header=1)
        payer_df = pd.read_excel(file, sheet_name="掛帳人清單")
        payer_df.iloc[:, 0] = payer_df.iloc[:, 0].ffill() # 處理類別合併儲存格
        
        # 建立人名對照表
        payer_map = {}
        for _, row in payer_df.iterrows():
            name = str(row['領用人']).strip()
            if name and name != 'nan':
                payer_map[name] = {
                    'type': str(row.iloc[0]).strip().upper(),
                    'id': str(row['掛帳人']).strip()
                }

        # 3. 準備模板分頁
        output_ws = {}
        for t in ['IEC', 'ICC']:
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_產出_{latest_date}"
                output_ws[t] = new_ws
            else:
                st.warning(f"缺少範例模板：{tmpl_name}")

        # 4. 執行雙向對位填寫
        name_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        
        for _, row in detail_df.iterrows():
            item_pn = row.get('IEC PN')
            if pd.isna(item_pn): continue
            
            for person in name_cols:
                qty = row[person]
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    info = payer_map[person.strip()]
                    p_type = "IEC" if "IEC" in info['type'] else "ICC"
                    
                    if p_type in output_ws:
                        ws = output_ws[p_type]
                        
                        # 縱向定位：在 E 欄 (第5欄) 找料號
                        target_row = get_row_idx_by_pn(ws, 5, item_pn)
                        # 橫向定位：在第 5 列找工號
                        target_col = get_col_idx_by_id(ws, 5, info['id'])
                        
                        if target_row and target_col:
                            ws.cell(row=target_row, column=target_col, value=qty)
                        else:
                            if not target_row:
                                st.warning(f"⚠️ 模板 {p_type} 找不到料號: {item_pn}")
                            if not target_col:
                                st.warning(f"⚠️ 模板 {p_type} 找不到工號: {info['id']} ({person})")

        # 5. 更新原分頁狀態
        ws_orig = wb[target_sheet_name]
        ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
        
        # 存檔
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue(), latest_date

    except Exception as e:
        st.error(f"執行出錯：{e}")
        return None, None

# --- UI 介面 ---
uploaded_file = st.file_uploader("請上傳您的 Excel 檔案", type=["xlsx"])

if uploaded_file:
    if st.button("✨ 點我自動產生領用清單"):
        with st.spinner("正在進行料號與工號雙向定位..."):
            processed_data, date = process_excel(uploaded_file)
            if processed_data:
                st.success(f"處理完畢！日期 {date} 的檔案已根據模板填寫完成。")
                st.download_button(
                    label="📥 下載產出檔案",
                    data=processed_data,
                    file_name=f"領用單產出_{date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
