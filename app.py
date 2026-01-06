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
        
        # 修改後的正則表達式：
        # .* 表示允許前方有任何文字（例如：(說明)、(緊急)）
        # \d+ 匹配日期數字
        # \(未開單\) 匹配結尾
        pattern = r".*領用明細_(\d+).*\(未開單\)"
        matches = []
        for s in sheet_names:
            m = re.search(pattern, s)
            if m:
                # 提取日期數字用於排序，並記錄完整分頁名稱
                matches.append((m.group(1), s))
        
        if not matches:
            st.error("找不到符合格式的分頁！請確保分頁名稱包含『領用明細_日期』且結尾為『(未開單)』")
            return None, None
        
        # 排序以取得最新日期的分頁
        latest_date, target_sheet_name = sorted(matches, key=lambda x: x[0])[-1]
        st.info(f"📍 偵測到目標分頁：{target_sheet_name}")
        
        # 2. 讀取資料對照表
        detail_df = pd.read_excel(file, sheet_name=target_sheet_name, header=1)
        payer_df = pd.read_excel(file, sheet_name="掛帳人清單")
        
        # 處理類別合併儲存格並清理資料
        payer_df.iloc[:, 0] = payer_df.iloc[:, 0].ffill() 
        
        # 建立領用人對照地圖
        payer_map = {}
        for _, row in payer_df.iterrows():
            name = str(row['領用人']).strip()
            if name and name != 'nan':
                payer_map[name] = {
                    'type': str(row.iloc[0]).strip().upper(),
                    'id': str(row['掛帳人']).strip()
                }

        # 3. 複製並準備模板分頁
        output_ws_dict = {}
        for t in ['IEC', 'ICC']:
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_產出_{latest_date}"
                output_ws_dict[t] = new_ws
            else:
                st.warning(f"缺少範例模板：{tmpl_name}")

        # 4. 執行雙向對位填寫邏輯
        # 找出明細表中存在於對照表的人名欄位
        valid_person_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        
        for _, row in detail_df.iterrows():
            item_pn = row.get('IEC PN')
            if pd.isna(item_pn): 
                continue
            
            for person in valid_person_cols:
                qty = row[person]
                # 僅處理有數量的項目
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    info = payer_map[person.strip()]
                    p_type = "IEC" if "IEC" in info['type'] else "ICC"
                    
                    if p_type in output_ws_dict:
                        ws = output_ws_dict[p_type]
                        
                        # 縱向：在 E 欄 (第5欄) 尋找對應料號
                        target_row = get_row_idx_by_pn(ws, 5, item_pn)
                        # 橫向：在第 5 列尋找對應工號
                        target_col = get_col_idx_by_id(ws, 5, info['id'])
                        
                        if target_row and target_col:
                            ws.cell(row=target_row, column=target_col, value=qty)
                        else:
                            # 輸出除錯資訊
                            if not target_row:
                                st.warning(f"⚠️ 在 {p_type} 模板中找不到料號: {item_pn}")
                            if not target_col:
                                st.warning(f"⚠️ 在 {p_type} 模板中找不到工號: {info['id']} ({person})")

        # 5. 更新原始分頁狀態並儲存
        ws_orig = wb[target_sheet_name]
        ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
        
        # 寫入二進位流
        output = io.BytesIO()
        wb.save(output)
        return output.getvalue(), latest_date

    except Exception as e:
        st.error(f"系統執行出錯：{str(e)}")
        return None, None

# --- Streamlit 介面渲染 ---
uploaded_file = st.file_uploader("請上傳您的領用單 Excel 檔案", type=["xlsx"])

if uploaded_file:
    if st.button("✨ 執行自動對位填表"):
        with st.spinner("正在比對料號與工號座標..."):
            processed_data, date = process_excel(uploaded_file)
            if processed_data:
                st.success(f"處理完成！日期 {date} 的檔案已準備好下載。")
                st.download_button(
                    label="📥 下載自動產出檔案",
                    data=processed_data,
                    file_name=f"領用單產出結果_{date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
