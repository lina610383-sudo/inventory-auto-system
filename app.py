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

def process_excel(file):
    try:
        # 1. 載入原始活頁簿
        wb = openpyxl.load_workbook(file)
        sheet_names = wb.sheetnames
        
        # 尋找目標明細分頁 (未開單)
        pattern = r".*領用明細_(\d+).*\(未開單\)"
        matches = []
        for s in sheet_names:
            m = re.search(pattern, s)
            if m:
                matches.append((m.group(1), s))
        
        if not matches:
            st.error("❌ 找不到符合格式的分頁！請確認分頁名稱包含『領用明細_日期』且結尾為『(未開單)』")
            return None, None
        
        latest_date, target_sheet_name = sorted(matches, key=lambda x: x[0])[-1]
        st.info(f"📍 偵測到目標明細分頁：{target_sheet_name}")
        
        # 2. 讀取資料
        # header=1 假設明細標題在第 2 列
        detail_df = pd.read_excel(file, sheet_name=target_sheet_name, header=1)
        
        if "掛帳人清單" not in sheet_names:
            st.error("❌ 找不到『掛帳人清單』分頁！")
            return None, None
            
        payer_df = pd.read_excel(file, sheet_name="掛帳人清單")
        payer_df.iloc[:, 0] = payer_df.iloc[:, 0].ffill() 
        
        payer_map = {}
        for _, row in payer_df.iterrows():
            name = str(row['領用人']).strip()
            unit_type = str(row.iloc[0]).strip().upper() 
            if name and name != 'nan':
                payer_map[name] = {
                    'type': "IEC" if "IEC" in unit_type else "ICC",
                    'id': str(row['掛帳人']).strip()
                }

        # 3. 準備產出分頁
        output_ws_dict = {}
        current_row_dict = {} # 紀錄每個模板目前寫到哪一行
        for t in ['IEC', 'ICC']:
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_領用單_{latest_date}"
                output_ws_dict[t] = new_ws
                current_row_dict[t] = 6 # 假設模板從第 6 行開始填寫資料
            else:
                st.warning(f"⚠️ 檔案中缺少模板：『{tmpl_name}』")

        # 4. 定位與回填 (從未開單分頁抓取料件資訊並填入模板)
        valid_person_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        filled_count = 0

        # 定義料件資訊回填至模板的欄位索引 (1-based)
        # 您可以根據實際模板結構調整
        item_mapping = {
            'Description': 2,   # B 欄
            'Supplier': 3,      # C 欄
            'Unit': 4,          # D 欄
            'IEC PN': 5,        # E 欄 (料號)
            'Unit Price': 6     # F 欄
        }

        for _, row in detail_df.iterrows():
            item_pn = row.get('IEC PN')
            if pd.isna(item_pn): continue
            
            # 檢查這列中是否有任何 IEC 或 ICC 的領用需求
            has_qty_iec = False
            has_qty_icc = False
            
            # 先掃描一次這列資料，確認哪些單位需要開單
            for person in valid_person_cols:
                qty = row[person]
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    unit_type = payer_map[str(person).strip()]['type']
                    if unit_type == "IEC": has_qty_iec = True
                    if unit_type == "ICC": has_qty_icc = True

            # 針對需要的單位模板，填入料件基本資訊與數量
            for t in ['IEC', 'ICC']:
                if (t == "IEC" and has_qty_iec) or (t == "ICC" and has_qty_icc):
                    if t in output_ws_dict:
                        ws = output_ws_dict[t]
                        target_row = current_row_dict[t]
                        
                        # 1. 填入料件基本資料 (從未開單分頁抓取)
                        for col_name, col_idx in item_mapping.items():
                            if col_name in row:
                                ws.cell(row=target_row, column=col_idx, value=row[col_name])
                        
                        # 2. 橫向對位填入該人的領用數量
                        for person in valid_person_cols:
                            person_name = str(person).strip()
                            info = payer_map[person_name]
                            
                            if info['type'] == t:
                                qty = row[person]
                                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                                    # 搜尋工號在第 5 列的欄位座標
                                    target_col = get_col_idx_by_id(ws, 5, info['id'])
                                    if target_col:
                                        ws.cell(row=target_row, column=target_col, value=qty)
                                        filled_count += 1
                        
                        # 完成這列填寫後，模板行數下移一行
                        current_row_dict[t] += 1

        # 5. 輸出
        ws_orig = wb[target_sheet_name]
        ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
        
        if filled_count > 0:
            st.success(f"✅ 成功從未開單分頁提取料件資訊，並填入 {filled_count} 筆領用數量。")
        else:
            st.warning("⚠️ 掃描完成，但未發現有效的領用數量（需大於 0）。")

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue(), latest_date

    except Exception as e:
        st.error(f"❌ 錯誤：{str(e)}")
        return None, None

# 使用者介面
uploaded_file = st.file_uploader("📂 請上傳 Excel 檔案", type=["xlsx"])
if uploaded_file:
    if st.button("✨ 執行自動化回填"):
        processed_data, date = process_excel(uploaded_file)
        if processed_data:
            st.download_button("📥 下載結果", data=processed_data, file_name=f"領用單_{date}.xlsx")
