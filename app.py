import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
import io
import re

# 頁面配置
st.set_page_config(page_title="領用單自動化生成系統", layout="wide")
st.title("🚀 領用單流程自動化系統")

def get_col_idx_by_header(ws, header_row_idx, target_header_name):
    """
    動態偵測：在指定的標題列搜尋對應名稱的欄位索引 (1-based)
    支援多種可能標題的模糊匹配
    """
    if not target_header_name:
        return None
    
    # 定義常見的標題同義詞
    synonyms = {
        "Vendor": ["VENDOR", "SUPPLIER", "廠商"],
        "Description": ["DESCRIPTION", "品名", "描述"],
        "HP PN": ["HP PN", "HPPN", "HP料號"],
        "IEC PN": ["IEC PN", "IECPN", "IEC料號"],
        "Unit": ["UNIT", "單位"],
        "No": ["NO", "NO.", "項次", "序號"]
    }
    
    search_list = synonyms.get(target_header_name, [target_header_name])
    search_list = [s.upper() for s in search_list]

    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row_idx, column=col).value
        if val:
            cell_text = str(val).strip().upper()
            if any(s in cell_text for s in search_list):
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
        current_row_dict = {} 
        for t in ['IEC', 'ICC']:
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_領用單_{latest_date}"
                output_ws_dict[t] = new_ws
                current_row_dict[t] = 6 # 資料填寫起始行
            else:
                st.warning(f"⚠️ 檔案中缺少模板：『{tmpl_name}』")

        # 4. 定位與回填
        valid_person_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        filled_count = 0

        # 需要從明細中提取的關鍵欄位名稱
        fields_to_sync = ["No", "Vendor", "Description", "HP PN", "IEC PN", "Unit"]

        for index, row in detail_df.iterrows():
            item_pn = row.get('IEC PN')
            if pd.isna(item_pn): continue
            
            unit_targets = set()
            for person in valid_person_cols:
                qty = row[person]
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    unit_targets.add(payer_map[str(person).strip()]['type'])

            for t in unit_targets:
                if t in output_ws_dict:
                    ws = output_ws_dict[t]
                    target_row = current_row_dict[t]
                    
                    # 自動偵測模板欄位位置並填入資料
                    for field in fields_to_sync:
                        col_idx = get_col_idx_by_header(ws, 5, field)
                        if col_idx:
                            if field == "No":
                                ws.cell(row=target_row, column=col_idx, value=target_row - 5)
                            else:
                                # 處理明細表中可能不同名的欄位 (如 Vendor vs Supplier)
                                source_val = row.get(field)
                                if pd.isna(source_val) and field == "Vendor":
                                    source_val = row.get("Supplier")
                                
                                if pd.notna(source_val):
                                    ws.cell(row=target_row, column=col_idx, value=source_val)
                    
                    # 回填領用數量 (掛帳人對位)
                    for person in valid_person_cols:
                        person_name = str(person).strip()
                        info = payer_map[person_name]
                        
                        if info['type'] == t:
                            qty = row[person]
                            if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                                # 動態搜尋工號所在欄位
                                target_col = None
                                target_id = info['id']
                                for col in range(1, ws.max_column + 1):
                                    header_val = ws.cell(row=5, column=col).value
                                    if header_val and str(header_val).strip().upper() == str(target_id).upper():
                                        target_col = col
                                        break
                                
                                if target_col:
                                    ws.cell(row=target_row, column=target_col, value=qty)
                                    filled_count += 1
                    
                    current_row_dict[t] += 1

        # 5. 輸出
        ws_orig = wb[target_sheet_name]
        ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
        
        if filled_count > 0:
            st.success(f"✅ 已完成動態對位回填。已偵測並同步：Vendor, Description, HP PN, IEC PN, Unit。")
        else:
            st.warning("⚠️ 處理完成，但未發現有效的領用資料。")

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue(), latest_date

    except Exception as e:
        st.error(f"❌ 發生錯誤：{str(e)}")
        return None, None

# UI
uploaded_file = st.file_uploader("📂 請上傳領用單 Excel 檔案", type=["xlsx"])
if uploaded_file:
    if st.button("✨ 執行智慧動態生成"):
        processed_data, date = process_excel(uploaded_file)
        if processed_data:
            st.download_button("📥 下載領用單結果", data=processed_data, file_name=f"領用單產出_{date}.xlsx")
