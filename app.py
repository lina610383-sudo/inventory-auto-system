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
    for row in range(1, ws.max_row + 1):
        val = ws.cell(row=row, column=pn_col_idx).value
        if val and str(val).strip().upper() == target_pn:
            return row
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
        for t in ['IEC', 'ICC']:
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_領用單_{latest_date}"
                output_ws_dict[t] = new_ws
            else:
                st.warning(f"⚠️ 檔案中缺少模板：『{tmpl_name}』")

        # 4. 定位與回填 (含料件資料同步)
        valid_person_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        missing_data = []
        filled_count = 0

        # 定義料件資訊欄位與模板對應欄位的映射 (範例：明細標題 -> 模板列索引)
        # 您可以根據實際 Excel 欄位調整這裡的數字
        item_info_mapping = {
            'Description': 2,   # 假設模板 B 欄是 Description
            'Supplier': 3,      # 假設模板 C 欄是 Supplier
            'Unit': 4,          # 假設模板 D 欄是 Unit
            'Unit Price': 6     # 假設模板 F 欄是 Unit Price
        }

        for _, row in detail_df.iterrows():
            item_pn = row.get('IEC PN')
            if pd.isna(item_pn): continue
            
            for person in valid_person_cols:
                qty = row[person]
                
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    person_name = str(person).strip()
                    info = payer_map[person_name]
                    target_type = info['type']
                    
                    if target_type in output_ws_dict:
                        ws = output_ws_dict[target_type]
                        
                        # 定位座標
                        target_row = get_row_idx_by_pn(ws, 5, item_pn) # PN 在 E 欄 (5)
                        target_col = get_col_idx_by_id(ws, 5, info['id']) # 工號在 第 5 列
                        
                        if target_row and target_col:
                            # 1. 回填數量
                            ws.cell(row=target_row, column=target_col, value=qty)
                            
                            # 2. 同步回填料件詳細資料 (從明細表填入模板對應列)
                            for detail_col_name, tmpl_col_idx in item_info_mapping.items():
                                if detail_col_name in row:
                                    ws.cell(row=target_row, column=tmpl_col_idx, value=row[detail_col_name])
                            
                            filled_count += 1
                        else:
                            reason = []
                            if not target_row: reason.append(f"料號 {item_pn} 不在模板 E 欄")
                            if not target_col: reason.append(f"工號 {info['id']} 不在模板第 5 列標題")
                            missing_data.append({
                                "類型": target_type, "領用人": person_name, "料號": item_pn, "原因": " & ".join(reason)
                            })

        # 5. 輸出
        ws_orig = wb[target_sheet_name]
        ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
        
        if missing_data:
            st.warning("📋 部分資料定位失敗，請檢查模板設定：")
            st.table(pd.DataFrame(missing_data))
        
        if filled_count > 0:
            st.success(f"✅ 成功填入 {filled_count} 筆領用資料及其詳細料件資訊。")

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
