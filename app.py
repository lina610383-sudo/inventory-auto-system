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
    # 遍歷所有行以尋找料號座標
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
        
        # 尋找目標明細分頁 (彈性匹配：領用明細_日期...未開單)
        pattern = r".*領用明細_(\d+).*\(未開單\)"
        matches = []
        for s in sheet_names:
            m = re.search(pattern, s)
            if m:
                matches.append((m.group(1), s))
        
        if not matches:
            st.error("❌ 找不到符合格式的分頁！請確認分頁名稱包含『領用明細_日期』且結尾為『(未開單)』")
            return None, None
        
        # 取得最新日期的分頁
        latest_date, target_sheet_name = sorted(matches, key=lambda x: x[0])[-1]
        st.info(f"📍 偵測到目標明細分頁：{target_sheet_name}")
        
        # 2. 讀取明細資料與掛帳人資訊
        # 假設明細標題在第 2 列 (Pandas header=1)
        detail_df = pd.read_excel(file, sheet_name=target_sheet_name, header=1)
        
        if "掛帳人清單" not in sheet_names:
            st.error("❌ 找不到『掛帳人清單』分頁！")
            return None, None
            
        payer_df = pd.read_excel(file, sheet_name="掛帳人清單")
        # 處理合併儲存格：補全第一欄的單位類型 (IEC/ICC)
        payer_df.iloc[:, 0] = payer_df.iloc[:, 0].ffill() 
        
        # 建立地圖：領用人 -> { 單位類型, 掛帳人工號 }
        payer_map = {}
        for _, row in payer_df.iterrows():
            name = str(row['領用人']).strip()
            unit_type = str(row.iloc[0]).strip().upper() 
            if name and name != 'nan':
                payer_map[name] = {
                    'type': "IEC" if "IEC" in unit_type else "ICC",
                    'id': str(row['掛帳人']).strip()
                }

        # 3. 準備產出分頁 (根據流程複製模板)
        output_ws_dict = {}
        for t in ['IEC', 'ICC']:
            tmpl_name = f"領用單格式範例 {t}"
            if tmpl_name in sheet_names:
                # 直接複製預設格式，保留框線、標題與公式
                new_ws = wb.copy_worksheet(wb[tmpl_name])
                new_ws.title = f"{t}_領用單_{latest_date}"
                output_ws_dict[t] = new_ws
            else:
                st.warning(f"⚠️ 檔案中缺少模板：『{tmpl_name}』，將無法產出此類別。")

        # 4. 資料比對與填入 (雙向對位)
        # 找出明細中符合「領用人」定義的欄位
        valid_person_cols = [c for c in detail_df.columns if str(c).strip() in payer_map]
        
        # 用於記錄缺漏資料
        missing_data = []
        filled_count = 0

        for _, row in detail_df.iterrows():
            item_pn = row.get('IEC PN')
            item_desc = row.get('Description', 'Unknown')
            
            if pd.isna(item_pn): continue
            
            for person in valid_person_cols:
                qty = row[person]
                
                # 只有當領用數量大於 0 才處理
                if pd.notna(qty) and isinstance(qty, (int, float)) and qty > 0:
                    person_name = str(person).strip()
                    info = payer_map[person_name]
                    target_type = info['type']
                    
                    if target_type in output_ws_dict:
                        ws = output_ws_dict[target_type]
                        
                        # A. 縱向定位：在 E 欄 (第 5 欄) 找料號 PN
                        target_row = get_row_idx_by_pn(ws, 5, item_pn)
                        # B. 橫向定位：在 第 5 列 找掛帳人工號
                        target_col = get_col_idx_by_id(ws, 5, info['id'])
                        
                        if target_row and target_col:
                            # 填入數量，保留原格式
                            ws.cell(row=target_row, column=target_col, value=qty)
                            filled_count += 1
                        else:
                            # 收集缺失座標的資料
                            reason = []
                            if not target_row: reason.append(f"料號 {item_pn} 不在 E 欄")
                            if not target_col: reason.append(f"工號 {info['id']} 不在第 5 列")
                            missing_data.append({
                                "類型": target_type,
                                "領用人": person_name,
                                "品名": item_desc,
                                "料號": item_pn,
                                "工號": info['id'],
                                "原因": " & ".join(reason)
                            })

        # 5. 完成處理：標示狀態並輸出
        # 將原始明細分頁更名為 (已開單)
        ws_orig = wb[target_sheet_name]
        ws_orig.title = target_sheet_name.replace("(未開單)", "(已開單)")
        
        # 顯示處理結果與缺漏報告
        if missing_data:
            st.warning("📋 部分資料因座標不匹配無法填入，請參考下方清單：")
            st.table(pd.DataFrame(missing_data))
        
        if filled_count > 0:
            st.success(f"✅ 成功填入 {filled_count} 筆資料至模板中。")
        else:
            st.error("❌ 未能在模板中找到對應的座標，請檢查料號欄(E)與工號列(5)。")

        output = io.BytesIO()
        wb.save(output)
        return output.getvalue(), latest_date

    except Exception as e:
        st.error(f"❌ 處理過程中發生非預期錯誤：{str(e)}")
        return None, None

# --- 使用者介面渲染 ---
uploaded_file = st.file_uploader("📂 請上傳包含『領用明細』與『掛帳人清單』的 Excel 檔案", type=["xlsx"])

if uploaded_file:
    if st.button("✨ 啟動自動化領用單生成"):
        with st.spinner("正在進行單位識別、格式複製與精準填表..."):
            processed_data, date = process_excel(uploaded_file)
            if processed_data:
                st.download_button(
                    label="📥 下載已開單之領用單結果",
                    data=processed_data,
                    file_name=f"領用單產出結果_{date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
