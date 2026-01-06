import streamlit as st
import pandas as pd
import re
import io
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# 頁面配置
st.set_page_config(page_title="領用單自動化系統", layout="wide")

st.title("📦 領用單流程自動化系統")
st.info("上傳 Excel 後，系統會自動比對掛帳人並根據最新明細產出 IEC/ICC 領用單。")

def process_logic(file):
    try:
        # 載入活頁簿
        wb = load_workbook(file)
        all_sheets = wb.sheetnames

        # 1. 偵測最新日期分頁 (未開單)
        pattern = r"\(說明\) 領用明細_(\d+) \(未開單\)"
        matches = []
        for s in all_sheets:
            m = re.search(pattern, s)
            if m:
                matches.append((m.group(1), s))
        
        if not matches:
            st.error("找不到符合『(說明) 領用明細_日期 (未開單)』的分頁！")
            return None, None
            
        latest_date, target_sheet = sorted(matches, key=lambda x: x[0])[-1]
        st.success(f"📍 偵測到最新明細：{target_sheet}")

        # 2. 處理掛帳人清單
        if "掛帳人清單" not in all_sheets:
            st.error("找不到『掛帳人清單』分頁！")
            return None, None
        
        df_payers = pd.read_excel(file, sheet_name="掛帳人清單")
        df_payers.iloc[:, 0] = df_payers.iloc[:, 0].ffill() # 處理合併儲存格
        
        # 建立對照字典: 領用人 -> {工號, 單位, 類型}
        payer_map = {}
        for _, row in df_payers.iterrows():
            name = str(row['領用人']).strip()
            payer_map[name] = {
                'id': str(row['掛帳人']).strip(),
                'unit': str(row['單位']).strip(),
                'type': str(row.iloc[0]).strip()
            }

        # 3. 讀取明細內容 (標題在第 2 列)
        df_detail = pd.read_excel(file, sheet_name=target_sheet, header=1)
        
        # 4. 準備輸出模板
        output_ws = {}
        for fmt in ["IEC", "ICC"]:
            t_name = f"領用單格式範例 {fmt}"
            if t_name in all_sheets:
                ws = wb.copy_worksheet(wb[t_name])
                ws.title = f"{fmt}_Output_{latest_date}"
                output_ws[fmt] = ws

        # 5. 填寫邏輯
        person_cols = [c for c in df_detail.columns if str(c).strip() in payer_map]
        row_counters = {"IEC": 6, "ICC": 6} # 資料從第 6 行開始填

        for _, row in df_detail.iterrows():
            desc = row.get('Description')
            pn = row.get('IEC PN')
            
            # 防呆檢查
            final_desc = desc if pd.notna(desc) else "【須補資料】"
            final_pn = pn if pd.notna(pn) else "【須補資料】"

            for person in person_cols:
                qty = row[person]
                if pd.notna(qty) and qty > 0:
                    info = payer_map[str(person).strip()]
                    p_type = "IEC" if "IEC" in info['type'].upper() else "ICC"
                    
                    if p_type in output_ws:
                        ws = output_ws[p_type]
                        curr_r = row_counters[p_type]
                        
                        # 填入品項資訊 (依據範例格式)
                        ws.cell(row=curr_r, column=1, value=final_desc) # Description
                        ws.cell(row=curr_r, column=5, value=final_pn)   # IEC Part No / 開單料號
                        
                        # 尋找工號對應的欄位 (在第 5 列尋找)
                        target_col = None
                        for col_idx in range(1, ws.max_column + 1):
                            if str(ws.cell(row=5, column=col_idx).value).strip() == info['id']:
                                target_col = col_idx
                                break
                        
                        if target_col:
                            ws.cell(row=curr_r, column=target_col, value=qty)
                        
                        row_counters[p_type] += 1

        # 6. 更新明細分頁狀態
        wb[target_sheet].title = f"(說明) 領用明細_{latest_date} (已開單)"

        # 儲存結果
        out_bio = io.BytesIO()
        wb.save(out_bio)
        return out_bio.getvalue(), latest_date

    except Exception as e:
        st.error(f"系統執行錯誤: {e}")
        return None, None

# UI 介面
file_input = st.file_uploader("📂 請上傳領用單 Excel 檔案", type=["xlsx"])
if file_input:
    if st.button("🚀 產出已開單文件"):
        with st.spinner("正在進行比對與填表..."):
            res_data, date_val = process_logic(file_input)
            if res_data:
                st.success(f"完成！已產出日期 {date_val} 的領用單。")
                st.download_button(
                    label="📥 下載更新後的 Excel",
                    data=res_data,
                    file_name=f"領用單處理結果_{date_val}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
