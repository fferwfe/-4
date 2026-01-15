import streamlit as st
import pandas as pd
import io
import re
import json
import os
from google.cloud import vision
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side

# --- 1. 初始化 Google AI (從 Secrets 讀取) ---
def init_vision():
    if "gcp_service_account" in st.secrets:
        key_dict = dict(st.secrets["gcp_service_account"])
        with open("key.json", "w") as f:
            json.dump(key_dict, f)
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'key.json'
        return vision.ImageAnnotatorClient()
    return None

# --- 2. 核心辨識邏輯：優先抓內容，否則抓發言者 ---
def parse_line_screenshot(file, client):
    content = file.read()
    image = vision.Image(content=content)
    response = client.text_detection(image=image)
    
    # 這裡的邏輯會分析文字的座標位置
    # 簡單化處理：偵測每行文字，並判斷是否帶有 '+' 
    full_text = response.text_annotations[0].description if response.text_annotations else ""
    lines = full_text.split('\n')
    
    orders = []
    current_sender = "未知用戶"
    
    for line in lines:
        # 簡單過濾掉時間、結單等字眼
        if "前的" in line or "結單" in line: continue
        
        # 如果這行有 + 號
        if "+" in line:
            qty_match = re.search(r'\+(\d+)', line)
            qty = int(qty_match.group(1)) if qty_match else 1
            
            # 判斷內容是否有名字 (例如: 婷茹 +1)
            name_in_content = re.search(r'([^\+\s\d]+)\s*\+', line)
            
            if name_in_content:
                final_name = name_in_content.group(1).strip()
            else:
                # 如果內容沒名字，就使用上一次偵測到的「發言者姓名」
                final_name = current_sender
            
            orders.append({"姓名": final_name, "數量": qty})
        else:
            # 如果沒有 + 號，這行通常是發言者的名字（小字）
            if len(line.strip()) > 0 and len(line.strip()) < 10:
                current_sender = line.strip()
                
    return orders

# --- 3. 網頁介面 ---
st.set_page_config(page_title="學界二班團購系統", layout="wide")
st.title("🛒 團購截圖 AI 自動化對帳 (正式版)")

# 商品設定區
with st.expander("⚙️ 商品設定", expanded=True):
    df_config = pd.DataFrame([{"品名": "長榮航空米果", "單價": 150, "單位": "顆"}])
    edited_df = st.data_editor(df_config)
    item = edited_df.iloc[0]

uploaded_files = st.file_uploader("📸 上傳截圖", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files:
    client = init_vision()
    if client:
        all_orders = []
        for f in uploaded_files:
            all_orders.extend(parse_line_screenshot(f, client))
        
        if all_orders:
            st.write("📋 辨識清單：", pd.DataFrame(all_orders))

            if st.button("🚀 下載 2025 格式 Excel"):
                output = io.BytesIO()
                wb = Workbook()
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

                # --- Sheet 1: 付款單 (橫向) ---
                ws1 = wb.active
                ws1.title = "付款單"
                ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(all_orders))
                ws1['A1'] = f"學 界 二 班   {item['品名']}"
                ws1['A1'].alignment = Alignment(horizontal='center')
                
                for i, res in enumerate(all_orders, 1):
                    data_rows = [f"學二  {item['品名']}", "N1", res['姓名'], res['數量'], item['單位'], item['單價'], "元"]
                    for r_idx, val in enumerate(data_rows, 2):
                        cell = ws1.cell(row=r_idx, column=i, value=val)
                        cell.border = thin_border
                        cell.alignment = Alignment(horizontal='center')

                # --- Sheet 2: 對帳單 (縱向) ---
                ws2 = wb.create_sheet("對帳單")
                ws2['A1'] = f"學 界 二 班   {item['品名']}"
                headers = ["姓名", "數量", "應付款項", "付款狀態"]
                for c, h in enumerate(headers, 1):
                    ws2.cell(row=2, column=c, value=h).border = thin_border
                
                total_q = 0
                for r, res in enumerate(all_orders, 3):
                    ws2.cell(row=r, column=1, value=res['姓名']).border = thin_border
                    ws2.cell(row=r, column=2, value=res['數量']).border = thin_border
                    ws2.cell(row=r, column=3, value=res['數量']*item['單價']).border = thin_border
                    ws2.cell(row=r, column=4).border = thin_border
                    total_q += res['數量']
                
                ws2.cell(row=len(all_orders)+3, column=1, value="總計").border = thin_border
                ws2.cell(row=len(all_orders)+3, column=3, value=total_q*item['單價']).border = thin_border

                # --- Sheet 3: 商品標籤 ---
                ws3 = wb.create_sheet("商品標籤")
                for i, res in enumerate(all_orders):
                    base_r = i * 2 + 1
                    ws3.cell(row=base_r, column=1, value=f"學二{item['品名']}")
                    ws3.cell(row=base_r+1, column=1, value=res['姓名'])
                    ws3.cell(row=base_r+1, column=2, value=res['數量'])

                wb.save(output)
                st.download_button("💾 下載 Excel", output.getvalue(), f"{item['品名']}_對帳表.xlsx")
