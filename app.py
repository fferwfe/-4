import streamlit as st
import pandas as pd
import io
import re
import json
import os
from google.cloud import vision
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side

# --- 1. 初始化 Google AI (加強容錯) ---
def init_vision():
    if "gcp_service_account" in st.secrets:
        key_dict = dict(st.secrets["gcp_service_account"])
        
        # 💡 自動修正私鑰格式錯誤
        if "private_key" in key_dict:
            p_key = key_dict["private_key"]
            # 修正換行符號被轉義的問題
            p_key = p_key.replace("\\n", "\n")
            # 確保有正確的開頭與結尾
            if "-----BEGIN PRIVATE KEY-----" not in p_key:
                p_key = "-----BEGIN PRIVATE KEY-----\n" + p_key
            if "-----END PRIVATE KEY-----" not in p_key:
                p_key = p_key + "\n-----END PRIVATE KEY-----"
            key_dict["private_key"] = p_key

        with open("key.json", "w") as f:
            json.dump(key_dict, f)
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'key.json'
        
        try:
            return vision.ImageAnnotatorClient()
        except Exception as e:
            st.error(f"AI 啟動失敗，請檢查金鑰格式。錯誤訊息: {e}")
    return None

# --- 2. 智慧辨識：無名字則抓發言者 ---
def parse_line_screenshot(file, client):
    content = file.read()
    image = vision.Image(content=content)
    response = client.text_detection(image=image)
    if not response.text_annotations: return []

    texts = response.text_annotations
    # 依座標 y 軸排序，確保由上而下讀取
    blocks = []
    for text in texts[1:]:
        y = text.bounding_poly.vertices[0].y
        blocks.append({'text': text.description, 'y': y})
    blocks.sort(key=lambda x: x['y'])

    orders = []
    current_sender = "未知"
    for b in blocks:
        txt = b['text']
        if "前的" in txt or ":" in txt or "已結單" in txt: continue
        
        if "+" in txt:
            qty_match = re.search(r'\+(\d+)', txt)
            qty = int(qty_match.group(1)) if qty_match else 1
            # 嘗試找內容裡的名字 (例如: 珮真+1)
            name_match = re.search(r'^([^\+\s\d]+)\s*\+', txt)
            final_name = name_match.group(1) if name_match else current_sender
            orders.append({"姓名": final_name, "數量": qty})
        else:
            # 短文字通常是發言者姓名
            if 1 < len(txt) < 8: current_sender = txt
    return orders

# --- 3. 介面與 Excel 生成 ---
st.set_page_config(page_title="學界二班團購系統", layout="wide")
st.title("🛒 團購截圖 AI 自動對帳 (正式版)")

with st.expander("⚙️ 商品設定", expanded=True):
    df_config = pd.DataFrame([{"品名": "長榮航空米果", "單價": 150, "單位": "顆"}])
    edited_df = st.data_editor(df_config)
    item = edited_df.iloc[0]

uploaded_files = st.file_uploader("📸 上傳 LINE 截圖", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files:
    client = init_vision()
    if client:
        all_results = []
        for f in uploaded_files:
            all_results.extend(parse_line_screenshot(f, client))
        
        if all_results:
            st.success(f"✅ 辨識成功！共 {len(all_results)} 筆訂單")
            st.table(pd.DataFrame(all_results))

            # Excel 生成
            output = io.BytesIO()
            wb = Workbook()
            thin = Side(style='thin')
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            align = Alignment(horizontal='center', vertical='center')

            # 付款單 (橫向排列)
            ws1 = wb.active
            ws1.title = "付款單"
            ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(all_results))
            ws1['A1'] = f"學 界 二 班   {item['品名']}"
            ws1['A1'].alignment = align
            for i, res in enumerate(all_results, 1):
                rows = [f"學二  {item['品名']}", "N1", res['姓名'], res['數量'], item['單位'], item['單價'], "元"]
                for r_idx, val in enumerate(rows, 2):
                    c = ws1.cell(row=r_idx, column=i, value=val)
                    c.border = border
                    c.alignment = align

            # 對帳單 (縱向)
            ws2 = wb.create_sheet("對帳單")
            ws2['A1'] = f"學 界 二 班   {item['品名']}"
            header = ["姓名", "數量", "應付款項", "付款狀態"]
            for c, h in enumerate(header, 1): ws2.cell(row=2, column=c, value=h).border = border
            total = 0
            for r, res in enumerate(all_results, 3):
                ws2.cell(row=r, column=1, value=res['姓名']).border = border
                ws2.cell(row=r, column=2, value=res['數量']).border = border
                ws2.cell(row=r, column=3, value=res['數量']*item['單價']).border = border
                total += res['數量']*item['單價']
            ws2.cell(row=len(all_results)+3, column=1, value="總計").border = border
            ws2.cell(row=len(all_results)+3, column=3, value=total).border = border

            wb.save(output)
            st.download_button("🚀 下載 2025 航空米果格式 Excel", output.getvalue(), f"{item['品名']}_對帳表.xlsx")
