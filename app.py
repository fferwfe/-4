import streamlit as st
import pandas as pd
import io
import re
import json
import os
from google.cloud import vision
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side

# --- 1. 初始化 Google AI ---
def init_vision():
    if "gcp_service_account" in st.secrets:
        key_dict = dict(st.secrets["gcp_service_account"])
        # 在伺服器端建立臨時金鑰檔
        with open("key.json", "w") as f:
            json.dump(key_dict, f)
        os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = 'key.json'
        return vision.ImageAnnotatorClient()
    return None

# --- 2. 智慧型辨識邏輯：內容優先，發言者補位 ---
def get_orders_from_ai(uploaded_file, client):
    content = uploaded_file.read()
    image = vision.Image(content=content)
    response = client.text_detection(image=image)
    
    if not response.text_annotations:
        return []

    texts = response.text_annotations
    # texts[0] 是整張圖的文字，後面的是個別區塊
    # 我們需要根據座標 y 軸來判斷誰在誰上面
    blocks = []
    for text in texts[1:]:
        vertices = text.bounding_poly.vertices
        y_top = vertices[0].y
        blocks.append({'text': text.description, 'y': y_top})
    
    # 依照 y 軸排序（從上到下）
    blocks.sort(key=lambda x: x['y'])
    
    orders = []
    last_potential_sender = "未知"
    
    for b in blocks:
        txt = b['text']
        # 排除時間與系統字
        if "前的" in txt or "結單" in txt or ":" in txt: continue
        
        if "+" in txt:
            qty_match = re.search(r'\+(\d+)', txt)
            qty = int(qty_match.group(1)) if qty_match else 1
            
            # 判斷內容是否有名字 (例如: 婷茹+1)
            name_in_msg = re.match(r'^([^\+\s\d]+)', txt)
            if name_in_msg and len(name_in_msg.group(1)) > 1:
                final_name = name_in_msg.group(1)
            else:
                final_name = last_potential_sender # 抓取上方發言者
            
            orders.append({"姓名": final_name, "數量": qty})
        else:
            # 這可能是一個發言者的名字
            if 1 < len(txt) < 10:
                last_potential_sender = txt
                
    return orders

# --- 3. 網頁介面 ---
st.set_page_config(page_title="學界二班團購系統", layout="wide")
st.title("🛒 團購截圖 AI 自動化對帳系統")

with st.expander("⚙️ 商品設定", expanded=True):
    df_config = pd.DataFrame([{"品名": "長榮航空米果", "單價": 150, "單位": "顆"}])
    edited_df = st.data_editor(df_config)
    item = edited_df.iloc[0]

uploaded_files = st.file_uploader("📸 上傳 LINE 截圖 (可多張)", type=['png', 'jpg', 'jpeg'], accept_multiple_files=True)

if uploaded_files:
    client = init_vision()
    if client:
        all_results = []
        for f in uploaded_files:
            all_results.extend(get_orders_from_ai(f, client))
        
        if all_results:
            st.success(f"✅ 辨識完成！共抓取 {len(all_results)} 筆訂單")
            st.table(pd.DataFrame(all_results))

            # --- 4. 生成 Excel 邏輯 ---
            # 使用 BytesIO 緩存 Excel 內容
            output = io.BytesIO()
            wb = Workbook()
            thin = Side(style='thin')
            border = Border(left=thin, right=thin, top=thin, bottom=thin)
            align = Alignment(horizontal='center', vertical='center')

            # --- 付款單 (橫向) ---
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

            # --- 對帳單 ---
            ws2 = wb.create_sheet("對帳單")
            ws2['A1'] = f"學 界 二 班   {item['品名']}"
            headers = ["姓名", "數量", "應付款項", "付款狀態"]
            for c_idx, h in enumerate(headers, 1):
                ws2.cell(row=2, column=c_idx, value=h).border = border
            
            total_sum = 0
            for r_idx, res in enumerate(all_results, 3):
                ws2.cell(row=r_idx, column=1, value=res['姓名']).border = border
                ws2.cell(row=r_idx, column=2, value=res['數量']).border = border
                amt = res['數量'] * item['單價']
                ws2.cell(row=r_idx, column=3, value=amt).border = border
                ws2.cell(row=r_idx, column=4).border = border
                total_sum += amt
            
            last_row = len(all_results) + 3
            ws2.cell(row=last_row, column=1, value="總計").border = border
            ws2.cell(row=last_row, column=3, value=total_sum).border = border

            # --- 商品標籤 ---
            ws3 = wb.create_sheet("商品標籤")
            for idx, res in enumerate(all_results):
                r = idx * 2 + 1
                ws3.cell(row=r, column=1, value=f"學二{item['品名']}")
                ws3.cell(row=r+1, column=1, value=res['姓名'])
                ws3.cell(row=r+1, column=2, value=res['數量'])

            wb.save(output)
            
            st.download_button(
                label="🚀 下載正式 Excel 表格",
                data=output.getvalue(),
                file_name=f"{item['品名']}_自動對帳表.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ 請檢查 Streamlit Secrets 是否已填入 Google 金鑰內容。")
