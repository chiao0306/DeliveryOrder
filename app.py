import streamlit as st
import os
import json
import base64
import requests
import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Gemini OCR 測試", layout="centered")

# ── 側邊欄：Gemini 模型選擇 ──────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 設定")
    gemini_model = st.selectbox(
        "Gemini 模型",
        options=[
            "gemini-3.5-flash",
            "gemini-3.1-flash-lite",
        ],
        index=1,
        format_func=lambda x: {
            "gemini-3.5-flash": "Gemini 3.5 Flash",
            "gemini-3.1-flash-lite": "Gemini 3.1 Flash Lite",
        }[x],
    )
    st.caption(f"目前選用：`{gemini_model}`")

st.title("👁️ 軋輥組裝報表 OCR 分析")

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

# ── 定義產生 Excel 的函數 ──────────────────────────────────────
def create_excel_report(parsed_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "軋輥組裝報表"

    # 樣式設定
    font_header = Font(name="微軟正黑體", size=11, bold=True, color="FFFFFF")
    font_body = Font(name="微軟正黑體", size=10)
    font_bold = Font(name="微軟正黑體", size=10, bold=True)
    fill_header = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    align_center = Alignment(horizontal="center", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    border_thin = Border(
        left=Side(style='thin', color='A6A6A6'), right=Side(style='thin', color='A6A6A6'),
        top=Side(style='thin', color='A6A6A6'), bottom=Side(style='thin', color='A6A6A6')
    )

    # 第一列：標題
    ws.cell(row=1, column=1, value="施工項目")
    ws.cell(row=1, column=2, value="型號")
    ws.cell(row=1, column=3, value="編號尺寸")
    ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=16)

    for col in range(1, 17):
        cell = ws.cell(row=1, column=col)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_thin

    # 指定的輸出順序與 JSON key 的對應 (Excel顯示名稱, JSON鍵名)
    categories_order = [
        ("再生", "再生"),
        ("軸頸再生", "軸位再生"),
        ("粗車", "粗車"),
        ("軸頸粗車", "軸位粗車"),
        ("精車", "精車"),
        ("軸頸精車", "軸位精車")
    ]

    current_row = 2

    for display_cat, json_cat in categories_order:
        # 如果該分類沒資料就跳過
        if json_cat not in parsed_data or not parsed_data[json_cat]:
            continue
            
        model_dict = parsed_data[json_cat]
        is_first_cat_row = True
        
        for model, rollers in model_dict.items():
            if not rollers:
                continue
                
            items = list(rollers.items())
            
            # 每 7 組一列 (佔據 14 欄)
            for chunk_idx in range(0, len(items), 7):
                chunk = items[chunk_idx:chunk_idx+7]
                
                # 第一欄 (施工項目)
                cell_a = ws.cell(row=current_row, column=1)
                if is_first_cat_row:
                    cell_a.value = display_cat
                    cell_a.font = font_bold
                    is_first_cat_row = False
                cell_a.alignment = align_center
                cell_a.border = border_thin
                
                # 第二欄 (型號)
                cell_b = ws.cell(row=current_row, column=2)
                if chunk_idx == 0:
                    cell_b.value = model
                    cell_b.font = font_bold
                cell_b.alignment = align_center
                cell_b.border = border_thin
                
                # 第三欄開始 (編號尺寸對)
                for pair_idx, (r_id, r_val) in enumerate(chunk):
                    col_id = 3 + pair_idx * 2
                    col_val = 4 + pair_idx * 2
                    
                    cell_id = ws.cell(row=current_row, column=col_id, value=r_id)
                    
                    # 判斷小數或整數以設定對齊與格式
                    try:
                        val_num = float(r_val) if '.' in r_val else int(r_val)
                    except ValueError:
                        val_num = r_val
                        
                    cell_val = ws.cell(row=current_row, column=col_val, value=val_num)
                    
                    cell_id.font = font_body
                    cell_id.alignment = align_center
                    cell_id.border = border_thin
                    
                    cell_val.font = font_body
                    cell_val.alignment = align_right if isinstance(val_num, (int, float)) else align_center
                    cell_val.border = border_thin
                    
                # 填補剩下的空白框線，保持表格完整性
                for remaining in range(len(chunk) * 2, 14):
                    col_empty = 3 + remaining
                    cell_empty = ws.cell(row=current_row, column=col_empty, value="")
                    cell_empty.border = border_thin
                    
                current_row += 1

    # 自動調整欄寬
    for col in ws.columns:
        max_len = 0
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 2, 10)

    # 存入記憶體
    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()


uploaded_file = st.file_uploader(
    "上傳報表圖檔（如：五號機軋輥組裝報表）", type=["jpg", "jpeg", "png", "pdf"]
)

if uploaded_file:
    file_content = uploaded_file.getvalue()

    st.subheader("🟢 Gemini 結構化辨識")
    if not GEMINI_API_KEY:
        st.error("⚠️ 請設定 GEMINI_API_KEY 環境變數。")
    else:
        PROMPT = """
你是工廠軋輥維修報表的資料擷取助手。
請分析這張「軋輥組裝報表」圖片，依照以下規則輸出 JSON：

規則：
1. 掃描每一筆軋輥記錄（每一列）。注意報表中有區段標示不同的「輥輪型號」（例如 30D, 30S, 30L）。
2. 欄位對應如下：
   - 粗車（Roll 粗車尺寸，非軸位）
   - 再生（Roll 再生尺寸）
   - 精車（Roll 精車尺寸）
   - 軸位粗車（軸位粗車尺寸）
   - 軸位再生（軸位再生尺寸）
   - 軸位精車（軸位精車尺寸）
3. 若該格為數字（含小數），代表有施做，請記錄該數字（字串格式）。
4. 若該格為「X」或空白，代表未施做，請略過（不要包含在輸出中）。
5. 每個類別只輸出「有施做」的項目。
6. JSON 結構必須為三層：【施工類別】 -> 【輥輪型號】 -> 【輥輪編號: 尺寸】。

輸出格式範例（只輸出 JSON，不要加任何說明文字）：
{
  "粗車": {
    "30D": { "N30DL90": "288", "30DL13": "288" },
    "30S": { "V30S58": "287", "M30S141": "288" }
  },
  "再生": {
    "30D": { "N30DL90": "307" }
  },
  "精車": {
    "30D": { "N30DL90": "300.06" }
  },
  "軸位粗車": {
    "30S": { "V30S58": "146" }
  },
  "軸位再生": {
    "30S": { "V30S58": "155" },
    "30L": { "M30L237": "155" }
  },
  "軸位精車": {
    "30S": { "V30S58": "149.97" }
  }
}

請直接輸出 JSON，不要加 markdown 代碼區塊。
""".strip()

        with st.spinner(f"使用 {gemini_model} 辨識中…"):
            try:
                fname = uploaded_file.name.lower()
                if fname.endswith(".pdf"):
                    mime = "application/pdf"
                elif fname.endswith(".png"):
                    mime = "image/png"
                else:
                    mime = "image/jpeg"

                b64_data = base64.b64encode(file_content).decode("utf-8")

                payload = {
                    "contents": [
                        {
                            "parts": [
                                {
                                    "inline_data": {
                                        "mime_type": mime,
                                        "data": b64_data,
                                    }
                                },
                                {"text": PROMPT},
                            ]
                        }
                    ]
                }

                api_url = (
                    f"https://generativelanguage.googleapis.com/v1beta/models/"
                    f"{gemini_model}:generateContent?key={GEMINI_API_KEY}"
                )
                resp = requests.post(
                    api_url,
                    headers={"Content-Type": "application/json"},
                    json=payload,
                    timeout=60,
                )
                resp.raise_for_status()
                raw_text = (
                    resp.json()
                    .get("candidates", [{}])[0]
                    .get("content", {})
                    .get("parts", [{}])[0]
                    .get("text", "")
                    .strip()
                )

                if raw_text.startswith("```"):
                    raw_text = raw_text.split("```")[1]
                    if raw_text.startswith("json"):
                        raw_text = raw_text[4:]
                    raw_text = raw_text.strip()

                parsed = json.loads(raw_text)
                st.success("✅ Gemini 辨識成功！")

                # ── 新增：匯出 Excel 下載按鈕 ──────────────────────
                excel_data = create_excel_report(parsed)
                st.download_button(
                    label="📥 下載 Excel 報表",
                    data=excel_data,
                    file_name="軋輥組裝報表整理.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary"
                )
                # ──────────────────────────────────────────────────

                CATEGORIES = ["粗車", "再生", "精車", "軸位粗車", "軸位再生", "軸位精車"]
                for cat in CATEGORIES:
                    types_dict = parsed.get(cat, {})
                    total_count = sum(len(rollers) for rollers in types_dict.values())
                    
                    label = f"**{cat}**（{total_count} 件）" if total_count else f"{cat}（無施做）"
                    with st.expander(label, expanded=(total_count > 0)):
                        if total_count > 0:
                            for r_type, rollers in types_dict.items():
                                if rollers:
                                    st.markdown(f"**🔹 型號：{r_type}**")
                                    for roller_id, size in rollers.items():
                                        st.write(f"　- `{roller_id}` → **{size}**")
                        else:
                            st.caption("本次無此項目。")

                st.divider()
                st.caption("📋 原始 JSON")
                st.json(parsed)

            except json.JSONDecodeError as e:
                st.error(f"JSON 解析失敗：{e}")
                st.code(raw_text, language="json")
            except Exception as e:
                st.error(f"Gemini 辨識錯誤：{e}")
