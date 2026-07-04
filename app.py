import streamlit as st
import os
import json
import base64
import requests

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

st.title("軋輥組裝報表 OCR 分析")

# 只保留 Gemini 的環境變數
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

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
1. 掃描每一筆軋輥記錄（每一列）。
2. 欄位對應如下：
   - 粗車（Roll 粗車尺寸，非軸位）
   - 再生（Roll 再生尺寸）
   - 精車（Roll 精車尺寸）
   - 軸位粗車（軸位粗車尺寸）
   - 軸位再生（軸位再生尺寸）
   - 軸位精車（軸位精車尺寸）
3. 若該格為數字（含小數），代表有施做，請記錄該數字（字串格式）。
4. 若該格為「X」或空白，代表未施做，請略過（不要包含在輸出中）。
5. 每個類別只輸出「有施做」的項目，以 { "編號": "尺寸" } 的鍵值對表示。

輸出格式範例（只輸出 JSON，不要加任何說明文字）：
{
  "粗車": { "M30S141": "307" },
  "再生": {},
  "精車": { "M30S141": "300.07", "V30S58": "300.02" },
  "軸位粗車": { "V30S58": "146" },
  "軸位再生": { "V30S58": "155" },
  "軸位精車": { "V30S58": "149.97" }
}

請直接輸出 JSON，不要加 markdown 代碼區塊。
""".strip()

        with st.spinner(f"使用 {gemini_model} 辨識中…"):
            try:
                # 判斷檔案類型
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

                # 清除可能殘留的 markdown fence
                if raw_text.startswith("```"):
                    raw_text = raw_text.split("```")[1]
                    if raw_text.startswith("json"):
                        raw_text = raw_text[4:]
                    raw_text = raw_text.strip()

                parsed = json.loads(raw_text)
                st.success("✅ Gemini 辨識成功！")

                CATEGORIES = ["粗車", "再生", "精車", "軸位粗車", "軸位再生", "軸位精車"]
                for cat in CATEGORIES:
                    items = parsed.get(cat, {})
                    count = len(items)
                    label = f"**{cat}**（{count} 件）" if count else f"{cat}（無施做）"
                    with st.expander(label, expanded=(count > 0)):
                        if items:
                            for roller_id, size in items.items():
                                st.write(f"- `{roller_id}` → **{size}**")
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
