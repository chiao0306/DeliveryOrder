import streamlit as st
import os
import json
import base64
import requests
from azure.core.credentials import AzureKeyCredential
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest

st.set_page_config(page_title="Azure OCR 測試", layout="wide")

# ── 側邊欄：Gemini 模型選擇 ──────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 設定")
    gemini_model = st.selectbox(
        "Gemini 模型",
        options=[
            "gemini-2.5-flash",
            "gemini-2.5-flash-lite-preview-06-17",
        ],
        index=0,
        format_func=lambda x: {
            "gemini-2.5-flash": "Gemini 2.5 Flash（預設）",
            "gemini-2.5-flash-lite-preview-06-17": "Gemini 2.5 Flash Lite",
        }[x],
    )
    st.caption(f"目前選用：`{gemini_model}`")

st.title("👁️ 軋輥組裝報表 OCR 分析")

AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")
AZURE_KEY = os.environ.get("AZURE_KEY", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")

uploaded_file = st.file_uploader(
    "上傳報表圖檔（如：五號機軋輥組裝報表）", type=["jpg", "jpeg", "png", "pdf"]
)

if uploaded_file:
    file_content = uploaded_file.getvalue()

    # ── 兩欄並排 ────────────────────────────────────────────────
    col_azure, col_gemini = st.columns(2)

    # ════════════════════════════════════════════════════════════
    # 左欄：Azure Document Intelligence
    # ════════════════════════════════════════════════════════════
    with col_azure:
        st.subheader("🔷 Azure Document Intelligence")
        if not AZURE_ENDPOINT or not AZURE_KEY:
            st.error("⚠️ 請設定 AZURE_ENDPOINT 與 AZURE_KEY 環境變數。")
        else:
            with st.spinner("傳送至 Azure 辨識中…"):
                try:
                    client = DocumentIntelligenceClient(
                        endpoint=AZURE_ENDPOINT,
                        credential=AzureKeyCredential(AZURE_KEY),
                        api_version="2024-02-29-preview",
                    )
                    poller = client.begin_analyze_document(
                        model_id="prebuilt-layout",
                        body=file_content,
                    )
                    result = poller.result()
                    st.success("✅ Azure 辨識成功！")

                    if result.tables:
                        st.write(f"偵測到 **{len(result.tables)}** 個表格")
                        for table_idx, table in enumerate(result.tables):
                            with st.expander(
                                f"📌 表格 {table_idx + 1}（{table.row_count} 列）",
                                expanded=(table_idx == 0),
                            ):
                                table_data = {}
                                for cell in table.cells:
                                    r, c = cell.row_index, cell.column_index
                                    if r not in table_data:
                                        table_data[r] = {}
                                    table_data[r][c] = cell.content.replace("\n", " ")
                                for r in sorted(table_data.keys()):
                                    row_text = " | ".join(
                                        [
                                            table_data[r].get(c, "")
                                            for c in range(table.column_count)
                                        ]
                                    )
                                    st.code(row_text, language="markdown")
                    else:
                        st.warning("⚠️ 未偵測到表格結構。")

                except Exception as e:
                    st.error(f"🔍 端點：[{AZURE_ENDPOINT}]")
                    st.error(f"辨識錯誤：{e}")

    # ════════════════════════════════════════════════════════════
    # 右欄：Gemini 視覺辨識 → JSON 結構化輸出
    # ════════════════════════════════════════════════════════════
    with col_gemini:
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