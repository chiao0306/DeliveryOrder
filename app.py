import streamlit as st
import os
from azure.core.credentials import AzureKeyCredential
# 💡 修正最新版微軟 SDK 的正確導入路徑
from azure.ai.documentintelligence import DocumentIntelligenceClient
from azure.ai.documentintelligence.models import AnalyzeDocumentRequest

st.set_page_config(page_title="Azure OCR 測試", layout="centered")
st.title("👁️ Azure OCR 表格切割測試")

AZURE_ENDPOINT = os.environ.get("AZURE_ENDPOINT", "")
AZURE_KEY = os.environ.get("AZURE_KEY", "")

uploaded_file = st.file_uploader("上傳報表圖檔 (如：五號機輥輪組裝報表)", type=['jpg', 'png', 'pdf'])

if uploaded_file:
    if not AZURE_ENDPOINT or not AZURE_KEY:
        st.error("⚠️ 請至 Cloud Run 設定 AZURE_ENDPOINT 與 AZURE_KEY 環境變數。")
    else:
        with st.spinner("正在將檔案傳送至 Azure Document Intelligence..."):
            try:
                # 💡 加上 api_version 參數確保雲端相容性
                client = DocumentIntelligenceClient(
                    endpoint=AZURE_ENDPOINT,
                    credential=AzureKeyCredential(AZURE_KEY),
                    api_version="2024-11-30-preview" # 使用穩定的最新預覽版
                )
                
                # 讀取二進位圖檔
                file_content = uploaded_file.getvalue()
                
                # 呼叫 layout 模型分析表格
                poller = client.begin_analyze_document(
                    model_id="prebuilt-layout",
                    body=file_content
                )
                    analyze_request=file_content,
                    content_type=uploaded_file.type
                )
                result = poller.result()

                st.success("✅ Azure 辨識成功！")

                if result.tables:
                    st.subheader(f"📊 偵測到 {len(result.tables)} 個表格區域")
                    for table_idx, table in enumerate(result.tables):
                        with st.expander(f"📌 表格 {table_idx + 1} (共 {table.row_count} 列)", expanded=True):
                            table_data = {}
                            for cell in table.cells:
                                r, c = cell.row_index, cell.column_index
                                if r not in table_data:
                                    table_data[r] = {}
                                table_data[r][c] = cell.content.replace("\n", " ")

                            for r in sorted(table_data.keys()):
                                # 💡 防止某些格子空值造成報錯，加上 get 保底
                                row_text = " | ".join([table_data[r].get(c, "") for c in range(table.column_count)])
                                st.code(row_text, language="markdown")
                else:
                    st.warning("⚠️ Azure 沒偵測到任何表格結構。")

            except Exception as e:
                st.error(f"辨識發生錯誤：{e}")
