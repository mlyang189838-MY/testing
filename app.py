# =====================================================
# MP3 語音轉文字 -> Word (.docx)
# Streamlit Web App
# =====================================================

import os
import tempfile
import streamlit as st
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from faster_whisper import WhisperModel

# =====================================================
# Streamlit 頁面設定
# =====================================================
st.set_page_config(
    page_title="MP3 語音轉 Word 文字檔",
    page_icon="🎧",
    layout="centered"
)

st.title("🎧 MP3 語音轉文字 → Word 文字檔")
st.write("上傳 MP3 檔案，選擇語言與模型，即可轉換為文字並下載 Word 檔。")

# =====================================================
# 側邊欄設定
# =====================================================
with st.sidebar:
    st.header("轉錄設定")

    model_size = st.selectbox(
        "Whisper 模型",
        ["tiny", "base", "small", "medium", "large-v3"],
        index=2,
        help="模型越大越準確，但速度較慢、需要較多資源。"
    )

    language = st.selectbox(
        "語音語言",
        ["自動偵測", "中文", "粵語", "英文"],
        index=0
    )

    language_code = {
        "自動偵測": None,
        "中文": "zh",
        "粵語": "yue",
        "英文": "en"
    }[language]

    st.info(
        "建議：\n\n"
        "- 短音檔或測試：`base`\n"
        "- 一般會議、訪談：`small`\n"
        "- 中文／粵語長錄音：`medium`\n"
        "- 高準確度需求：`large-v3`"
    )

# =====================================================
# 上傳 MP3
# =====================================================
uploaded_file = st.file_uploader(
    "請上傳 MP3 語音檔",
    type=["mp3"],
    accept_multiple_files=False
)

if uploaded_file is None:
    st.warning("請先上傳 MP3 檔案。")
    st.stop()

st.success(f"已上傳：{uploaded_file.name}")

# =====================================================
# 開始轉錄按鈕
# =====================================================
if st.button("開始語音轉文字", type="primary", use_container_width=True):

    with st.status("正在處理語音檔，請稍候...", expanded=True) as status:

        st.write("1. 儲存上传的 MP3 檔案...")
        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=".mp3"
        ) as tmp_file:
            tmp_file.write(uploaded_file.getvalue())
            mp3_path = tmp_file.name

        try:
            st.write(f"2. 載入 Whisper 模型：{model_size} ...")
            model = WhisperModel(
                model_size,
                device="cpu",
                compute_type="int8"
            )

            st.write("3. 正在進行語音轉文字...")
            segments, info = model.transcribe(
                mp3_path,
                language=language_code,
                task="transcribe",
                vad_filter=True
            )

            transcript_parts = []
            progress_bar = st.progress(0, text="正在轉錄...")
            segment_count = 0

            for segment in segments:
                transcript_parts.append(segment.text.strip())
                segment_count += 1
                progress_bar.progress(
                    min(segment_count * 5, 95),
                    text=f"已處理 {segment_count} 個語音段落..."
                )

            transcript = "\n".join(transcript_parts).strip()

            if not transcript:
                status.update(
                    label="未能從音檔辨識出文字。",
                    state="error"
                )
                st.error("未能從音檔辨識出文字，請確認 MP3 是否包含清晰語音。")
                st.stop()

            progress_bar.progress(100, text="語音轉文字完成！")

            st.write("4. 正在建立 Word 文件...")
            doc = Document()

            style = doc.styles["Normal"]
            style.font.name = "Microsoft JhengHei"
            style.font.size = Pt(12)

            title = doc.add_heading("MP3 語音轉文字結果", level=0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph(f"來源語音檔：{uploaded_file.name}")
            doc.add_paragraph(f"辨識語言：{info.language}")
            doc.add_paragraph("")

            doc.add_heading("轉錄內容", level=1)
            doc.add_paragraph(transcript)

            word_path = "語音轉文字結果.docx"
            doc.save(word_path)

            with open(word_path, "rb") as word_file:
                word_bytes = word_file.read()

            status.update(
                label="轉錄完成！",
                state="complete"
            )

            st.subheader("轉錄結果")
            st.text_area(
                "文字內容",
                value=transcript,
                height=350
            )

            st.download_button(
                label="⬇️ 下載 Word 文字檔",
                data=word_bytes,
                file_name="語音轉文字結果.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
                type="primary",
                use_container_width=True
            )

        except Exception as e:
            status.update(
                label="轉錄失敗。",
                state="error"
            )
            st.error(f"處理時發生錯誤：{e}")

        finally:
            if os.path.exists(mp3_path):
                os.remove(mp3_path)

            if os.path.exists("語音轉文字結果.docx"):
                os.remove("語音轉文字結果.docx")

# =====================================================
# 頁尾說明
# =====================================================
st.divider()
st.caption(
    "提示：轉錄結果可能受音質、背景噪音、多人同時說話及語言混合程度影響。"
)
