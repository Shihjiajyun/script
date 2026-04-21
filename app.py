import os
import tempfile
from pathlib import Path

import streamlit as st
import streamlit.components.v1 as components

import card

st.set_page_config(
    page_title="司儀稿手卡生成器",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ── 全站樣式 ──────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,600;1,400;1,600&family=Noto+Sans+TC:wght@400;500;600&display=swap');

/* ── reset ── */
*, *::before, *::after { box-sizing: border-box; }

/* ── body / app shell ── */
html, body { background: #f8f7f4 !important; }
.stApp    { background: #f8f7f4 !important; }
[data-testid="stAppViewContainer"] { background: #f8f7f4 !important; }
[data-testid="stHeader"] { background: transparent !important; display: none; }
[data-testid="stToolbar"] { display: none; }
#MainMenu, footer { display: none !important; }
.block-container {
    max-width: 780px;
    padding: 3rem 2rem 5rem;
}

/* ── typography base ── */
html, body, .stApp, .stMarkdown, p, span, label, div {
    font-family: 'Noto Sans TC', 'PingFang TC', 'Microsoft JhengHei', sans-serif !important;
    color: #18160e;
}

/* ── site header ── */
.site-header {
    padding-bottom: 2.5rem;
    margin-bottom: .5rem;
}
.site-header-cn {
    font-family: 'Noto Sans TC', sans-serif !important;
    font-size: clamp(1.6rem, 4vw, 2.2rem);
    font-weight: 600;
    color: #18160e;
    line-height: 1.2;
    letter-spacing: -.02em;
    margin: 0 0 .25rem;
}
.site-header-en {
    font-family: 'Cormorant Garamond', serif !important;
    font-style: italic;
    font-size: 1.05rem;
    color: #1d35a8;
    letter-spacing: .02em;
    margin: 0 0 1rem;
}
.site-header-rule {
    border: none;
    border-top: 1.5px solid #e6e3dc;
    margin: 1.5rem 0 0;
}

/* ── steps ── */
.step-row {
    display: flex;
    align-items: flex-start;
    gap: 1.75rem;
    padding: 2.25rem 0;
    border-bottom: 1px solid #e6e3dc;
}
.step-num {
    font-family: 'Cormorant Garamond', serif !important;
    font-weight: 600;
    font-size: 3.5rem;
    line-height: 1;
    color: #dedad3;
    flex-shrink: 0;
    width: 3.5rem;
    text-align: right;
    padding-top: .1rem;
    user-select: none;
}
.step-body { flex: 1; min-width: 0; }
.step-title {
    font-size: .8rem;
    font-weight: 600;
    letter-spacing: .1em;
    text-transform: uppercase;
    color: #1d35a8;
    margin: 0 0 .35rem;
}
.step-desc {
    font-size: .9rem;
    color: #5c5a54;
    line-height: 1.7;
    margin: 0 0 1.1rem;
}
.step-desc a { color: #1d35a8; text-decoration: underline; text-underline-offset: 2px; }

/* ── prompt box ── */
.prompt-wrap {
    position: relative;
    border-left: 3px solid #1d35a8;
    background: #edeae3;
    border-radius: 0 6px 6px 0;
    padding: 1rem 1.25rem;
}
.prompt-label {
    font-size: .7rem;
    font-weight: 600;
    letter-spacing: .08em;
    text-transform: uppercase;
    color: #1d35a8;
    margin-bottom: .5rem;
}
.prompt-text {
    font-size: .82rem;
    color: #3a3830;
    line-height: 1.8;
    white-space: pre-wrap;
    word-break: break-word;
    margin: 0;
}

/* ── section divider ── */
.section-sep {
    border: none;
    border-top: 1px solid #e6e3dc;
    margin: 2rem 0 1.5rem;
}

/* ── preview header ── */
.preview-head {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    margin-bottom: 1rem;
}
.preview-title {
    font-size: .8rem;
    font-weight: 600;
    letter-spacing: .1em;
    text-transform: uppercase;
    color: #1d35a8;
    margin: 0;
}
.preview-meta {
    font-size: .8rem;
    color: #9c9890;
}

/* ── download button ── */
div[data-testid="stDownloadButton"] {
    margin-top: 1.25rem;
}
div[data-testid="stDownloadButton"] > button {
    width: 100% !important;
    background: #1d35a8 !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 6px !important;
    padding: .75rem 1.5rem !important;
    font-size: .9rem !important;
    font-weight: 500 !important;
    font-family: 'Noto Sans TC', sans-serif !important;
    letter-spacing: .03em !important;
    transition: background .15s ease !important;
    cursor: pointer !important;
}
div[data-testid="stDownloadButton"] > button:hover {
    background: #162990 !important;
}
div[data-testid="stDownloadButton"] > button:active {
    background: #0f1e6e !important;
}

/* ── slider ── */
[data-testid="stSlider"] [data-testid="stWidgetLabel"] p {
    font-size: .82rem !important;
    font-weight: 500 !important;
    color: #3a3830 !important;
    margin-bottom: .3rem !important;
}
[data-testid="stSlider"] .st-emotion-cache-1gv3huu,
[data-testid="stSlider"] span {
    color: #18160e !important;
}

/* ── checkbox ── */
[data-testid="stCheckbox"] label {
    font-size: .82rem !important;
    color: #3a3830 !important;
}
[data-testid="stCheckbox"] label p {
    font-size: .82rem !important;
    color: #3a3830 !important;
}

/* ── file uploader ── */
[data-testid="stFileUploader"] section {
    border: 1.5px dashed #c8c4bb !important;
    border-radius: 8px !important;
    background: #f0ede8 !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #1d35a8 !important;
    background: #eceaf8 !important;
}
[data-testid="stFileUploaderDropzone"] p,
[data-testid="stFileUploaderDropzone"] span,
[data-testid="stFileUploaderDropzoneInstructions"] p,
[data-testid="stFileUploaderDropzoneInstructions"] span {
    color: #5c5a54 !important;
    font-size: .85rem !important;
}
[data-testid="stFileUploader"] button {
    background: #18160e !important;
    color: #f8f7f4 !important;
    border: none !important;
    border-radius: 5px !important;
    font-size: .8rem !important;
    font-family: 'Noto Sans TC', sans-serif !important;
}
.uploadedFileName { color: #18160e !important; }

/* ── alert / info boxes ── */
[data-testid="stAlert"] { border-radius: 6px !important; }

/* ── footer ── */
.site-footer {
    margin-top: 4rem;
    padding-top: 1.25rem;
    border-top: 1px solid #e6e3dc;
    font-size: .78rem;
    color: #b0ada6;
    display: flex;
    gap: 1.5rem;
}
.site-footer a {
    color: #b0ada6 !important;
    text-decoration: none !important;
}
.site-footer a:hover { color: #1d35a8 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────
CLAUDE_PROMPT = """請幫我將以下司儀稿整理成 MD 格式，規則如下：
- 用 ## 標記每個段落的標題
- 每一段對話格式為「姓名：台詞內容」（全形冒號）
- 同一位司儀的連續對話，如果太長請適當拆成多個「姓名：」段落，每段盡量不超過 100 字
- 台詞中的 ★ 請保留
- 空行用來分隔不同說話者

以下是司儀稿內容：
[請在這裡貼上你的司儀稿]"""

# ── 標題 ─────────────────────────────────────────────────────────
st.markdown("""
<div class="site-header">
    <p class="site-header-cn">司儀稿手卡生成器</p>
    <p class="site-header-en">MC Script Card Generator</p>
    <hr class="site-header-rule">
</div>
""", unsafe_allow_html=True)

# ── 步驟一 ────────────────────────────────────────────────────────
prompt_escaped = CLAUDE_PROMPT.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
st.markdown(f"""
<div class="step-row">
    <div class="step-num">01</div>
    <div class="step-body">
        <p class="step-title">整理稿子</p>
        <p class="step-desc">
            前往 <a href="https://claude.ai" target="_blank">claude.ai</a>，
            複製下方 Prompt，貼入後將司儀稿附在最後，讓 Claude 自動整理成 MD 格式。
        </p>
        <div class="prompt-wrap">
            <p class="prompt-label">Prompt</p>
            <p class="prompt-text">{prompt_escaped}</p>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── 步驟二 ────────────────────────────────────────────────────────
st.markdown("""
<div class="step-row">
    <div class="step-num">02</div>
    <div class="step-body">
        <p class="step-title">上傳 MD 檔案</p>
        <p class="step-desc">將 Claude 輸出的內容存成 .md 檔案後上傳。</p>
    </div>
</div>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("", type=["md"], label_visibility="collapsed")

# ── 步驟三 ────────────────────────────────────────────────────────
st.markdown("""
<div class="step-row">
    <div class="step-num">03</div>
    <div class="step-body">
        <p class="step-title">設定選項</p>
        <p class="step-desc">調整每格字數上限，以及是否在右下角顯示格號。</p>
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns([3, 2], gap="large")
with col1:
    max_chars = st.slider("每格最大字數", min_value=60, max_value=150, value=120, step=5)
with col2:
    st.markdown("<div style='height:1.6rem'></div>", unsafe_allow_html=True)
    show_pagenum = st.checkbox("每格右下角顯示格號")


# ── 輔助：HTML 手卡預覽 ───────────────────────────────────────────

def _esc(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def render_line_html(text: str) -> str:
    parts = text.split("★")
    out = []
    for i, part in enumerate(parts):
        out.append(_esc(part))
        if i < len(parts) - 1:
            out.append('<span class="star">★</span>')
    return "".join(out)


def build_cell_html(cell_blocks: list, cell_number: int | None, is_empty: bool) -> str:
    if is_empty:
        return '<div class="card-cell card-cell--empty"></div>'

    inner = []
    first_section = next((b.section for b in cell_blocks if b.section), "")
    if first_section:
        inner.append(f'<div class="section-hd">《{_esc(first_section)}》</div>')

    for blk in cell_blocks:
        first_line = blk.lines[0] if blk.lines else ""
        inner.append(
            f'<p class="dialogue">'
            f'<span class="spk">{_esc(blk.speaker)}：</span>'
            f'{render_line_html(first_line)}</p>'
        )
        for extra in blk.lines[1:]:
            inner.append(f'<p class="cont">{render_line_html(extra)}</p>')

    num_html = f'<div class="cell-num">{cell_number}</div>' if cell_number is not None else ""
    return f'<div class="card-cell">{"".join(inner)}{num_html}</div>'


def build_preview_html(cells: list, show_pagenum: bool) -> str:
    CELLS_PER_PAGE = 8
    padded = list(cells)
    if len(padded) % 2 != 0:
        padded.append([])

    pages_html = []
    for page_start in range(0, len(padded), CELLS_PER_PAGE):
        page_cells = padded[page_start: page_start + CELLS_PER_PAGE]
        while len(page_cells) < CELLS_PER_PAGE:
            page_cells.append([])

        cells_html = []
        for local_idx, cb in enumerate(page_cells):
            global_idx = page_start + local_idx
            is_empty = len(cb) == 0
            num = (global_idx + 1) if (show_pagenum and not is_empty) else None
            cells_html.append(build_cell_html(cb, num, is_empty))

        page_n = page_start // CELLS_PER_PAGE + 1
        pages_html.append(
            f'<div class="page">'
            f'<div class="page-label">Page {page_n}</div>'
            f'<div class="card-grid">{"".join(cells_html)}</div>'
            f'</div>'
        )

    return f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="utf-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    background: #f0ede8;
    font-family: '標楷體', 'DFKai-SB', 'BiauKai', serif;
    padding: 16px 10px 24px;
  }}
  .page {{ margin: 0 auto 24px; max-width: 740px; }}
  .page-label {{
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-size: 9px;
    font-weight: 600;
    letter-spacing: .1em;
    text-transform: uppercase;
    color: #9c9890;
    margin-bottom: 5px;
    padding-left: 1px;
  }}
  .card-grid {{
    display: grid;
    grid-template-columns: 1fr 1fr;
    grid-template-rows: repeat(4, auto);
    gap: 0;
    border: 1.5px solid #2a2820;
    border-radius: 3px;
    overflow: hidden;
    box-shadow: 0 2px 12px rgba(24,22,14,.1);
  }}
  .card-cell {{
    background: #fffffe;
    min-height: 145px;
    padding: 10px 12px 22px;
    position: relative;
    display: flex;
    flex-direction: column;
    border: 1px solid #dedad3;
    font-size: 11.5px;
    line-height: 1.65;
    color: #18160e;
    word-break: break-all;
  }}
  .card-cell:nth-child(odd)  {{ border-right: 1.5px solid #2a2820; }}
  .card-cell:nth-child(even) {{ border-left: none; }}
  .card-cell--empty {{ background: #f5f2ed; }}
  .section-hd {{
    font-size: 10px;
    font-weight: bold;
    color: #1d35a8;
    margin-bottom: 3px;
    letter-spacing: .02em;
  }}
  .dialogue {{ text-align: justify; margin-bottom: 1px; }}
  .spk {{ font-weight: bold; color: #18160e; }}
  .star {{ color: #c00; font-weight: bold; }}
  .cont {{ text-align: justify; padding-left: 2.4em; margin-bottom: 1px; }}
  .cell-num {{
    position: absolute;
    bottom: 6px;
    right: 10px;
    font-size: 8.5px;
    color: #aaa9a4;
    font-family: 'Helvetica Neue', Arial, sans-serif;
    font-weight: 500;
    letter-spacing: .02em;
  }}
</style>
</head>
<body>{"".join(pages_html)}</body>
</html>"""


# ── 預覽 & 下載 ───────────────────────────────────────────────────
if uploaded_file is not None:
    md_text = uploaded_file.read().decode("utf-8")
    blocks = card.parse_md(md_text)
    cells = card.layout_blocks(blocks, max_chars=max_chars)
    total_pages = (len(cells) + 7) // 8

    st.markdown('<hr class="section-sep">', unsafe_allow_html=True)
    st.markdown(f"""
<div class="preview-head">
    <p class="preview-title">手卡預覽</p>
    <span class="preview-meta">{len(cells)} 格 &nbsp;·&nbsp; {total_pages} 頁</span>
</div>
""", unsafe_allow_html=True)

    preview_html = build_preview_html(cells, show_pagenum)
    iframe_h = max(520, min(total_pages * 660, 900))
    components.html(preview_html, height=iframe_h, scrolling=True)

    # 產生 DOCX
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        card.generate_docx(cells, tmp_path, show_pagenum=show_pagenum)
        with open(tmp_path, "rb") as f:
            docx_bytes = f.read()
    finally:
        os.unlink(tmp_path)

    stem = Path(uploaded_file.name).stem
    st.download_button(
        label="下載手卡 DOCX",
        data=docx_bytes,
        file_name=f"{stem}_手卡.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )

# ── 頁尾 ──────────────────────────────────────────────────────────
st.markdown("""
<div class="site-footer">
    <a href="mailto:shihjiajyun@gmail.com">shihjiajyun@gmail.com</a>
    <a href="https://github.com/Shihjiajyun/script" target="_blank">GitHub</a>
</div>
""", unsafe_allow_html=True)
