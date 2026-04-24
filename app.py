"""
신한 금융시장 Brief 웹 앱 (Streamlit)
"""

import os
import contextlib
from pathlib import Path

import streamlit as st
import yaml

# Streamlit Cloud: st.secrets, 로컬: .env
try:
    if "GEMINI_API_KEY" in st.secrets:
        os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

from news_collector import NewsCollector
from doc_generator import DocGenerator

TEMPLATE_PATH = "금융시장브리프 양식.docx"

# ─────────────────────────────────────────────
# Page config & global styles
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Shinhan Financial Market Brief",
    layout="centered",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        #MainMenu, footer, header {visibility: hidden;}

        .main .block-container {
            padding-top: 3.5rem;
            padding-bottom: 4rem;
            max-width: 760px;
        }

        h1 {
            font-weight: 700;
            letter-spacing: -0.025em;
            font-size: 2.25rem !important;
            margin-bottom: 0.25rem !important;
        }
        h2, h3 { font-weight: 600; letter-spacing: -0.015em; }

        .stCaption, .caption { color: #6b7280; }

        section[data-testid="stSidebar"] {
            background-color: #fafafa;
            border-right: 1px solid #f0f0f0;
        }
        section[data-testid="stSidebar"] .block-container {
            padding-top: 2rem;
        }

        .stButton > button {
            border-radius: 10px;
            font-weight: 600;
            padding: 0.65rem 1rem;
            transition: all 0.15s ease;
            border: 1px solid #e5e7eb;
        }
        .stButton > button[kind="primary"] {
            background-color: #0064FF;
            border-color: #0064FF;
        }
        .stButton > button[kind="primary"]:hover {
            background-color: #0052d4;
            border-color: #0052d4;
        }

        .stDownloadButton > button {
            border-radius: 10px;
            font-weight: 600;
        }

        [data-testid="stExpander"] {
            border: 1px solid #eef0f3;
            border-radius: 10px;
        }

        .stProgress > div > div > div { background-color: #0064FF; }

        code { font-size: 0.82rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ─────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────
@st.cache_data
def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


class LogStream:
    """print() 출력을 Streamlit 영역에 실시간 스트리밍."""

    def __init__(self, placeholder, max_lines: int = 200):
        self.placeholder = placeholder
        self.max_lines = max_lines
        self.buffer: list[str] = []

    def write(self, text: str):
        if not text:
            return
        for line in text.splitlines():
            if line.strip():
                self.buffer.append(line.rstrip())
        self.placeholder.code(
            "\n".join(self.buffer[-self.max_lines :]),
            language="text",
        )

    def flush(self):
        pass


# ─────────────────────────────────────────────
# Sidebar (navigation)
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### Shinhan")
    st.caption("Research Tools")
    st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)

    page = st.radio(
        "navigation",
        ["금융시장브리프"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.caption("v1.0 · Research Center")


# ─────────────────────────────────────────────
# Page: 금융시장브리프
# ─────────────────────────────────────────────
if page == "금융시장브리프":
    config = load_config()

    st.title("금융시장브리프")
    st.caption(f"{config['issue']['subtitle']} · 자동 생성")
    st.markdown("<div style='height:1.25rem'></div>", unsafe_allow_html=True)

    col1, col2 = st.columns([2, 1])
    with col1:
        issue_num = st.number_input(
            "호수",
            min_value=1,
            value=int(config["issue"]["number"]),
            step=1,
        )
    with col2:
        st.markdown("<div style='height:1.8rem'></div>", unsafe_allow_html=True)
        dry_run = st.toggle(
            "테스트 모드",
            value=False,
            help="API 호출 없이 더미 데이터로 실행",
        )

    with st.expander("주제 구성 확인"):
        for t in config["topics"]:
            st.markdown(
                f"**{t['name']}** &nbsp;·&nbsp; {t['count']}개 "
                f"&nbsp;·&nbsp; <span style='color:#6b7280'>{t.get('author', '')}</span>",
                unsafe_allow_html=True,
            )

    st.markdown("<div style='height:0.75rem'></div>", unsafe_allow_html=True)
    run = st.button("문서 생성 시작", type="primary", use_container_width=True)

    if run:
        config["issue"]["number"] = int(issue_num)

        output_path = None
        articles = None
        error = None

        with st.status("실행 중입니다…", expanded=True) as status:
            log_area = st.empty()
            stream = LogStream(log_area)

            try:
                with contextlib.redirect_stdout(stream):
                    if dry_run:
                        print("[TEST] 더미 데이터 생성 중...")
                        from main import generate_dummy_articles

                        articles = generate_dummy_articles(config)
                        print(f"[TEST] 더미 기사 {len(articles)}개 생성 완료")
                    else:
                        if not os.getenv("GEMINI_API_KEY"):
                            raise RuntimeError(
                                "GEMINI_API_KEY가 설정되지 않았습니다. Secrets를 확인하세요."
                            )
                        print("[1/2] 뉴스 수집 시작")
                        collector = NewsCollector(config)
                        articles = collector.collect_all()

                    print("")
                    print("[2/2] Word 문서 생성 중...")
                    generator = DocGenerator(config)
                    if Path(TEMPLATE_PATH).exists():
                        output_path = generator.generate_from_template(
                            articles, template_path=TEMPLATE_PATH
                        )
                    else:
                        output_path = generator.generate(articles)
                    print(f"[완료] 생성 파일: {output_path}")

                status.update(label="완료", state="complete", expanded=False)

            except Exception as e:
                error = e
                import traceback
                stream.write(f"\n[ERROR] {e}\n{traceback.format_exc()}")
                status.update(label="오류 발생", state="error", expanded=True)

        if error is None and output_path:
            success = sum(1 for a in articles if a.get("source") != "수집 실패")
            total = len(articles)

            st.markdown("<div style='height:0.5rem'></div>", unsafe_allow_html=True)
            st.success(f"문서 생성 완료 · 기사 {success} / {total}개 수록")

            with open(output_path, "rb") as f:
                st.download_button(
                    label="Word 파일 다운로드",
                    data=f.read(),
                    file_name=Path(output_path).name,
                    mime=(
                        "application/vnd.openxmlformats-officedocument."
                        "wordprocessingml.document"
                    ),
                    use_container_width=True,
                )

            if success < total:
                st.warning(
                    f"{total - success}개 기사 수집에 실패했습니다. 수동 확인이 필요합니다."
                )
        elif error is not None:
            st.error(f"실행 중 오류가 발생했습니다: {error}")
