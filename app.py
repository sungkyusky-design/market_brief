"""
신한 금융시장 Brief 웹 앱 (Streamlit)

로컬 실행:
    streamlit run app.py

배포:
    Streamlit Community Cloud에서 이 repo 연결 → Secrets에 GEMINI_API_KEY 등록
"""

import os
import sys
from pathlib import Path
from datetime import datetime

import streamlit as st
import yaml

# Streamlit Cloud에선 st.secrets, 로컬에선 .env 사용
try:
    if "GEMINI_API_KEY" in st.secrets:
        os.environ["GEMINI_API_KEY"] = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass  # secrets.toml 없음 → .env로 fallback

from news_collector import NewsCollector
from doc_generator import DocGenerator

TEMPLATE_PATH = "금융시장브리프 양식.docx"


@st.cache_data
def load_config():
    with open("config.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


st.set_page_config(page_title="금시브 생성기", page_icon="📊", layout="centered")
st.title("📊 신한 금융시장 Brief 생성기")

config = load_config()

col1, col2 = st.columns(2)
with col1:
    issue_num = st.number_input(
        "호수", min_value=1, value=int(config["issue"]["number"]), step=1
    )
with col2:
    dry_run = st.toggle("테스트 모드 (API 미호출)", value=False)

with st.expander("주제 목록 확인", expanded=False):
    for t in config["topics"]:
        st.write(f"- **{t['name']}** ({t['count']}개) — {t.get('author', '')}")

st.divider()

if st.button("🚀 문서 생성", type="primary", use_container_width=True):
    config["issue"]["number"] = int(issue_num)

    progress = st.progress(0, text="준비 중...")
    log = st.empty()

    try:
        # 1단계: 기사 수집
        if dry_run:
            from main import generate_dummy_articles
            progress.progress(20, text="더미 기사 생성 중...")
            articles = generate_dummy_articles(config)
        else:
            if not os.getenv("GEMINI_API_KEY"):
                st.error("GEMINI_API_KEY가 설정되지 않았습니다. Secrets를 확인하세요.")
                st.stop()
            progress.progress(10, text="뉴스 수집 시작 (몇 분 걸립니다)...")
            collector = NewsCollector(config)
            articles = collector.collect_all()

        progress.progress(70, text=f"기사 {len(articles)}개 수집 완료. 문서 생성 중...")

        # 2단계: 문서 생성
        generator = DocGenerator(config)
        if Path(TEMPLATE_PATH).exists():
            output_path = generator.generate_from_template(
                articles, template_path=TEMPLATE_PATH
            )
        else:
            output_path = generator.generate(articles)

        progress.progress(100, text="완료!")

        success = sum(1 for a in articles if a.get("source") != "수집 실패")
        st.success(f"✅ 생성 완료 — 기사 {success}/{len(articles)}개 수록")

        with open(output_path, "rb") as f:
            st.download_button(
                label="📥 Word 파일 다운로드",
                data=f.read(),
                file_name=Path(output_path).name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )

        if success < len(articles):
            st.warning(f"⚠️ {len(articles) - success}개 기사 수집 실패 — 수동 확인 필요")

    except Exception as e:
        progress.empty()
        st.error(f"오류 발생: {e}")
        import traceback
        st.code(traceback.format_exc())
