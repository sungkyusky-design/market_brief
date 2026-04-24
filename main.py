"""
신한 금융시장 Brief 자동 생성기

실행 방법:
    python main.py                  # 양식 파일 기반 자동 실행 (기본)
    python main.py --no-template    # 양식 없이 프로그래매틱 생성
    python main.py --dry-run        # API 호출 없이 더미 데이터로 테스트

사전 준비:
    1. pip install -r requirements.txt
    2. .env 파일에 GEMINI_API_KEY 입력
    3. config.yaml에서 호수(number) 확인/수정
"""

import sys
from pathlib import Path

import yaml

from news_collector import NewsCollector
from doc_generator import DocGenerator

# 양식 파일 경로 (프로젝트 루트 기준)
# .docx 우선 사용 (서식 보존), 없으면 .doc 사용
TEMPLATE_PATH = "금융시장브리프 양식.docx"
TEMPLATE_PATH_DOC = "금융시장브리프 양식.doc"


def load_config(path: str = "config.yaml") -> dict:
    """설정 파일 로드"""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return yaml.safe_load(f)
    except FileNotFoundError:
        print(f"오류: 설정 파일 '{path}'을(를) 찾을 수 없습니다.")
        sys.exit(1)
    except yaml.YAMLError as e:
        print(f"오류: 설정 파일 파싱 실패 - {e}")
        sys.exit(1)


def generate_dummy_articles(config: dict) -> list[dict]:
    """API 호출 없이 문서 레이아웃 테스트용 더미 데이터 생성"""
    articles = []
    topics = config["topics"]

    for topic in topics:
        name = topic["name"]
        count = topic.get("count", 1)
        author = topic.get("author", "")
        for i in range(count):
            articles.append({
                "topic": name,
                "title": f"{name} 주요 뉴스 제목입니다",
                "lines": [
                    f"{name} 관련 첫 번째 요약 줄입니다. 구체적 수치와 팩트를 포함한 긴 문장으로 70~90자 범위 테스트용입니다.",
                    f"{name} 관련 두 번째 요약 줄입니다. 2026년 전망 수치, 전년 대비 증감률 등 구체적 데이터를 포함합니다.",
                    f"{name} 관련 세 번째 요약 줄입니다. 주요 기업명과 구체적 금액, 날짜 등 팩트 중심으로 서술합니다.",
                ],
                "source": "테스트 데이터",
                "author": author,
            })
    return articles


def main():
    """메인 실행 함수"""
    config = load_config()
    issue_num = config["issue"]["number"]
    title = config["issue"]["title"]

    is_dry_run = "--dry-run" in sys.argv
    use_template = "--no-template" not in sys.argv

    print("=" * 52)
    print(f"  {title} {issue_num}호 자동 생성기")
    if use_template:
        print(f"  모드: 양식 파일 기반 생성")
    else:
        print(f"  모드: 프로그래매틱 생성")
    print("=" * 52)

    # ── 1단계: 뉴스 수집 ──
    if is_dry_run:
        print("\n[DRY RUN] API 호출 없이 더미 데이터로 문서를 생성합니다.")
        articles = generate_dummy_articles(config)
        print(f"  더미 기사 {len(articles)}개 생성 완료")
    else:
        print("\n[1/2] 뉴스 수집 시작")
        try:
            collector = NewsCollector(config)
            articles = collector.collect_all()
        except ValueError as e:
            print(f"\n오류: {e}")
            sys.exit(1)
        except Exception as e:
            print(f"\n예기치 않은 오류: {e}")
            sys.exit(1)

    # ── 2단계: 문서 생성 ──
    step = "DRY RUN" if is_dry_run else "2/2"
    print(f"\n[{step}] Word 문서 생성 중...")

    try:
        generator = DocGenerator(config)

        if use_template:
            # .docx 양식 우선, 없으면 .doc 사용
            template_to_use = None
            if Path(TEMPLATE_PATH).exists():
                template_to_use = TEMPLATE_PATH
            elif Path(TEMPLATE_PATH_DOC).exists():
                template_to_use = TEMPLATE_PATH_DOC
                print(f"  주의: .docx 양식이 없어 .doc를 변환합니다. 서식 손실 가능성 있음.")
            
            if template_to_use:
                print(f"  양식 파일 사용: {template_to_use}")
                output_path = generator.generate_from_template(
                    articles, template_path=template_to_use
                )
            else:
                print(f"  주의: 양식 파일을 찾을 수 없어 프로그래매틱 생성으로 전환합니다.")
                output_path = generator.generate(articles)
        else:
            output_path = generator.generate(articles)
    except Exception as e:
        print(f"\n문서 생성 오류: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    # ── 완료 ──
    total = len(articles)
    success = sum(1 for a in articles if a.get("source") != "수집 실패")

    print("\n" + "=" * 52)
    print(f"  생성 완료!")
    print(f"  파일: {output_path}")
    print(f"  기사: {success}/{total}개 수록")
    if success < total:
        print(f"  주의: {total - success}개 기사 수집 실패 (수동 수정 필요)")
    print("=" * 52)


if __name__ == "__main__":
    main()
