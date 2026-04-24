"""
뉴스 수집 모듈
Gemini API + Google Search grounding을 사용하여
주제별 최신 뉴스를 검색하고 요약합니다.
"""

import os
import json
import re
import time
from pathlib import Path

from google import genai
from google.genai.types import GenerateContentConfig, GoogleSearch, Tool
from dotenv import load_dotenv


class NewsCollector:
    """Gemini + Google Search 기반 뉴스 수집/요약 클래스"""

    def __init__(self, config: dict):
        # 스크립트 위치 기준으로 .env 파일 로드
        env_path = Path(__file__).resolve().parent / ".env"
        load_dotenv(dotenv_path=env_path)
        api_key = os.getenv("GEMINI_API_KEY")
        if not api_key or api_key == "your_api_key_here":
            raise ValueError(
                "GEMINI_API_KEY가 설정되지 않았습니다.\n"
                ".env 파일을 열어 API 키를 입력하세요.\n"
                "발급: https://aistudio.google.com/app/apikey"
            )

        self.client = genai.Client(api_key=api_key)
        self.config = config
        self.model_name = config["model"]["name"]
        self.temperature = config["model"]["temperature"]
        self.delay = config["model"]["delay_between_calls"]
        self.fmt = config["format"]

        # Google Search 도구 설정
        self.search_tool = Tool(google_search=GoogleSearch())

    def _build_prompt(self, topic_name: str, keywords: str, exclude_titles: list[str] = None, exclude_articles: list[dict] = None) -> str:
        """주제별 뉴스 검색 및 요약 프롬프트 생성"""
        title_min = self.fmt["title_min_chars"]
        title_max = self.fmt["title_max_chars"]
        body_min = self.fmt["body_min_chars_per_line"]
        body_max = self.fmt["body_max_chars_per_line"]
        body_lines = self.fmt["body_lines"]

        # 제외할 기사 정보 구성
        exclude_section = ""
        if exclude_articles:
            exclude_details = []
            for article in exclude_articles:
                title = article.get("title", "")
                lines = article.get("lines", [])
                # 첫 번째 불릿만 표시 (내용 파악용)
                content_preview = lines[0][:100] if lines else ""
                exclude_details.append(f"  제목: {title}\n  내용: {content_preview}...")
            
            exclude_list = "\n\n".join(exclude_details)
            exclude_section = f"""

[🚨 필수: 이미 수집한 기사와 중복 절대 금지 🚨]
아래는 이미 선택된 기사입니다:
{exclude_list}

⛔ 절대 금지 사항:
1. 위 제목과 같거나 유사한 주제의 기사 선택 금지
2. 위 기사와 같은 내용, 같은 기관/기업, 같은 통계/발표를 다루는 기사 금지
3. 단순히 표현만 바꾼 같은 소재의 기사 금지
4. 위 기사와 같은 키워드(GDP, 금리, 환율 등)가 중복되는 기사 금지

✅ 반드시 준수:
- '{topic_name}' 카테고리 내에서 **완전히 다른 하위 주제**를 선택할 것
- 예: 거시경제라면 → GDP성장률, 수출입통계, 물가동향, 소비지표, 고용지표, 산업생산, 기업심리지수, 통화정책 등은 모두 **별개의 주제**
- 같은 카테고리 안에서도 세부 영역이 전혀 겹치지 않는 뉴스를 찾을 것
- 다른 기관의 다른 발표, 다른 지표, 다른 산업 영역을 다루는 기사를 선택할 것
"""
        elif exclude_titles:
            # 하위 호환성: 제목만 있는 경우
            exclude_list = "\n".join([f"  - {title}" for title in exclude_titles])
            exclude_section = f"""

[🚨 필수: 이미 수집한 기사와 중복 절대 금지 🚨]
아래는 이미 선택된 기사 제목입니다:
{exclude_list}

⛔ 절대 금지: 위 제목과 같거나 유사한 주제의 기사 선택 금지
✅ 반드시: '{topic_name}' 카테고리 내 **완전히 다른 하위 주제** 선택
"""

        return f"""너는 한국 금융/경제 뉴스 요약 전문가야.
'{topic_name}' 관련 가장 중요한 뉴스 1건을 검색해서 아래 규칙에 맞게 요약해줘.

[검색 우선순위]
1. 최우선: 최근 1주일 이내 뉴스
2. 1주일 이내 없으면: 최근 1달 이내로 확장
3. 1달 이내도 없으면: 최근 2달 이내까지 확장
가능한 한 최신 뉴스를 찾되, 위 범위 내에서 가장 중요한 뉴스를 선택할 것.

[검색 키워드 참고]
{keywords}{exclude_section}

[요약 규칙]
1. 제목: 공백 포함 {title_min}~{title_max}자. 핵심 키워드만 간결하게 압축.
2. 본문: 정확히 {body_lines}개 불릿. 각 불릿은 공백 포함 {body_min}~{body_max}자.
   - 반드시 구체적 숫자를 포함할 것: 금액(조원, 억원, 달러), 비율(%, YoY, QoQ), 날짜, 수량 등.
   - 하나의 불릿에 2~3개 문장을 담아 팩트 위주로 상세히 서술.
   - 모호한 표현("크게 상승", "상당한 규모") 대신 정확한 수치를 사용.
3. 말투: 간결한 문어체 (~함, ~임, ~전망, ~예정, ~예상)
4. 불릿 기호(●, -, *) 사용 금지. 순수 텍스트만.
5. 실제 뉴스 기사를 기반으로 작성. 허구 금지.

[참고 예시 - 이 수준의 구체성과 길이가 필요함]
제목: "LGD, 4년 만에 흑자전환"
불릿1: "2025년 매출 25.8조원(-3% YoY), 영업이익 5,170억원으로 4년 만의 흑자전환"
불릿2: "중국의 LCD 저가 공세 이후 OLED 올인. 3Q25 중소형 OLED M/S 20.3%로 2년 전 동기 대비 2배 이상 상승. OLED 매출 비중 61% 기록"
불릿3: "2026년 영업이익 컨센서스 1.2조원. OLED 및 AX 집중 방침"

[출력 형식]
반드시 아래 JSON 형식으로만 응답. JSON 외 다른 텍스트 절대 금지:
```json
{{
  "title": "구체적 뉴스 제목 ({title_min}~{title_max}자)",
  "lines": [
    "첫 번째 불릿 - 구체적 수치와 팩트 포함 ({body_min}~{body_max}자)",
    "두 번째 불릿 - 구체적 수치와 팩트 포함 ({body_min}~{body_max}자)",
    "세 번째 불릿 - 구체적 수치와 팩트 포함 ({body_min}~{body_max}자)"
  ],
  "source": "원본 기사 출처 언론사명"
}}
```"""

    def _parse_response(self, text: str) -> dict | None:
        """Gemini 응답에서 JSON을 추출하여 파싱"""
        # 1차: 전체 텍스트를 JSON으로 파싱 시도
        cleaned = text.strip()
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError:
            pass

        # 2차: ```json ... ``` 코드블록 안의 JSON 추출
        code_block = re.search(r"```(?:json)?\s*(\{[\s\S]*?\})\s*```", cleaned)
        if code_block:
            try:
                return json.loads(code_block.group(1))
            except json.JSONDecodeError:
                pass

        # 3차: 텍스트 내 첫 번째 JSON 객체 추출
        json_match = re.search(r"\{[\s\S]*\}", cleaned)
        if json_match:
            try:
                return json.loads(json_match.group())
            except json.JSONDecodeError:
                pass

        # 4차: 줄 단위 폴백 파싱
        lines = [line.strip() for line in cleaned.split("\n") if line.strip()]
        if len(lines) >= 4:
            return {
                "title": lines[0].replace("제목:", "").replace('"', "").strip()[:18],
                "lines": [
                    line.replace("-", "").replace("●", "").strip()[:25]
                    for line in lines[1:4]
                ],
                "source": "파싱 실패 - 수동 확인 필요",
            }

        return None

    def collect_article(self, topic_name: str, keywords: str, exclude_titles: list[str] = None, exclude_articles: list[dict] = None) -> dict | None:
        """단일 주제에 대한 뉴스 1건 검색 및 요약"""
        prompt = self._build_prompt(topic_name, keywords, exclude_titles, exclude_articles)

        try:
            # Google Search grounding을 사용한 생성
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt,
                config=GenerateContentConfig(
                    tools=[self.search_tool],
                    temperature=self.temperature,
                ),
            )

            result = self._parse_response(response.text)
            if result:
                return result

            # 파싱 실패 시 1회 재시도 (Search 없이)
            print(f"    응답 파싱 실패, 재시도 중...")
            response = self.client.models.generate_content(
                model=self.model_name,
                contents=prompt,
                config=GenerateContentConfig(
                    temperature=self.temperature,
                ),
            )
            return self._parse_response(response.text)

        except Exception as e:
            print(f"    API 오류: {e}")
            return None

    def collect_all(self) -> list[dict]:
        """config에 정의된 모든 주제에 대해 뉴스 수집"""
        articles = []
        topics = self.config["topics"]
        total = sum(t.get("count", 1) for t in topics)
        idx = 0

        print(f"\n  총 {total}개 기사를 {len(topics)}개 주제에서 수집합니다.\n")

        for topic in topics:
            name = topic["name"]
            keywords = topic["keywords"]
            count = topic.get("count", 1)
            author = topic.get("author", "")  # 저자 정보 가져오기
            
            # 같은 주제에서 여러 기사 수집 시 중복 방지용
            collected_articles = []

            for i in range(count):
                idx += 1
                suffix = f" ({i + 1}/{count})" if count > 1 else ""
                print(f"  [{idx:2d}/{total}] {name}{suffix} 검색 중...", end=" ", flush=True)

                # 이전에 수집한 기사 전체를 제외하고 검색
                article = self.collect_article(
                    name, keywords, 
                    exclude_articles=collected_articles if i > 0 else None
                )

                if article:
                    article["topic"] = name
                    article["author"] = author  # 저자 정보 추가
                    articles.append(article)
                    
                    # 중복 방지를 위해 전체 기사 정보 저장
                    collected_articles.append(article)
                    
                    title_preview = article.get("title", "")[:20]
                    print(f"-> \"{title_preview}\"")
                else:
                    # 실패 시 빈 항목 추가 (수동 수정 가능)
                    articles.append({
                        "topic": name,
                        "title": f"{name} 관련 뉴스 (수동 입력 필요)",
                        "lines": [
                            "해당 주제 뉴스 수집에 실패함",
                            "수동으로 내용을 입력해주세요",
                            "",
                        ],
                        "source": "수집 실패",
                        "author": author,  # 저자 정보 추가
                    })
                    print("-> 실패 (빈 항목 추가됨)")

                # Rate limit 방지 대기
                if idx < total:
                    time.sleep(self.delay)

        success = sum(1 for a in articles if a.get("source") != "수집 실패")
        print(f"\n  수집 완료: 성공 {success}/{total}개")

        return articles
