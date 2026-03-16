# Material Taxonomy

## 목적
자재 추천과 템플릿 추천에서 사용하는 분류 체계를 정리한다.

## 핵심 구성요소

### Materials
실제 자재 마스터.
가격, 규격, 이미지, trade_code, price_band, expose_to_prequote, recommendation_score_base 등을 가진다.

### MaterialGroups
legacy/대표 빠른 매핑용.
관리 편의와 대표 분류 용도로 사용한다.
정교한 추천 정확도의 기준으로 단독 사용하지 않는다.

### MaterialTags
실제 추천 품질의 핵심.
한 자재에 여러 태그를 붙일 수 있다.

## tag_type 의미
- trade: 공정 분류
- space: 공간 분류
- mood: 분위기
- tone: 색감/톤
- texture: 질감
- style: 스타일
- usage: 사용 목적
- budget: 예산대
- feature: 기능 특성
- group_key: 그룹 연결용 키
- group_label: 그룹 표시 라벨

## 운영 원칙
- 추천은 MaterialTags를 기준으로 한다.
- MaterialGroups는 관리 편의와 빠른 대표 매핑 용도다.
- tags_summary는 검색/요약 보조값이며, 원본 태그 테이블을 대체하지 않는다.
- derived tag와 manual tag를 구분할 수 있으면 구분한다.
- 비활성 태그는 과거 데이터 참고용으로만 남기고 추천 점수에는 반영하지 않는다.

## 점수 원칙
초안
- trade 가중치 가장 높음
- space / mood / tone / style 순으로 반영
- budget / feature / usage는 보정값으로 사용
- recommendation_score_base 로 개별 자재 기본점수를 조정

## 변경 원칙
- tag_type 추가는 공유 문서와 견적앱 docs를 동시에 수정한다.
- UI에서 허용하는 자유 입력은 저장 전 normalize 규칙을 거친다.