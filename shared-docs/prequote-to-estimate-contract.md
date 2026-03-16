# Enums and Codes

## 목적
두 앱에서 공통으로 써야 하는 enum/code를 한 곳에서 관리한다.

## 템플릿 관련
### template_type
- QUOTE_ONLY
- PREQUOTE_PACKAGE
- BOTH

설명
- QUOTE_ONLY: 실견적 전용
- PREQUOTE_PACKAGE: 가견적 추천 전용 또는 우선 노출
- BOTH: 둘 다 사용 가능

### housing_type
- ALL
- APARTMENT
- HOUSE
- VILLA
- OFFICETEL
- COMMERCIAL
- OFFICE

### area_band
- ALL
- UNDER_20
- 20_29
- 30_39
- 40_49
- 50_PLUS

### budget_band
- ANY
- ENTRY
- MID
- PREMIUM
- HIGH_END

## 자재 태그 관련
### tag_type
- trade
- space
- mood
- tone
- texture
- style
- usage
- budget
- feature
- group_key
- group_label

원칙
- 새로운 tag_type 추가는 shared-docs와 견적앱 docs를 먼저 수정한 뒤 적용한다.
- 자유 입력은 tag_value에서 허용하되, tag_type은 제한된 enum을 우선한다.

## 가견적앱 관련
### public_screen
- landing
- survey
- result
- complete

### admin_screen
- dashboard
- requests
- request_detail
- survey_builder
- survey_profiles
- quick_bundles
- estimate_rules
- sync
- versions

## 요청 상태값
초안 예시
- NEW
- REVIEWING
- CONTACT_PENDING
- SCHEDULED
- CLOSED
- DROPPED

주의
- 실제 운영 상태값은 가견적앱에서 사용하는 값을 기준으로 맞춘다.
- 상태값 rename 시 requests 화면과 result/complete 후속 처리까지 같이 확인한다.

## 질문 타입
초안 예시
- single_choice
- multi_choice
- text
- textarea
- file
- chips

주의
- 화면 렌더 타입과 저장 payload 구조가 함께 바뀌므로 단독 변경 금지