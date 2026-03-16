# 가견적앱 작업 규칙

이 프로젝트는 고객용 가견적 및 상담 리드 수집 앱이다.
목표는 고객 이탈을 줄이면서 상담과 실견적으로 이어질 수 있는 충분한 정보를 수집하는 것이다.

## 우선순위
1. 제출 완료율
2. 질문 흐름의 자연스러움
3. 모바일 UX
4. 결과 화면의 신뢰감
5. 견적앱과의 데이터 연결 안정성

## 이 프로젝트에서 중요한 도메인
- public flow: landing / survey / result / complete
- survey builder
- survey profiles
- quick bundles
- estimate rules
- sync
- versions
- requests / request detail

## 반드시 지킬 원칙
- single choice는 자동 전진 UX를 유지한다.
- multi choice는 자동 제출/자동 전진을 강제하지 않는다.
- visible_if_expr 또는 유사 분기 조건을 수정할 때는 숨겨진 질문의 answer pruning 영향을 함께 본다.
- 질문 수를 늘릴 때는 반드시 이유를 적는다.
- 고객이 이해하기 어려운 운영 용어는 화면에 직접 노출하지 않는다.
- 결과 화면의 estimate range, flags, recommendation은 견적앱 계약과 어긋나지 않도록 유지한다.

## 작업 시작 전 필수 절차
1. 요청사항을 요약한다.
2. public/admin 중 어느 영역인지 구분한다.
3. 영향 파일을 적는다.
4. 설문 분기/결과/동기화/버전 영향 여부를 적는다.
5. 구현 단계와 QA 항목을 만든 뒤 수정한다.

## 변경 시 체크포인트
### 공개 설문 관련
- landing → survey → result → complete 흐름이 유지되는가
- single choice 자동 전진이 유지되는가
- multi choice는 다음 버튼 흐름이 유지되는가
- draft 저장/복원에 영향이 있는가
- file upload / text input / chip 입력이 유지되는가

### 설문 빌더 관련
- question / option / logic / tag rules 저장이 유지되는가
- profile_scope / project_scope 영향이 반영되는가
- quick bundle preset_answers_json / preset_tags_json 규칙이 유지되는가
- 질문 삭제/비활성 처리 시 공개 설문 표시 순서가 깨지지 않는가

### 운영 관련
- requests 목록/상세 저장이 유지되는가
- sync 로그/상태 표시에 영향이 있는가
- versions의 published / draft / restore 흐름이 유지되는가
- publish 전 검토 없이 draft 구조를 바로 바꾸지 않는가

## 금지사항
- 결과 필드명을 설명 없이 변경하지 않는다.
- 견적앱 연동용 코드/enum을 임의 rename 하지 않는다.
- 분기 로직을 수정하면서 QA 없이 배포하지 않는다.
- 질문 수를 늘리는 방향을 기본값으로 잡지 않는다.

## 출력 형식
작업 결과는 아래 순서로 정리한다.
1. 변경 요약
2. 수정 파일
3. UX 영향
4. 연동 영향
5. QA 체크리스트
6. 남은 리스크