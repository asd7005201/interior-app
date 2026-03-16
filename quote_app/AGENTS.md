# 견적앱 작업 규칙

이 프로젝트는 관리자 전용 견적 작성 및 운영 앱이다.
목표는 상담 중 빠르고 정확하게 견적을 작성하고, 자재/템플릿/메타 정보를 안정적으로 유지하는 것이다.

## 우선순위
1. 데이터 정합성
2. 기존 관리자 플로우 안정성
3. 템플릿/자재 추천 메타 호환성
4. UI 개선
5. 성능 최적화

## 이 프로젝트에서 중요한 도메인
- Quotes / QuoteItems
- Materials
- MaterialGroups
- MaterialTags
- TemplateCatalog
- TemplateVersions
- TemplateItemsCache

## 반드시 지킬 원칙
- 시트 컬럼명, 저장 키, enum 값은 임의 변경하지 않는다.
- MaterialGroups는 legacy/대표 빠른 매핑으로 취급하고, 추천 품질 기준은 MaterialTags를 우선한다.
- 템플릿 메타(template_type, housing_type, area_band, budget_band, expose_to_prequote, prequote_priority 등)는 가견적앱 연결 계약에 영향을 주므로 수정 전 영향 범위를 먼저 정리한다.
- 관리자 화면(edit, catalog, templates, templateslist, dashboard, materialgroups)의 기존 흐름은 최소 수정 원칙으로 다룬다.
- UI 문구/레이아웃 변경보다 데이터 정합성과 저장 안정성을 우선한다.

## 작업 시작 전 필수 절차
1. 요구사항을 3~7줄로 요약한다.
2. 영향 파일 목록을 먼저 적는다.
3. 시트/캐시/버전 영향 여부를 적는다.
4. 구현 단계를 번호로 나눈다.
5. 수동 QA 항목을 먼저 만든 뒤 수정한다.

## 변경 시 체크포인트
### 자재 관련
- Materials 저장/조회에 영향이 있는가
- MaterialTags 생성/수정/비활성 로직에 영향이 있는가
- tags_summary 자동 생성과 수동 보조 입력 규칙이 깨지지 않는가
- MaterialGroups legacy 매핑이 깨지지 않는가

### 템플릿 관련
- TemplateCatalog 메타 필드가 유지되는가
- TemplateVersions와 snapshot/cache 무효화가 필요한가
- expose_to_prequote / is_featured_prequote / prequote_priority 규칙이 유지되는가
- 가견적 추천에 사용되는 housing_type / area_band / budget_band 정규화가 유지되는가

### 관리자 플로우 관련
- edit 화면에서 견적 작성/수정/저장 흐름이 유지되는가
- dashboard 검색/필터/메모 상태가 깨지지 않는가
- catalog에서 자재 정보와 태그가 함께 저장되는가
- templates / templateslist / materialgroups 이동 동선이 깨지지 않는가

## 금지사항
- 검증 없이 시트 헤더를 변경하지 않는다.
- enum 값을 새 문자열로 임의 추가하지 않는다.
- prequote 연동 메타를 설명 없이 삭제하거나 이름을 바꾸지 않는다.
- 캐시 키/버전 키 로직을 영향 분석 없이 수정하지 않는다.

## 출력 형식
작업 결과는 아래 순서로 정리한다.
1. 변경 요약
2. 수정 파일
3. 데이터 영향
4. QA 체크리스트
5. 남은 리스크