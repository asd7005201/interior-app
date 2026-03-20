/**
 * Claude Agent SDK - 에이전트 팀 데모
 * 이 파일은 견적앱 코드베이스를 분석하는 에이전트 팀 예시입니다.
 *
 * 실행 방법:
 *   1. npm install @anthropic-ai/claude-agent-sdk
 *   2. export ANTHROPIC_API_KEY="sk-..."
 *   3. node docs/agent_team_demo.mjs
 */

import { query } from "@anthropic-ai/claude-agent-sdk";

// ─────────────────────────────────────────────
//  에이전트 팀 정의
//  각 팀원은 description, prompt, tools 를 가집니다.
// ─────────────────────────────────────────────
const agents = {

  /**
   * 팀원 1: 코드 리뷰어
   * - 역할: GAS 코드의 품질/안정성 검토
   * - 도구: 파일 읽기 전용 (쓰기 없음)
   */
  "code-reviewer": {
    description: "Google Apps Script 코드의 품질, 데이터 정합성, 캐시 로직을 검토합니다.",
    prompt: `당신은 Google Apps Script 전문 코드 리뷰어입니다.
코드를 검토할 때 다음에 집중하세요:
- 캐시 무효화 로직 누락 여부
- 시트 헤더 컬럼명 하드코딩 위험
- 에러 처리 누락
- 성능 병목 (getValues 배치 vs 단건 호출)
반드시 파일명과 라인 번호를 포함해 구체적으로 지적하세요.`,
    tools: ["Read", "Grep"],
  },

  /**
   * 팀원 2: 데이터 흐름 분석가
   * - 역할: QuoteItems/Materials/Templates 간 데이터 흐름 추적
   * - 도구: 파일 읽기 + 검색
   */
  "data-flow-analyst": {
    description: "견적(Quotes), 자재(Materials), 템플릿(Templates) 사이의 데이터 흐름과 의존성을 분석합니다.",
    prompt: `당신은 이 견적 앱의 데이터 흐름 전문 분석가입니다.
Google Sheets를 DB로 사용하는 앱의 특성을 이해하고 있습니다.
분석 시 집중 포인트:
- 시트 간 ID 참조 무결성
- tags_summary 자동 생성 로직의 입력/출력
- prequote 연동 메타 필드(expose_to_prequote, prequote_priority 등)의 흐름
- 캐시 키가 실제 데이터 변경을 올바르게 반영하는지
결과는 "A → B → C" 형태의 흐름도로 정리하세요.`,
    tools: ["Read", "Grep", "Glob"],
  },

  /**
   * 팀원 3: 문서 작성자
   * - 역할: 분석 결과를 마크다운 문서로 정리
   * - 도구: 파일 읽기 + 쓰기
   */
  "doc-writer": {
    description: "분석 결과를 구조화된 마크다운 문서로 작성합니다.",
    prompt: `당신은 기술 문서 전문 작성자입니다.
다음 형식으로 문서를 작성하세요:
1. ## 요약 (3줄 이내)
2. ## 발견된 이슈 (심각도: 높음/중간/낮음 분류)
3. ## 데이터 흐름
4. ## 개선 권고사항
5. ## QA 체크리스트
간결하고 실행 가능한 내용만 포함하세요.`,
    tools: ["Read", "Write"],
  },
};

// ─────────────────────────────────────────────
//  메인 실행: 오케스트레이터 에이전트
// ─────────────────────────────────────────────
async function runAgentTeam() {
  console.log("🚀 에이전트 팀 시작...\n");

  // 작업 디렉토리: 이 프로젝트 루트
  const projectRoot = new URL("..", import.meta.url).pathname;

  for await (const message of query({
    // 메인 에이전트(오케스트레이터)에게 주는 작업 지시
    prompt: `
견적앱 코드베이스를 다음 순서로 팀과 함께 분석해줘:

1. code-reviewer 에이전트를 호출해서 Code.js와 utils.js의 핵심 로직을 검토해줘.
2. data-flow-analyst 에이전트를 호출해서 Materials → TemplateCatalog → Quotes 데이터 흐름을 추적해줘.
3. doc-writer 에이전트를 호출해서 위 두 분석 결과를 docs/analysis_report.md 파일로 저장해줘.

각 팀원의 결과를 기다린 뒤 다음 단계로 넘어가.
    `.trim(),

    options: {
      cwd: projectRoot,
      // 메인 에이전트가 사용할 수 있는 도구 (Agent 포함 필수!)
      allowedTools: ["Read", "Glob", "Grep", "Agent"],
      permissionMode: "acceptEdits",
      maxTurns: 30,
      agents,  // 위에서 정의한 팀원들 등록
      systemPrompt: `당신은 팀 오케스트레이터입니다.
각 전문 에이전트를 적절히 활용해 작업을 분담하고 결과를 취합하세요.
에이전트 호출 전에 반드시 어떤 에이전트를 왜 호출하는지 한 줄로 설명하세요.`,
    },
  })) {
    // 메시지 타입별 처리
    if (message.type === "system" && message.subtype === "init") {
      console.log(`📌 세션 ID: ${message.session_id ?? "(없음)"}\n`);
    } else if (message.type === "result") {
      console.log("\n✅ 최종 결과:\n");
      console.log(message.result);
    } else if (message.type === "assistant") {
      // 어시스턴트 메시지에서 텍스트만 출력
      for (const block of message.content ?? []) {
        if (block.type === "text") {
          process.stdout.write(block.text);
        }
      }
    }
  }

  console.log("\n\n🏁 에이전트 팀 작업 완료.");
}

runAgentTeam().catch(console.error);
