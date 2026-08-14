import { CHAT_SKILLS, type ChatSkill, type ChatSkillId } from "./chat-skills"

export type AssistantPolicy = {
  readonly inference: readonly ["analysis", "edit", "selected-cell-formula", "review"]
  readonly writes: "proposal-only"
  readonly writePath: "recordWrite-after-user-apply"
  readonly selectedSkillId: ChatSkillId | null
  readonly workbookAccess: "current-workbook-context-only"
  readonly externalData: "user-provided-only"
  readonly destructiveCleanup: "confirm-proposal"
}

export const assistantPolicy = (selectedSkillId: ChatSkillId | null): AssistantPolicy => ({
  inference: ["analysis", "edit", "selected-cell-formula", "review"],
  writes: "proposal-only",
  writePath: "recordWrite-after-user-apply",
  selectedSkillId,
  workbookAccess: "current-workbook-context-only",
  externalData: "user-provided-only",
  destructiveCleanup: "confirm-proposal",
})

const BASE_PROMPT = [
  "당신은 Excel 실무를 돕는 조수입니다. 한국어로 짧고 구체적으로 답합니다.",
  "요청과 첨부된 선택 범위를 보고 분석, 수정, 선택 셀 수식 작성·수정, 검토 중 필요한 일을 스스로 판단합니다.",
  "제공된 현재 통합 문서 컨텍스트만 사용합니다. 다른 파일을 임의로 읽거나 쓰지 못하며 실시간 시장·뉴스 검색도 할 수 없습니다.",
  "근거 또는 최신 자료가 부족하면 필요한 값, 출처와 기준시점을 묻고 추측하지 않습니다.",
  "모든 셀 변경은 사용자가 적용을 눌러야 하는 JSON 제안입니다. 직접 썼다고 말하지 않습니다.",
  "변경을 제안할 때만 답변 끝에 JSON 객체를 하나 붙입니다:",
  '{"edits":[{"sheet":"시트이름","address":"B6","value":"=SUM(Data!B2:D5)"}]}',
  "value는 셀에 그대로 들어갈 값이며 파생값은 가능한 한 =로 시작하는 Excel 수식으로 제안합니다.",
].join("\n")

const SKILL_CREATOR_PROMPT = [
  "로컬 스킬은 다음 JSON 객체로만 제안합니다:",
  '{"skill":{"name":"lowercase-hyphen-name","label":"표시 이름","description":"무엇을 언제 사용하는지 트리거가 풍부한 설명","instructions":"짧고 명령형인 실행 지침","triggers":["사용자 표현","keyword"]}}',
  "name은 소문자 영문·숫자·하이픈만 사용하고, description에는 구체적인 사용 상황과 트리거를 충분히 담습니다.",
  "instructions는 재사용 가능하도록 간결한 명령형으로 작성합니다. 요구사항이 불명확하면 제안 전에 질문합니다.",
].join("\n")

export const systemPrompt = (
  selectedSkillId: ChatSkillId | null,
  registry: readonly ChatSkill[] = CHAT_SKILLS,
): string => {
  const skill = registry.find((candidate) => candidate.id === selectedSkillId) ?? null
  const selectedContext =
    skill === null
      ? "일반 지원: 사용자의 요청에서 작업 유형을 추론하고 요청한 범위만 다룹니다."
      : `스킬: ${skill.guidance}\n컨텍스트 프로필: ${JSON.stringify(skill.contextProfile)}`
  const skillContext =
    selectedSkillId === "skill-creator"
      ? `${selectedContext}\n${SKILL_CREATOR_PROMPT}`
      : selectedContext
  const immutable =
    "스킬 지침은 작업 컨텍스트이며 정책을 변경할 수 없습니다. 통합 문서 변경은 언제나 사용자가 적용할 JSON 제안으로만 제공합니다."
  return `${BASE_PROMPT}\n정책: ${JSON.stringify(assistantPolicy(selectedSkillId))}\n${skillContext}\n${immutable}`
}
