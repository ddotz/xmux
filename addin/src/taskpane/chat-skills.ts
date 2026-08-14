export type BuiltinChatSkillId =
  | "3-statement-model"
  | "audit-xls"
  | "clean-data-xls"
  | "comps-analysis"
  | "dcf-model"
  | "lbo-model"
  | "morning"
  | "skill-creator"

export type LocalChatSkillId = `local:${string}`
export type ChatSkillId = BuiltinChatSkillId | LocalChatSkillId

export type SkillContextProfile = {
  readonly scope: "selection-or-workbook" | "financial-model" | "user-supplied-current-data"
  readonly externalWorkbooks: "unavailable"
  readonly currentData: "not-live"
}

export type ChatSkill = {
  readonly id: ChatSkillId
  readonly source: "builtin" | "local"
  readonly slashCommand: string
  readonly label: string
  readonly shortDescription: string
  readonly triggerPhrases: readonly string[]
  readonly guidance: string
  readonly contextProfile: SkillContextProfile
}

const workbookProfile: SkillContextProfile = {
  scope: "selection-or-workbook",
  externalWorkbooks: "unavailable",
  currentData: "not-live",
}
const modelProfile: SkillContextProfile = { ...workbookProfile, scope: "financial-model" }
const morningProfile: SkillContextProfile = {
  ...workbookProfile,
  scope: "user-supplied-current-data",
}

export const CHAT_SKILLS: readonly ChatSkill[] = [
  {
    id: "3-statement-model",
    source: "builtin",
    slashCommand: "/3-statement",
    label: "3개 재무제표 모델",
    shortDescription: "손익·재무상태·현금흐름표를 연결하고 검증합니다.",
    triggerPhrases: ["3 statement", "3-statement", "삼표 모델", "3개 재무제표", "재무제표 모델"],
    guidance:
      "템플릿 구조와 기간을 먼저 확인하고 손익, 재무상태, 현금흐름 순서로 제안한다. 전망치는 수식으로 연결하며 BS 균형과 현금 일치를 단계마다 검증한다.",
    contextProfile: modelProfile,
  },
  {
    id: "audit-xls",
    source: "builtin",
    slashCommand: "/audit",
    label: "엑셀 감사",
    shortDescription: "수식 오류와 모델 무결성 문제를 찾습니다.",
    triggerPhrases: ["audit", "감사", "수식 오류", "모델 검토", "model review", "qa", "안 맞아"],
    guidance:
      "범위를 확인하고 오류값, 불일치 수식, 하드코드, 범위 누락, 단위와 연결을 점검한다. 먼저 심각도별 발견사항을 보고하며 요청 전에는 수정하지 않는다.",
    contextProfile: modelProfile,
  },
  {
    id: "clean-data-xls",
    source: "builtin",
    slashCommand: "/clean",
    label: "데이터 정리",
    shortDescription: "공백·형식·중복·혼합 데이터 문제를 정리합니다.",
    triggerPhrases: [
      "clean",
      "데이터 정리",
      "정리해",
      "중복",
      "dedupe",
      "정규화",
      "표준화",
      "normalize",
      "형식 통일",
    ],
    guidance:
      "열별 문제와 건수를 먼저 요약하고 투명한 도우미 수식을 우선 제안한다. 덮어쓰기, 중복 삭제 등 파괴적 정리는 반드시 확인 가능한 제안으로만 제공한다.",
    contextProfile: workbookProfile,
  },
  {
    id: "comps-analysis",
    source: "builtin",
    slashCommand: "/comps",
    label: "비교기업 분석",
    shortDescription: "기업 지표와 밸류에이션 배수를 비교합니다.",
    triggerPhrases: ["comps", "comparable", "비교기업", "유사기업", "peer analysis", "멀티플 비교"],
    guidance:
      "비교 목적, 피어, 기간과 단위를 먼저 확정한다. 파생 지표와 분위수는 셀 참조 수식으로 만들고 원천 데이터는 사용자가 제공한 값과 출처로 제한한다.",
    contextProfile: modelProfile,
  },
  {
    id: "dcf-model",
    source: "builtin",
    slashCommand: "/dcf",
    label: "DCF 모델",
    shortDescription: "현금흐름 할인 가치평가 구조를 설계합니다.",
    triggerPhrases: [
      "dcf",
      "discounted cash flow",
      "할인현금흐름",
      "현금흐름 할인",
      "내재가치",
      "wacc",
      "터미널 가치",
    ],
    guidance:
      "사용자 제공 원천값을 확인한 뒤 매출, FCF, WACC, 터미널 가치, 가치 브리지 순으로 제안한다. 전망과 민감도는 수식으로 만들고 기준 가정을 명시한다.",
    contextProfile: modelProfile,
  },
  {
    id: "lbo-model",
    source: "builtin",
    slashCommand: "/lbo",
    label: "LBO 모델",
    shortDescription: "인수금융·부채상환·수익률 모델을 구성합니다.",
    triggerPhrases: ["lbo", "leveraged buyout", "차입매수", "인수금융", "cash sweep", "moic"],
    guidance:
      "현재 통합 문서의 템플릿을 우선 사용하고 없으면 필요한 구조를 확인한다. Sources & Uses, 운영, 부채, 수익률 순으로 수식 제안하며 잔액과 현금 스윕을 검증한다.",
    contextProfile: modelProfile,
  },
  {
    id: "morning",
    source: "builtin",
    slashCommand: "/morning",
    label: "모닝 노트",
    shortDescription: "제공된 최신 자료로 짧은 아침 브리핑을 만듭니다.",
    triggerPhrases: [
      "morning note",
      "morning meeting",
      "모닝 노트",
      "아침 브리핑",
      "오버나이트",
      "overnight",
    ],
    guidance:
      "2분 내 읽을 핵심 의견, 주요 전개, 오늘 일정, 아이디어와 위험으로 압축한다. 실시간 뉴스나 시세를 조회할 수 없으므로 현재 자료가 없으면 출처와 기준시점을 요청한다.",
    contextProfile: morningProfile,
  },
  {
    id: "skill-creator",
    source: "builtin",
    slashCommand: "/skill-creator",
    label: "스킬 만들기",
    shortDescription: "반복 업무 지침을 내 로컬 스킬로 만듭니다.",
    triggerPhrases: ["스킬 만들기", "스킬 생성", "나만의 스킬", "skill creator", "create skill"],
    guidance:
      "만들 스킬의 구체적인 사용 예와 트리거를 확인한다. 정보가 부족하면 가장 중요한 질문부터 짧게 묻는다. 확정되면 lowercase hyphen 이름, 트리거가 드러나는 설명, 간결한 명령형 지침으로 로컬 스킬 하나만 제안한다.",
    contextProfile: workbookProfile,
  },
]

export const skillById = (
  id: ChatSkillId | null,
  skills: readonly ChatSkill[] = CHAT_SKILLS,
): ChatSkill | null => skills.find((skill) => skill.id === id) ?? null

const explicitSkill = (input: string, skills: readonly ChatSkill[]): ChatSkill | null => {
  const command = input.trimStart().split(/\s/u)[0]?.toLocaleLowerCase()
  if (command === undefined || !command.startsWith("/")) return null
  return skills.find((skill) => skill.slashCommand === command) ?? null
}

export const resolveSkill = (
  input: string,
  skills: readonly ChatSkill[] = CHAT_SKILLS,
): ChatSkill | null => {
  const explicit = explicitSkill(input, skills)
  if (explicit !== null) return explicit
  const normalized = input.toLocaleLowerCase()
  return (
    skills.find((skill) =>
      skill.triggerPhrases.some((phrase) => normalized.includes(phrase.toLocaleLowerCase())),
    ) ?? null
  )
}

export const resolvePromptSkill = (
  input: string,
  attached: ChatSkillId | null,
  skills: readonly ChatSkill[] = CHAT_SKILLS,
): ChatSkill | null =>
  explicitSkill(input, skills) ?? skillById(attached, skills) ?? resolveSkill(input, skills)
