export type BuiltinChatSkillId =
  | "3-statement-model"
  | "audit-xls"
  | "clean-data-xls"
  | "comps-analysis"
  | "credit-review"
  | "dcf-model"
  | "lbo-model"
  | "loan-schedule"
  | "morning"
  | "reconcile"
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
    id: "credit-review",
    source: "builtin",
    slashCommand: "/credit",
    label: "여신 심사 분석",
    shortDescription: "재무제표에서 안정성·수익성·상환능력 지표를 뽑습니다.",
    triggerPhrases: [
      "여신",
      "신용분석",
      "심사",
      "재무분석",
      "부채비율",
      "이자보상배율",
      "상환능력",
      "credit review",
    ],
    guidance:
      "먼저 재무제표의 기간, 단위, 계정 위치를 조회해 확인한다. 3개년 비교표를 만들고 부채비율, 유동비율, 차입금의존도, 이자보상배율, 영업이익률, ROA, EBITDA와 차입금 대비 배수를 셀 참조 수식으로 계산한다. 각 지표는 계산식이 보이도록 두고 판단 근거와 한계를 함께 적는다. 신용등급이나 승인 여부를 단정하지 않는다.",
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
    id: "loan-schedule",
    source: "builtin",
    slashCommand: "/loan",
    label: "상환 스케줄",
    shortDescription: "원리금균등·원금균등 상환표를 수식으로 만듭니다.",
    triggerPhrases: [
      "상환",
      "원리금",
      "원금균등",
      "거치",
      "대출 스케줄",
      "amortization",
      "상환표",
      "이자 계산",
    ],
    guidance:
      "원금, 연이자율, 기간, 상환 주기, 거치기간, 기산일을 먼저 확인하고 하나라도 모르면 묻는다. 가정은 셀에 따로 적고 스케줄은 그 셀을 참조하는 수식으로만 만든다. 원리금균등은 PMT, 이자는 잔액×주기이자율, 원금은 상환액에서 이자를 뺀 값으로 쓰고 마지막 회차에서 잔액이 정확히 0이 되는지 확인한다. 금액은 표시 형식으로만 단위를 맞춘다.",
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
    id: "reconcile",
    source: "builtin",
    slashCommand: "/recon",
    label: "대사·정합성 점검",
    shortDescription: "두 표를 키로 맞춰 차이 나는 건을 찾습니다.",
    triggerPhrases: [
      "대사",
      "정합성",
      "차이 확인",
      "맞춰봐",
      "reconcile",
      "대조",
      "검증",
      "잔액 확인",
    ],
    guidance:
      "두 표의 위치, 키 열, 비교할 금액 열을 먼저 확인한다. 키가 중복인지 먼저 점검하고, 한쪽에만 있는 건과 양쪽에 있으나 금액이 다른 건을 각각 나눠 새 시트에 낸다. 비교는 XLOOKUP 또는 INDEX·MATCH 수식으로 남겨 원본이 바뀌면 따라 움직이게 하고, 금액 비교는 부동소수 오차를 감안해 ROUND로 자리를 맞춘 뒤 비교한다. 차이 합계와 건수를 요약에 적는다.",
    contextProfile: workbookProfile,
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

/**
 * The request with its slash command taken off.
 *
 * `/dcf-model 3년치 전망 만들어줘` selects the skill *and* used to be sent verbatim, so the
 * model read `/dcf-model` as part of the instruction and answered about the command instead
 * of doing the work. The skill arrives through the system prompt; the request should be
 * only what the user actually asked for.
 */
export const stripSlashCommand = (
  input: string,
  skills: readonly ChatSkill[] = CHAT_SKILLS,
): string => {
  const trimmed = input.trimStart()
  const command = trimmed.split(/\s/u)[0]?.toLocaleLowerCase()
  if (command === undefined || !command.startsWith("/")) return input
  if (!skills.some((skill) => skill.slashCommand === command)) return input
  const rest = trimmed.slice(command.length).trim()
  return rest === "" ? input.trim() : rest
}

export const resolvePromptSkill = (
  input: string,
  attached: ChatSkillId | null,
  skills: readonly ChatSkill[] = CHAT_SKILLS,
): ChatSkill | null =>
  explicitSkill(input, skills) ?? skillById(attached, skills) ?? resolveSkill(input, skills)
