# Windows LTSC 추가 기능 취득 문제 — 케이스별 결과와 대응 플레이북

작성 시점: 2026-08-28. 대상: Office LTSC 2024 2408 (16.0.17932.20842), Windows 10/11,
망분리 사내 PC. 이 문서는 이후 세션(에이전트 교체 포함)이 같은 실험을 반복하지 않도록
**시도한 모든 케이스의 결과·증거·판정**과, 파일럿 결과별 대응, 그리고 WEF(웹 추가 기능)
바깥의 대체 구현 채널 검토를 남긴다. 측정 원본은 `FINDINGS.md` F8·F9.

## 문제 정의

Developer 레지스트리(`HKCU\...\Wef\Developer`)로 사이드로드한 웹 추가 기능을 콜드 WEF
프로필에서 처음 추가하면 로드가 실패하고, 하단 오류를 클릭해 열리는 팝업은
**"보안 센터에서 추가 기능 카탈로그를 추가하거나 활성화하세요."** 를 표시한다. 같은 세션에서
다시 추가하면 성공한다. 첫 추가 실패 동안 로컬 서비스(`service.log`)에는 요청이 전혀 없다
(SourceLocation 이전 단계 실패).

## 증거 원본

- 캡처 런 2회(제품 매니페스트 / 최소 매니페스트): `firstrun-20260827-143308`,
  `firstrun-20260827-193239` — 조작자 PC 보관(다운로드 폴더). A/B0/B/C/D 스냅샷과
  `diff-*` 파일 포함. 핵심 수치는 아래와 F9에 옮겨 적었으므로 원본 없이도 판단 가능.
- 진단 도구: `addin/scripts/diagnose-wef-firstrun.ps1`(스냅샷/diff/reset),
  `run-wef-investigation.ps1`, `analyze-wef-run.ps1`(증거 기반 자동 판정).
- Trusted Catalog 파일럿 키트: 조작자 PC `땡땡엑셀-TrustedCatalog-파일럿`
  (setup / verify-reboot / status / restore / collect / admin).

## 케이스별 결과

| # | 케이스 | 결과 | 근거 | 판정 |
|---|---|---|---|---|
| 1 | 매니페스트 축소 (권한·도메인·GetStarted 제거한 v0-minimal) | 동일 실패 | 두 런의 A→B0 동형 | 매니페스트 내용은 원인 아님. 변형 추가 실험 불필요 |
| 2 | 오류 화면 열고 닫은 뒤 **같은 세션**에서 재추가 | 성공 | 두 런의 D에서 `GET /index.html -> 200`, `Activated App` | 유일하게 재현된 복구 경로. 현행 위저드(`initialize.ps1`)의 근거 |
| 3 | 첫 추가 실패 후 오류 화면 **없이** Excel 재시작 → 재추가 | **실패** | 조작자 실험(키트 미캡처) | 첫 추가가 남긴 디스크 상태는 충분조건 아님 |
| 4 | 성공 상태의 레지스트리+WEF 캐시를 백업 → 클린 프로필에 재생(pre-seed) → 추가 | **실패** | 조작자 실험 | **pre-seed는 원리적으로 불가.** B0→B 델타 전체가 `UserIdentityCache` 7개 값뿐이고 그 `ExcelCacheExpire`는 기록 시점에 이미 만료(−4초). 성공 조건은 디스크에 없음. 재시도 금지 |
| 5 | 팝업 닫기(B→C)가 뭔가를 기록하는지 | 레지스트리 변화 0 | `diff-B-C-registry.txt` = 차이 없음 | 팝업 자체는 무의미. **대화상자가 열리는 순간** 프로세스 내 카탈로그/신원 서브시스템이 초기화되는 것이 인과 |
| 6 | 워밍업 위저드를 설치기에 내장 (현행 배포) | 세션 내 동작 확인 | 설치 현장 보고 | 임시 완화책. 리스크: ① 사용자 5단계 수동 절차 ② provider `Entitlements`가 취득+24h FILETIME — **하루 뒤 재발 미검증** ③ `WefCacheId` 리셋(Office 업데이트/복구) 시 재수행 |
| 7 | Trusted Catalog(공유 폴더) 채널 | **미판정** — 파일럿 키트 준비 완료, 실행 대기 | 실패 팝업이 카탈로그를 직접 요구 | **최우선 실험.** 성공 기준: 클린 프로필에서 첫 추가 즉시 활성화 + 새 프로세스·재부팅 후 지속 |

## 파일럿 결과별 대응 (판정 트리)

**케이스 7 성공 (첫 추가 성공 + 재부팅 지속)** 이면:

1. 배포 기본 채널을 Trusted Catalog로 전환. 설치 메뉴 6번(채널 전환)이 이미 구현되어
   있으므로, `install-windows-local.ps1`의 기본값을 카탈로그 채널로 뒤집고 워밍업
   위저드(`initialize.ps1`)와 메뉴 5번을 **삭제**한다 (fallback으로도 남기지 않는다 —
   근원 해결 뒤 증상 패치를 유지하는 것은 부채다).
2. 카탈로그 위치는 사내 파일서버 UNC(관리자 불필요)를 1순위, 로컬 공유 생성(관리자 1회)을
   2순위로 문서화한다. `docs/INSTALL.md` 갱신.
3. `FINDINGS.md` F8/F9에 파일럿 결과를 추가하고 이 문서의 케이스 7을 판정으로 채운다.

**케이스 7 실패** 면 웹 추가 기능 사이드로드는 이 환경에서 두 채널 모두 막힌 것이므로:

1. 사내 IT에 GPO/정책 배포(신뢰할 수 있는 카탈로그 중앙 등록) 가능 여부를 먼저 타진 —
   조직 관리 채널은 사용자 프로필 상태와 무관하게 동작할 수 있다.
2. 불가하면 아래 "WEF 외 대체 채널"의 권고 순서로 네이티브 채널 설계를 시작한다.
3. 그동안 위저드는 유일한 동작 경로이므로 유지하되, 24h 재발 여부(케이스 6 리스크 ②)를
   먼저 재검한다 — 재발한다면 위저드는 "한 번만"이 아니라 상시 절차가 되므로 대체 채널
   결정을 앞당겨야 한다.

## 미검증 항목 (다음 Windows 세션 체크리스트)

- [ ] 워밍업 후 **24시간 뒤** 추가 기능이 그냥 열리는지 (`Entitlements` TTL 재발 검증)
- [ ] Trusted Catalog 파일럿 `setup` → 첫 추가 판정
- [ ] 파일럿 `verify-reboot` → 재부팅 지속성 판정
- [ ] 사내 UNC 변형(파일서버 공유에 카탈로그 배치, 관리자 불필요) 동일 검증
- [ ] 파일럿 `collect` 산출물 회수 → 이 문서와 FINDINGS 갱신

## WEF 외 대체 채널 검토 — "같은 제품을 다른 추가 기능 방식으로 만들 수 있나"

전제가 되는 코드 사실: 이 저장소는 `Excel.run`을 `taskpane/main.ts` 한 곳에서만 호출하고,
나머지 전부(excel/·formula/·ai/·taskpane 뷰)는 `office-shapes.ts`의
`InspectContext`/`OperateContext` 계약에 의존한다. `excel/eval-context.ts`는 이 계약을
**Office 없이** 구현한 626 LOC 스탠드인이고 평가 하네스가 이미 그 위에서 무수정으로 돈다.
즉 **Excel I/O 백엔드는 설계상 교체 가능**하며, 대체 채널의 비용은 "브리지 1개 + 호스팅
셸"로 수렴한다. UI(pane 번들)는 어느 채널에서든 WebView2로 그대로 렌더링한다.

| 채널 | 설치 권한 | UI 재사용 | 코드 재사용 | 정책 리스크 | 판정 |
|---|---|---|---|---|---|
| COM 추가 기능 (VSTO/C#) | HKCU 등록 가능 (`HKCU\Software\Microsoft\Office\Excel\Addins`), VSTO 런타임은 Office에 동봉 | Custom Task Pane + WebView2에 기존 번들 로드 | InspectContext/OperateContext를 COM 오브젝트 모델로 구현한 브리지 필요. 나머지 TS 전량 재사용 | 서명 요구 가능성(은행), regasm 없는 per-user 등록 설계 필요 | **가능.** WEF 전면 실패 시 1순위 |
| XLL (Excel-DNA) | 관리자 불필요 — `HKCU\...\Excel\Options` OPEN 등록 또는 사용자가 추가 기능 대화상자에서 .xll 선택. Excel-DNA는 레지스트리 없는 in-process COM(CTP 포함) 지원 | 동일 (ExcelDna CTP + WebView2) | COM과 동일한 브리지 | Trust Center의 "서명된 추가 기능만" 정책, 엔드포인트 보안의 unsigned DLL 민감도 | **가능.** 단일 .xll 파일 배포라 가장 가볍다. COM과 사실상 동급 — 서명 체계만 확보되면 이쪽 |
| VBA 추가 기능 (.xlam) | 불필요 | **불가** — WebView2를 UserForm에 올릴 공식 경로 없음(IE WebBrowser는 사망) | 로직 전부 재작성 필요 | 매크로 보안 정책 | **부적합.** UI 없는 미니 기능 launcher 정도만 |
| Office Scripts | — | — | — | M365 전용, LTSC 미지원 | **불가** |
| 추가 기능 없이 외부 자동화 (독립 창 + COM Automation) | exe 실행만 (이미 node.exe 서비스 전례 있음) | pane을 일반 데스크톱 창(WebView2)에 호스팅, Excel 옆에 도킹 흉내 | 브리지를 COM Automation(`GetActiveObject`)으로 구현. Office 통합 표면 자체가 0 | 작업 창이 Excel 내부가 아님(UX 저하). 셀 편집 중 COM 호출 차단은 Office.js와 동일 | **가능하나 UX 타협.** F7의 Windows 컴패니언(UIA+키훅)과 한 몸으로 만들면 시너지 — 최후이자 최속 프로토타입 경로 |

주의: 네이티브 채널(COM/XLL)도 셀 편집기 내부 키 입력은 못 본다(F1–F4와 동일 한계).
다만 in-process면 `SetWindowsHookEx` 로컬 훅이 쉬워져 Windows 컴패니언 기능(F2/Tab)을
같은 바이너리에 흡수할 수 있다 — 채널 전환이 컴패니언 과제를 겸하는 셈.

### 권고 — 추가 기능으로 만든다면 **Excel-DNA XLL + 호스트 어댑터**

전제: 케이스 7 파일럿 판정이 항상 먼저다. 성공하면 코드 변경 0이므로 아래는 전부 불필요하다.
파일럿이 실패했을 때의 목표 설계만 여기 고정한다.

**왜 VSTO/COM이 아니라 XLL인가**

- 배포 표면이 가장 작다: `.xll` 파일 1개 + HKCU `Excel\Options` OPEN 등록. 관리자 불필요,
  런타임 설치 불필요(.NET Framework 4.8은 Windows in-box). VSTO는 ClickOnce + VSTO 런타임 +
  서명된 매니페스트까지 따라온다.
- Excel-DNA는 CTP를 얻기 위해 in-process COM 추가 기능을 런타임에 등록해 `ICTPFactory`를
  받는다 — **COM 레지스트리 등록이 아예 없다.**
- 측정 근거(대상 PC A 스냅샷): `HKCU/HKLM\SOFTWARE\Policies\Microsoft\Office` 둘 다 **없음**,
  `Excel\Security`에는 `DisableDDEServerLaunch=1` 하나뿐이고 `RequireAddinSig`는 없다.
  즉 **미서명 XLL을 정책이 막고 있지 않다.** 설치기가 이미 `Unblock-File`로 MOTW를 지우므로
  인터넷 유래 XLL 차단에도 걸리지 않는다. (서명은 "있으면 좋음"이지 선행 조건이 아니다.)

**아키텍처**

```
.xll (Excel-DNA)
 ├─ CTP ← WinForms UserControl ← WebView2 ← 기존 dist/ 번들 그대로
 ├─ 자산 서빙: SetVirtualHostNameToFolderMapping (가상 https 오리진 → 로컬 폴더)
 └─ Excel I/O: AddHostObjectToScript → .NET 브리지 → QueueAsMacro → COM
```

- 자산을 가상 호스트로 매핑하면 **로컬 HTTPS 서비스·자체 CA·포트 3927·로그온 자동 시작이 통째로
  사라진다.** 지금 설치 복잡도의 대부분이 그것이다. 동일 오리진 의미라 fetch/CSP도 그대로 산다.
- COM 호출은 Excel 매크로 컨텍스트에서만 안전하므로 `ExcelAsyncUtil.QueueAsMacro`로 태워야 한다.
- in-process라 `SetWindowsHookEx` 로컬 후킹 + UIA가 같은 바이너리에서 된다 → **F7의 "Windows
  컴패니언 별도 제작" 과제가 흡수된다.**

**비용 (측정값)**

Office.js 결합면은 `taskpane/main.ts`의 약 13개 지점 + `office-shapes.ts`의 타입 패리티
assert 6개 + 나머지 4개 파일의 부수적 참조 1개씩이 전부다. **pane·수식·AI·뷰 코드는
손대지 않는다.** 실제 작업량은 TS가 아니라 C# 브리지(두 인터페이스 구현) + XLL 셸이다.
`eval-context.ts`(626 LOC)가 계약의 참조 구현이다.

**리스크 2개 — 반드시 먼저 깨라**

1. **WebView2 × CTP 초기화** (Excel-DNA #682, WebView2Feedback #405): CTP 안에서
   `EnsureCoreWebView2Async`가 실패한 사례가 보고돼 있다. 원인은 STA/메시지 펌프 재진입이고
   회피 패턴(핸들 생성·가시화 이후 초기화, 펌프 블로킹 금지, `SynchronizationContext.Post`)이
   알려져 있다. **VSTO를 골라도 동일한 리스크다(CTP 메커니즘이 같음) — 채널 선택의 변수가
   아니라 선행 검증 대상이다.**
2. **macOS 지원 상실**: XLL/COM은 Excel for Mac에 존재하지 않는다. 현재 제품은 Mac Excel
   16.111에서 실동작 검증된 상태고 개발기도 Mac이다. 따라서 XLL은 "교체"가 아니라
   **두 번째 호스트 어댑터**로 붙여야 한다 — Office.js 어댑터(Mac·개발·평가 하네스),
   host-object 어댑터(Windows 배포). 시임이 이미 있으니 공존 비용은 어댑터 1개다.

**스파이크 게이트 (이거 통과 전에 본 구현 금지)**

1. 빈 Excel-DNA XLL이 CTP를 띄우고, 그 안 WebView2가 가상 호스트 매핑으로 현재 `dist/index.html`을
   렌더링한다.
2. host object 왕복 1회: pane JS → `.NET` → `QueueAsMacro` → COM으로 실제 셀 값을 읽어 반환.
3. 미서명 `.xll`을 대상 PC에서 경고 없이(또는 1회 승인으로) 로드.

3개 다 통과해야 브리지 전면 구현으로 간다. 1이 끝내 막히면 CTP 대신 Excel 창에 소유된
WinForms Form(도킹 후내) 또는 독립 창 + Automation으로 후퇴한다.
