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

1. 배포 기본 채널을 Trusted Catalog로 전환. 구현은 **`windows/trusted-catalog-pilot` 브랜치**에
   살아 있다(`catalog-windows-local.ps1` + 설치/제거/메뉴 6번/패키징/테스트). 파일럿이
   통과하면 그 브랜치를 Windows에서 검증한 뒤 main으로 병합하고, `install-windows-local.ps1`의
   기본값을 카탈로그 채널로 뒤집은 다음 워밍 위저드(`initialize.ps1`)와 메뉴 5번을
   **삭제**한다 (fallback으로도 남기지 않는다 — 근원 해결 뒤 증상 패치를 유지하는 것은 부채다).
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

## 진행 방향 (2026-08-28 결정)

회사 PC 검증이 막혀 있으므로, **Windows 없이 맥에서 할 수 있는 일을 먼저 하고 채널 선택은
나중에 한다.** 그러려면 채널이 교체 가능해야 하고, 그게 호스트 어댑터 분리다.

### 브랜치 분리 상태 (2026-08-28 정리 후)

| 브랜치 | 내용 | 검증 |
|---|---|---|
| `main` | 진단 킷 + Developer 워밍 완화 + 근거 문서 + **`ExcelHost` 포트(단계 1·1.5·2)** | 857 tests |
| `windows/trusted-catalog-pilot` | main + Trusted Catalog 채널 delta 7파일 (`catalog-windows-local.ps1`, 설치/제거/메뉴 6번/패키징/테스트/INSTALL) | 864 tests |

`adapter/excel-host`는 main에 흡수되고 삭제됐다 — 단계 1·2는 파일럿 결과와 무관하게
유지하기로 이미 결정한 변경이라 파생 브랜치에 남길 이유가 없었다. 단계 3(XLL 스파이크)은
돌릴 Windows 기계가 생길 때 브랜치를 새로 딴다. `wip/2026-08-28-snapshot`은 main+pilot로
완전히 재구성 가능함을 diff로 확인하고 삭제했다.

배포 zip은 더 이상 커밋하지 않는다(빌드 산출물 36 MB × 38 리비전 = 저장소 1.3 GiB 중
1.27 GiB). 회사 PC가 생기면 **파일럿 브랜치에서 `pnpm package:windows-local`을 돌려** 나온
zip을 가져가 케이스 7을 판정하고, 통과하면 Windows에서 검증 후 main으로 병합한다.

### 목표 구조 — 호스트 어댑터

현재(측정됨): 모든 소비자가 이미 `run`을 주입받고 있고, 리터럴 `Excel.run`/`Office.*`는
`taskpane/main.ts` 한 파일의 13개 지점에만 있다. 즉 시임은 이미 있고, 그 지점만 포트 뒤로
숨기면 된다.

```
           ┌────────────── 공유 (채널과 무관, 손 안 댈) ────────────┐
           │ taskpane view/sheet/chat · formula/* · ai/* · excel/*  │
           └──────────────────────┬──────────────────────┘
                                   │  ExcelHost 포트
                                   │  (run / requirements / onReady)
           ┌──────────────────────┼──────────────────────┐
           │                       │                        │
    host-office.ts          host-hostobject.ts          eval-context.ts
    (Office.js)             (WebView2 host object)      (테스트·평가, 이미 존재)
           │                       │                        │
  Excel Mac/Win (WEF)   Excel Win (XLL in-process COM)   node/vitest
```

자산 경로도 채널마다 다르다: WEF는 `https://localhost:3927`(로컬 서비스+자체 CA),
XLL은 `SetVirtualHostNameToFolderMapping`(서비스·인증서·포트 전부 불필요).

### 단계

| 단계 | 내용 | Windows PC | 상태 |
|---|---|---|---|
| 0 | 파일럿 브랜치 분리 + 근거 문서화 | 불필요 | **완료** |
| 1 | `ExcelHost` 포트 추출 — main.ts 13개 지점을 포트 뒤로, Office.js 어댑터 1개만 존재. 동작 동일 | 불필요 | **완료** |
| 1.5 | 런타임 전역 완전 격리 — `Excel.SheetVisibility` 제거, 가드레일이 enum 읽기까지 잡도록 강화 | 불필요 | **완료** |
| 2 | 컬텍스트 타입 탈Office — `HostContext`를 프로젝트 소유 타입으로 선언, Office 캌스트는 어댑터 한 군데로 | 불필요 | **완료** |
| 2.5 | 어휘·프로토콜 명시와 검증 — 열거형 단어 17개 유니온 + `WordsFit` parity, load/sync 5개 조항 | 불필요 | **완료** |
| 2.6 | **브리지 와이어 확정** — `host-bridge.ts`(op 리스트 프로토콜) + `bridge-memory.ts`(디스패치 참조 구현). 읽기 소비자 전수가 와이어만 통해 동작, 트랜스크립트로 고정 | 불필요 | **완료** |
| 2.7 | 쓰기 경로 op 확장 — `operate.ts` 멤버를 디스패치 테이블에 추가 | 불필요 | 다음 |
| 2.8 | `/xmux/external`·`/xmux/state` 대체 설계 (로컬 서비스가 사라지므로) | 불필요 | 미착수 |
| 3 | XLL 스파이크 — CTP+WebView2 렌더·host object 왕복·미서명 로드 게이트 3개 | **필요** | 대기 |
| 4 | C# 브리지 — 2.6 디스패치 테이블을 `QueueAsMacro` → COM으로 구현. Windows=XLL, Mac=Office.js 병행 | **필요** | 대기 |
| P | (병행, 게이트 아님) 파일럿 브랜치 검증 | **필요** | 대기 |

### 2단계에서 밝혀진 것 (이후 어댑터 작업의 전제)

결정적 측정: **`Excel.RequestContext`는 `InspectContext`도 `OperateContext`도 만족하지
않는다.** (타입 프로브로 확인: `Excel.RequestContext extends InspectContext` → `false`.)
Office의 `autoFill`/`clear` 오버로드가 우리가 이름 붙인 슬라이스보다 넓어서 반공변으로 거부된다.
즉 패인은 처음부터 **경계에서 구조적 캐스트로** Office 컬텍스트를 넘기고 있었고(테스트도
`as unknown as Excel.RequestContext`로 가짜 컬텍스트를 넣고 있었다), 계약을 교차해서 양쪽을
동시에 만족시키려던 시도는 애초에 잘못된 전제였다.

그래서 2단계는 대통합이 아니라 세 가지로 끝난다:

1. `HostContext`를 프로젝트 소유 타입으로 **명시 선언**한다. 교차 시 멤버가 경합하는
   곳(`getRange`, `getUsedRangeOrNullObject`, `autoFill`, `worksheet`)은 `Omit`으로 걷어내고
   한 번만 선언한다 — TypeScript가 교차된 오버로드 중 **첫 번째**를 고르기 때문에 생기는
   문제였다.
2. Office와의 간극은 **어댑터 안 캡스트 한 군데**로 몰아넣는다
   (`context as unknown as HostContext`). 그 캡스트를 정직하게 만드는 근거는
   `office-shapes.ts`의 `KeysFit` parity assert다 — Office가 멤버 이름을 바꾸면 사용자 통합
   문서가 아니라 빌드가 깨진다.
3. 가드레일을 최종형으로 조인다: **`Excel.`/`Office.`를 이름으로 언급할 수 있는 파일은
   딜 둘**(`host-office.ts` 런타임, `office-shapes.ts` 타입 대조). 타입 위치까지 포함해
   전면 금지이며, 캡스트가 단 1회임을 테스트가 검사한다.

**단계 1·2는 파일럿 결과와 무관하게 이득이다** — 테스트 용이성이 오르고, 평가 하네스가 지금
암묵적으로 의존하는 계약이 명시되며, 어느 채널로 가든 필요하다. 그래서 포트는 파생 브랜치가
아니라 `main`에 있다. 단계 2.6–2.8도 같은 이유로 파일럿 판정을 기다리지 않는다.

### 어댑터 유효성 재검토 (2026-08-28)

시임 자체는 유효하다 — 소비자가 전부 `run`을 주입받고, 오직 두 파일만 Office를 이름으로
부를 수 있으며, 테스트가 그걸 강제한다. 다만 **"두 번째 어댑터의 의무는 host.ts 한 파일"**는
과장이었고, 측정해 보니 세 군데가 계약 밖에 있었다. 세 개 모두 이번에 메웠다.

1. **값의 어휘가 `string`으로 지워져 있었다.** Office 열거형은 문자열이라 `autoFill(type: string)`은
   호스트에게 아무것도 알려주지 않는다 — 두 번째 어댑터는 `"FillDefault"`라는 단어를
   `operate.ts`를 grep해서 알아내야 했다. 이제 17개 어휘 유니온을 `office-shapes.ts`에
   선언하고, `WordsFit` parity assert가 **단어 하나하나를 설치된 Office 타이핑과 대조**한다.
   `KeysFit`이 멤버가 존재함을 증명했다면 `WordsFit`은 그 멤버에 무엇을 넘길 수 있는지를
   증명한다. 콜사이트는 한 줄도 고치지 않고 통과했다 — 어휘는 이미 맞았고, 단지 적혀 있지
   않았을 뿐이다.
2. **프로토콜이 계약에도, 테스트에도 없었다.** `HostContext`는 동기 API가 아니라 **지연 객체 그래프**다:
   접근자는 즉시 핸들을 돌려주고, `load`는 의도 선언이며, 값은 `sync` 지점에서만 생긴다.
   멤버를 즉시 읽기로 구현한 호스트는 **타입체크를 통과하고 런타임에 빈 문자열을 돌려준다.**
   타입으로는 잡을 수 없는 종류의 오류라 `host.ts` 상단에 다섯 조항으로 명시했다.
   이게 두 번째 어댑터의 **실제 비용**이다: COM 브리지는 핸들/로드를 큐에 쌓았다가 `sync`에서
   해소하는 지연 계층을 직접 만들어야 한다. `eval-context.ts`는 인메모리라 `sync`가 거의
   no-op이어서 이 비용을 치르지 않았고, 그래서 "읽기 절반의 참조 구현"이지 증명이 아니다
   (타입도 `InspectContext`까지만 맞춘다).

   더 나쁜 건 **테스트도 이걸 안 보고 있었다는 것**이다. 이 저장소의 페이크는 전부
   `load: () => {}`에 값이 미리 채워진 평범한 객체라, 소비자가 **로드 안 한 속성을 읽어도
   sync 전에 읽어도 전부 통과**했다. 이젠 `host-bridge.ts`가 그 자리에 있다(아래 참조) —
   응답에 없는 값은 읽을 수가 없으므로 프로토콜이 규칙이 아니라 **구조**로 강제된다.
   검증: `sheets.ts`의 로드에서 `items/visibility`를 빼면 저장소의 다른 모든 테스트는
   그대로 통과하고 이 스위트만 정확히 그 지점을 짚는다. 쓰기 경로(`operate.ts`)는 아직
   미검증이고, 그게 다음 차례다.
3. **포트 밖에 남은 호스트 의존이 두 개 있다.** 패인은 로컬 서비스에 HTTP로도 말한다:
   `/xmux/external`(외부 통합문서 읽기)과 `/xmux/state`(맥 컴패니언 F2/Tab). XLL은 자산을
   가상 호스트 매핑으로 서빙해서 **로컬 서비스를 없애는 것이 장점**인데, 그러면 이 두 기능이
   같이 죽는다. 채널 비용을 받을 때 빼먹지 않도록 포트 문서에 적어둔다.

고치지 **않은** 것 하나: `main.ts`는 `startOfficeHost`를 직접 import하고 가드레일 테스트가
그 import 문자열을 검사한다. 즉 호스트 **선택** 계층은 없고, 핸드쉐이크(`onReady`)도 `ExcelHost`에
속하지 않는다. 둘 다 두 번째 어댑터가 생기는 순간 5줄짜리 변경이고, 없는 호스트를 위해 지금
팩토리를 두는 건 투기적 추상화다. 비용으로만 기록한다.

그래서 두 번째 어댑터가 지는 의무는 정확히: `host.ts`의 **멤버 목록 + 어휘 + 다섯 조항의
프로토콜**, 그리고 **로컬 서비스가 사라진다면 `/xmux/*` 두 기능의 대체**다. `ExcelHost` 네
메서드는 그중 가장 쉬운 부분이다.

**포트가 왜 RPC가 아니라 프록시 모양인가** — 다시 꾺지 않게 적어둔다. "사각형 하나 읽어라 /
셀 묶음을 써라" 수준의 굵은 포트였다면 COM 브리지가 훨썬 쌀다. 대신 `excel/*`·패인·평가
하네스 전부가 이미 프록시/`load`/`sync` 모양을 전제로 쓰여 있으므로, 굵은 포트는 소비자를
전면 재작성해야 한다. 즉 현재 설계는 **패인을 안 고치는 대가로 어댑터가 지연 계층을 진다**는
교환이고, 의도된 교환이다. 단계 3 스파이크에서 이 지연 계층이 예상보다 비싸게 나오면 그때
재론의하고, 그 전에는 하지 않는다.

### 결정 규칙 (2026-08-28 개정 — 기본값을 뒤집음)

이전 규칙은 "파일럿 판정 먼저, 실패하면 XLL"이었다. **그건 막힌 걸 게이트로 세운 것**이다:
파일럿은 Windows PC가 있어야 돌아가고, 그 PC는 지금 없으며, 파일럿이 검증하려는 Trusted
Catalog도 결국 **WEF 채널**이다. 이 환경에서 WEF 취득은 케이스 1·3·4가 실패,
5가 무의미, 6은 세션 한정 완화책(24h TTL 미검증, 사용자 5단계 수동), 7은 미판정 —
즉 **동작이 확인된 WEF 취득 경로가 하나도 없다.**

개정 규칙:

- **기본 방향은 비WEF(XLL + WebView2 host object)다.** 파일럿은 게이트가 아니라
  **병렬로 긁어보는 싸고 가능성 있는 복권**이다. 통과하면 배포가 당장 쉬워지므로 버리지
  않고, 실패해도 일정이 밀리지 않는다.
- **Windows PC 없이 할 수 있는 XLL 준비를 먼저 끝낸다.** 패인 쪽 절반은 전부 맥에서
  설계·구현·검증된다. 그 작업은 **`adapter/xll-host` 브랜치**에서 진행한다 — 스파이크가
  통과하기 전까지는 미검증 베팅이므로 `main`을 오염시키지 않는다. Windows가 생기면
  "설계"가 아니라 "그 목록 구현"만 남게 한다.
- **Mac은 Office.js로 계속 산다.** 교체가 아니라 병행이고, 포트가 있으므로 공존 비용은
  어댑터 1개다.
- 단계 1·2·2.5는 어느 쪽이든 되돌리지 않는다.

근거가 되는 측정값은 이미 있다 — 대상 PC A 스냅샷에서 `Policies\Microsoft\Office`가
HKCU/HKLM 둘 다 없고 `RequireAddinSig`도 없다. 즉 **미서명 XLL을 정책이 막고 있지 않다.**
XLL 쪽 미검증 리스크는 정책이 아니라 CTP 안 WebView2 초기화 하나로 좁혀있다.

### 비WEF 준비 — 맥에서 끝낸 것과 남은 것

기준: "Windows가 없으면 못 한다"가 진실인 항목만 Windows로 미룬다.

| 항목 | Windows | 상태 |
|---|---|---|
| 패인↔호스트 와이어 프로토콜 (op 리스트 + 응답) | 불필요 | **완료** — `host-bridge.ts` |
| host object가 질 디스패치 테이블 명세 | 불필요 | **완료** — 멤버 8개 + `load`, 트랜스크립트로 고정 |
| 그 테이블의 참조 구현 | 불필요 | **완료** — `bridge-memory.ts`, 읽기 소비자 전수 통과 |
| 쓰기 경로(`operate.ts`) op 확장 | 불필요 | 다음 차례 |
| `/xmux/external`·`/xmux/state` 대체 설계 | 불필요 | 미착수 |
| C# 브리지(디스패치 → `QueueAsMacro` → COM) | **필요** | 대기 |
| XLL 셸 + CTP + WebView2 초기화 | **필요** | 대기 (유일한 진짜 미지수) |
| 미서명 `.xll` 로드 확인 | **필요** | 대기 (정책 스냅샷상 통과 예상) |

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
