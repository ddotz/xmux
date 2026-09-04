# Windows LTSC 웹 추가 기능 삽입 — 증거와 실험 플레이북

갱신: 2026-09-04. 대상: Office LTSC 2024 2408 (16.0.17932.20842), Windows 10/11,
망분리 사내 PC. 측정 원본은 `FINDINGS.md` F8–F10이다.

## 확인된 문제

Developer 레지스트리와 사용자 단위 Trusted Catalog는 서로 다른 등록 채널이지만,
지금까지의 실험은 둘 다 **Office 추가 기능 대화상자에서 Add를 눌러 취득**했다.
Developer 콜드 실패에서는 Excel이 `SourceLocation`을 요청하기도 전에 멈췄다.

F8/F9가 확인한 경계:

- 첫 Add는 실패했고 `service.log`에 `/index.html` 요청이 없었다.
- 오류 화면을 여는 순간 이후 같은 Excel 프로세스의 두 번째 Add는 성공했다.
- 오류 화면을 닫을 때 레지스트리 변화는 없었다.
- 성공 상태의 WEF 캐시와 레지스트리를 콜드 프로필에 재생해도 실패했다.
- B0→B 디스크 델타는 `UserIdentityCache`뿐이며 `ExcelOmexUserIdentity=Anonymous`가
  나타났다. 활성화 조건은 디스크 pre-seed가 아니라 Excel 프로세스 메모리에 있었다.
- 2026-09-01 Trusted Catalog 파일럿도 실패했다. 상세 실패 단계 캡처는 회수되지
  않았으므로, 이것은 해당 end-to-end 후보를 닫지만 Office 내부 단계를 특정하지 않는다.

따라서 확인된 실패는 **WEF 전체**가 아니라 **Add 취득 경로**다. 이 구분 없이 XLL로
전환한 이전 결정은 철회한다.

## 새 우선순위

### A. 웹익스텐션 임베드 통합문서 — 기본 경로

Microsoft `office-addin-dev-settings sideload`가 Excel 데스크톱에서 사용하는 것과 같은
OOXML 구조를 설치 시 생성한다.

```xml
<!-- xl/webextensions/webextension.xml -->
<we:reference id="6374B2A1-D997-4BB0-B23B-17F28561827B"
              version="1.0.0.0"
              store="developer"
              storeType="Registry" />
```

`xl/webextensions/taskpanes.xml`은 `visibility="1"`이다. 설치기는 매니페스트의 실제
ID와 4-part 버전을 읽어 `%LOCALAPPDATA%\DdotExcel\땡땡엑셀 시작.xlsx`를 만든다.
Office Add-ins 대화상자의 카탈로그 열거와 Add 클릭을 거치지 않고 문서 내 참조가
`HKCU\...\Wef\Developer`를 직접 해석하게 한다.

구현:

- `create-wef-bootstrap-workbook.ps1`: Microsoft 공식 sideload 템플릿과 같은 8개 OOXML
  파트를 생성하며 바이너리 템플릿을 저장소에 넣지 않는다.
- `initialize-windows-local.ps1`: Excel COM으로 시작 통합문서를 열고, 새
  `GET /index.html -> 200`만 성공으로 인정한다.
- 설치와 메뉴 5번 모두 이 경로를 사용한다.

**상태: 구현·OOXML 구조 검증 완료, 대상 Windows 콜드 프로필 미검증.**

### B. COM 자동 워밍업 — A 실패 시 같은 실행의 fallback

A에서 20초 동안 새 요청이 없으면 새 Excel 프로세스에서 빈 통합문서를 만든 뒤 다음을
실행한다.

```powershell
$excel.CommandBars.ExecuteMso("OfficeExtensionsDialog")
```

별도 숨김 PowerShell이 해당 Excel 프로세스를 활성화하고 Escape를 보내 모달을 닫는다.
그 다음 **같은 Excel 프로세스**에서 임베드 시작 통합문서를 연다. F9의 인과는 대화상자를
여는 행위이고 닫기는 상태를 기록하지 않으므로, 기존 사용자 5단계 절차를 자동화한 것이다.

**상태: 구현·PowerShell 구문 검증 완료, 대상 Windows UI 동작 미검증.**

### C. Omex 조회 차단 정책 — 명시적 실험만

```text
HKCU\Software\Policies\Microsoft\Office\16.0\WEF\TrustedCatalogs
  DisableOmexCatalogs = 1 (DWORD)
```

메뉴 6번만 이 값을 설정한다. 기본 설치에는 적용하지 않는다. 기존 값의 존재 여부와 값을
`HKCU\Software\DdotExcel`에 기록하고 제거 시 복원한다. 설치 후 누군가 값을 바꾸면 제거기는
덮어쓰지 않는다.

금지:

- `Common\Privacy\DisconnectedState=2`: 추가 기능 UI와 sideload까지 죽일 수 있다.
- `UseOnlineContent=0`: 인증 경로 부작용이 있어 이 실험에 섞지 않는다.
- `DisableAllCatalogs=1`: Developer/Trusted Catalog까지 막으므로 목적과 반대다.

**상태: 구현 완료, 가설 미검증.** F9의 Omex 익명 신원 델타와 정합하지만 성공 증거는 없다.

## 판정 순서

대상 PC의 WEF 상태를 진단 도구로 초기화하고 각 실험을 독립 실행한다.

1. 정책 없이 설치한다. Add-ins를 열지 말고 설치기가 연 시작 통합문서의 새
   `/index.html` 요청을 확인한다.
2. A가 실패하면 설치기가 자동으로 B를 실행한다. 사용자 입력 없이 요청이 생기는지,
   Escape가 Excel 본창을 닫지 않는지 확인한다.
3. A+B가 실패하면 WEF 상태를 다시 초기화하고 메뉴 6번으로 C+A+B를 실행한다.
4. 성공한 최초 단계에서 Excel을 완전히 종료하고 일반 통합문서로 재시작, 재부팅,
   24시간 뒤를 각각 확인한다.
5. `diagnose-wef-firstrun.ps1 snapshot`과 `service.log`를 회수해 `FINDINGS.md`에 결과를
   추가한다.

성공 기준은 **클린 WEF 프로필에서 사용자 Add 없이 새 `/index.html` 요청과 작업창 표시**, 그리고
새 프로세스·재부팅 후 재현이다. 시작 통합문서에서만 열리는 경우도 Add 취득 우회로서는
성공이지만, 일반 통합문서 자동 노출과는 별도 판정한다.

## 폐기한 방향

XLL + WebView2 host object 경로는 제품 방향에서 폐기했다. 로컬 브랜치와 생성 산출물을
제거했으며 원격 실험 브랜치는 병합하지 않는다. 이유는 새 증거가 네이티브 호스트의 필요성을
보인 것이 아니라, 지금까지 서로 다른 것으로 취급한 두 WEF 실험이 같은 Add 취득 표면이었다는
점을 보였기 때문이다. 검증 비용이 가장 낮고 기존 제품 전체를 유지하는 A→B→C를 먼저 판정한다.

`ExcelHost` 포트는 배포 베팅과 무관한 테스트·격리 개선이므로 `main`에 유지한다.
