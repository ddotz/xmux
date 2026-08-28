# 땡땡엑셀 설치 가이드

> `xmux`는 저장소와 개발 프로젝트의 이름이고, 사용자가 Excel에서 보는 제품명은
> **땡땡엑셀**입니다. `%LOCALAPPDATA%\DdotExcel` 같은 영문 경로는 제품의 내부
> 식별자입니다.

## 1. 설치 방식

땡땡엑셀은 각 Windows PC에서 `https://localhost:3927` 로컬 서비스를 실행합니다.

- 로컬 관리자 권한이 필요하지 않습니다.
- Node.js나 pnpm을 별도로 설치하지 않습니다.
- SMB 공유 폴더나 Microsoft 365 관리자 배포를 사용하지 않습니다.
- 서비스는 이 PC의 loopback 주소(`127.0.0.1`, `::1`)에만 연결되므로 다른 PC에서는
  접근할 수 없습니다.
- Synology Drive는 ZIP 파일을 전달하고 보관하는 용도로만 사용합니다.

## 2. 지원 환경과 사전 확인

- 배포 ZIP 기준: Windows x64
- Windows 10 1607 이상 또는 Windows 11
- Microsoft 365용 Excel 또는 Excel 2019 이상
- Excel을 사용할 Windows 계정으로 로그인
- TCP 3927번 포트가 비어 있어야 함

3927번 포트 사용 여부는 일반 PowerShell에서 확인할 수 있습니다.

```powershell
Get-NetTCPConnection -LocalPort 3927 -State Listen -ErrorAction SilentlyContinue
```

아무것도 표시되지 않으면 포트를 사용할 수 있습니다. 다른 프로세스가 표시되면
그 프로그램을 종료하거나 담당자에게 확인한 뒤 설치합니다.

## 3. 설치

1. Synology Drive에서 `ddot-excel-windows-x64.zip`을 로컬 PC로 복사합니다.
2. ZIP을 우클릭하고 **속성 > 차단 해제 > 확인**을 선택합니다.
   **차단 해제**가 보이지 않으면 이 단계를 건너뜁니다.
3. ZIP 전체를 로컬 폴더에 압축 해제합니다. 네트워크 드라이브나 ZIP 파일 안에서
   직접 실행하지 마십시오.
4. 열려 있는 Excel 창을 모두 닫습니다.
5. 압축을 푼 폴더에서 **`땡땡엑셀 설치.bat`을 두 번 클릭**합니다. 관리자 권한은
   필요하지 않습니다.
6. 메뉴에서 **1. 설치 / 업데이트**를 선택합니다.

압축을 푼 폴더의 구성은 다음과 같습니다. 사용자가 실행하는 파일은 맨 위의
`땡땡엑셀 설치.bat` 하나뿐입니다.

```text
땡땡엑셀 설치.bat     <- 두 번 클릭
app\                  프로그램 파일
runtime\              내장 Node 런타임
scripts\              설치·관리·제거 스크립트
```

정상 설치되면 다음 주소가 표시됩니다.

```text
Service: https://localhost:3927
```

7. 설치가 끝나면 **Office 첫 실행 초기화**가 Excel을 엽니다. 새 통합 문서에서
   **홈 > 추가 기능 > 더 많은 추가 기능 > 개발자 추가 기능**의
   **땡땡엑셀 > 추가**를 선택합니다. Excel 버전에 따라 첫 메뉴가
   **추가 기능 가져오기**로 표시될 수 있습니다.
8. 작업창이 바로 열리면 초기화 창으로 돌아와 Enter를 누릅니다. Excel 아래쪽에
   **추가 기능 로드 오류**가 표시되면 다음 순서를 한 번만 수행합니다.
   1. 아래쪽 오류를 클릭합니다.
   2. 열린 **Office 추가 기능** 화면을 닫습니다.
   3. 개발자 추가 기능에서 **땡땡엑셀 > 추가**를 다시 선택합니다.
   4. 작업창이 열린 것을 확인하고 초기화 창으로 돌아와 Enter를 누릅니다.
9. 안내에 따라 Excel을 완전히 닫습니다. 초기화 도구는 이번 시도 이후의
   `/index.html` 요청과 Office WEF 캐시 ID를 확인한 뒤에만 완료 상태를 기록합니다.
10. Excel을 다시 실행하고 **홈 > 땡땡엑셀** 리본 버튼을 눌러 작업창을 엽니다.

메뉴 대신 PowerShell을 직접 쓰려면 압축을 푼 폴더에서 다음 명령을 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\install.ps1
```

이 절차는 현재 Windows 사용자의 현재 Office WEF 캐시마다 최초 1회만 필요합니다.
Office 캐시가 초기화되어 다시 필요해지면 설치 메뉴의
**5. Office 첫 실행 초기화 다시 실행**을 선택합니다.

## 4. 설치 확인

`땡땡엑셀 설치.bat`을 실행하고 **2. 상태 확인**을 선택합니다. PowerShell에서
직접 확인하려면 압축을 푼 폴더에서 다음 명령을 실행합니다.

```powershell
.\scripts\manage.ps1 status
```

설치 폴더의 관리 스크립트를 직접 실행할 수도 있습니다.

```powershell
& "$env:LOCALAPPDATA\DdotExcel\manage.ps1" status
```

정상 상태:

```text
DdotExcel local service is running at https://localhost:3927.
```

브라우저에서 `https://localhost:3927/health`를 열었을 때 정상 응답은 다음과
같습니다.

```json
{"service":"ddot-excel","status":"running"}
```

## 5. 설치되는 항목

모든 항목은 현재 Windows 사용자 범위에만 설치됩니다.

| 항목 | 위치 또는 이름 |
|---|---|
| 앱·런타임 | `%LOCALAPPDATA%\DdotExcel` |
| Office 등록 | `HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer` |
| 로그인 자동 실행 | `HKCU\Software\Microsoft\Windows\CurrentVersion\Run` |
| 인증서 | `Cert:\CurrentUser\My`, `Cert:\CurrentUser\Root` |
| 서비스 주소 | `https://localhost:3927` |

설치기는 시스템 전역 레지스트리, Windows 서비스, 방화벽 규칙 또는 다른 사용자
프로필을 변경하지 않습니다.

## 6. 서비스 제어

`땡땡엑셀 설치.bat` 메뉴의 **3. 서비스 다시 시작**으로 대부분 해결됩니다.
세부 제어는 압축을 푼 폴더에서 다음 명령을 사용합니다.

```powershell
.\scripts\manage.ps1 status
.\scripts\manage.ps1 start
.\scripts\manage.ps1 stop
.\scripts\manage.ps1 restart
```

서비스는 Windows 로그인 시 자동으로 시작됩니다.

## 7. 업데이트

1. 새 ZIP을 내려받고 **속성 > 차단 해제**를 적용합니다.
2. 새 로컬 폴더에 전체 압축 해제합니다.
3. Excel을 완전히 종료합니다.
4. 새 폴더에서 `땡땡엑셀 설치.bat`을 실행하고 **1. 설치 / 업데이트**를
   선택합니다.
5. Excel을 다시 실행합니다.

설치기는 기존 프로세스를 정지한 뒤 제품 소유 파일만 교체합니다. 아직 유효한
인증서는 재사용합니다. 기존 Office 등록과 AI 연결 설정도 같은 origin을 사용하는
동안 유지됩니다.

## 8. 제거

`땡땡엑셀 설치.bat`을 실행하고 **4. 제거**를 선택한 뒤 확인 질문에 `y`를
입력합니다. 압축을 푼 폴더가 없으면 설치 폴더의 제거 스크립트를 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File "$env:LOCALAPPDATA\DdotExcel\uninstall.ps1"
```

제거 후 Excel을 완전히 종료했다가 다시 실행합니다.

제거기는 자신이 등록한 시작 항목, Office 등록, 인증서와 설치 파일만 삭제합니다.
같은 이름의 항목이 다른 경로를 가리키면 삭제하지 않고 경고를 표시합니다.

## 9. 문제 해결

### `TCP port 3927 is already used by process ...`

다른 프로그램이 3927번 포트를 사용 중입니다.

```powershell
Get-NetTCPConnection -LocalPort 3927 -State Listen
```

표시된 프로세스를 확인하고 종료한 뒤 다시 설치합니다.

### `The deployment package is incomplete`

ZIP이 부분적으로 압축 해제되었습니다. 새 로컬 폴더에 전체 압축 해제한 뒤 다시
실행합니다.

### 서비스가 15초 안에 시작되지 않음

백신, AppLocker 또는 WDAC가 포함된 `node.exe` 실행을 차단했을 수 있습니다.
회사 정책 담당자에게 `%LOCALAPPDATA%\DdotExcel\runtime\node.exe` 실행 허용을
문의합니다.

### HTTPS health check 실패

현재 사용자 인증서 신뢰가 차단되었을 수 있습니다. 제거 후 다시 설치해도
반복되면 회사의 인증서 정책을 확인해야 합니다.

### 개발자 추가 기능에 땡땡엑셀이 보이지 않음

모든 Office 앱을 완전히 종료한 뒤 Excel을 다시 실행합니다. 설치 메뉴의
**5. Office 첫 실행 초기화 다시 실행**도 시도합니다. 계속 보이지 않으면 회사
정책이 Office 개발자 추가 기능을 차단한 것입니다.

### Office 첫 실행 초기화가 완료되지 않음

설치 도중 작업창 요청이 확인되지 않으면 설치기는 로컬 서비스는 남겨 두되 설치
완료로 보고하지 않습니다. `땡땡엑셀 설치.bat`을 다시 실행하고
**5. Office 첫 실행 초기화 다시 실행**을 선택합니다. 작업창이 열린 뒤에도 실패하면
`%LOCALAPPDATA%\DdotExcel\service.log`와 Office 정책을 확인합니다.

### Excel 시작 시 "추가 기능 로드 중 오류 발생" · 등록이 계속 풀림

Excel이 시작될 때 `https://localhost:3927`이 응답하지 않으면 Excel은 이 오류를
표시하고 **개발자 등록을 스스로 삭제**합니다. 즉 원인은 등록이 아니라, 로그인
시 서비스가 실행되지 못했거나 Excel보다 늦게 시작된 것입니다.

**먼저 버전을 확인합니다.** 1.10.5 이전 패키지의 서비스는 IPv4(`127.0.0.1`)에만
바인딩했습니다. Windows는 `localhost`를 `::1`(IPv6)로 먼저 해석하므로, Excel의
시작 시 요청은 실패하고 작업창을 직접 열 때만 성공했습니다. 증상은 이렇습니다.

- Excel을 껐다 켤 때마다 리본에서 땡땡엑셀 단추가 사라짐
- 추가 기능 목록에서 한 번 오류를 보고 다시 추가해야 사용 가능
- 그렇게 추가하면 정상 동작

`manage.ps1 status`는 이 상태에서도 전부 정상으로 보입니다. 서비스와 상태 점검이
모두 IPv4로 접속하기 때문입니다. 새 ZIP으로 다시 설치하면 서비스가 두 주소를 모두
수신하므로 해결됩니다. 확인:

```powershell
Get-NetTCPConnection -LocalPort 3927 -State Listen | Select-Object LocalAddress
```

`127.0.0.1`과 `::1`이 함께 나오면 정상입니다. `127.0.0.1`만 나오면 예전 패키지입니다.

일반 PowerShell에서 시작 체인을 점검합니다.

```powershell
& "$env:LOCALAPPDATA\DdotExcel\manage.ps1" status
```

| 출력 | 의미와 조치 |
|---|---|
| `Office registration: MISSING` | Excel이 로드 실패 후 등록을 지운 상태입니다. 서비스가 실행 중이면 자동으로 복원되므로 Excel을 완전히 종료했다가 다시 실행합니다. 재설치는 필요 없습니다. |
| `Logon autostart: MISSING` | 보안 도구가 로그인 시작 항목을 삭제했습니다. `땡땡엑셀 설치.bat`에서 **1. 설치 / 업데이트**를 다시 실행합니다. |
| `Logon autostart approval: DISABLED` | 작업 관리자 > 시작 앱에서 `DdotExcelLocalService`가 "사용 안 함"으로 바뀐 상태입니다. 다시 사용으로 바꾸거나 **1. 설치 / 업데이트**를 다시 실행합니다. |
| `Windows Script Host: DISABLED` | 회사 정책이 wscript 실행을 차단했습니다. **1. 설치 / 업데이트**를 다시 실행하면 PowerShell 경유 시작으로 자동 전환됩니다. |

서비스가 멈춰 있으면 시작한 뒤 Excel을 다시 실행합니다.

```powershell
& "$env:LOCALAPPDATA\DdotExcel\manage.ps1" start
```

서비스는 실행 중인 동안 Office 등록을 주기적으로 복원합니다. 이 오류를 본
뒤에는 Excel을 다시 시작하는 것으로 충분하며, 그래도 반복되면 위 표에서
로그인 시작이 막힌 원인을 찾아야 합니다.

### 작업창이 열리지 않거나 빈 화면

```powershell
& "$env:LOCALAPPDATA\DdotExcel\manage.ps1" restart
```

재시작 후 `https://localhost:3927/health`를 확인합니다.

## 10. 회사 PC 정책 관련 제한

다음 항목 중 하나라도 회사 정책으로 차단되면 IT 정책 변경 없이는 설치할 수
없습니다.

- PowerShell 스크립트 실행
- 현재 사용자 인증서 등록
- `%LOCALAPPDATA%`의 실행 파일
- Office 개발자 추가 기능 sideload
- 현재 사용자 로그인 시작 항목

이 패키지는 Microsoft의 사용자별 **개발자 sideload** 채널을 사용하는 소수 내부
PC용 배포 방식입니다. Microsoft가 공식 지원하는 조직 배포 경로는 Marketplace
또는 Microsoft 365 관리자 배포입니다.

AI 대화 기능에 필요한 API 키는 패키지에 포함되어 있지 않습니다.
