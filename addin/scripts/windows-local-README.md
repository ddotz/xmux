# 땡땡엑셀 Windows 로컬 설치

이 배포판은 각 PC에서 `https://localhost:3927` 서비스를 자체 실행합니다. 대상
PC에는 Node.js나 pnpm을 설치할 필요가 없습니다.

- 상세 설치·업데이트·제거: [`INSTALL.md`](INSTALL.md)
- 제품 기능과 사용법: [`USER-GUIDE.md`](USER-GUIDE.md)

## 설치

1. 내려받은 ZIP을 우클릭해 **속성 > 차단 해제 > 확인**을 누릅니다. 차단 해제
   항목이 보이지 않으면 이 단계를 건너뜁니다.
2. ZIP 파일을 로컬 폴더에 완전히 압축 해제합니다.
3. Excel 창을 모두 닫습니다.
4. **일반 Windows PowerShell**을 실행합니다. 관리자 권한은 필요하지 않습니다.
5. 압축을 푼 폴더에서 다음 명령을 실행합니다.

   ```powershell
   powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install.ps1
   ```

6. Excel을 실행합니다.
7. **홈 > 추가 기능 > 더 많은 추가 기능 > 개발자 추가 기능**에서 `땡땡엑셀`을
   선택하고 **추가**를 누릅니다. Excel 버전에 따라 첫 메뉴가 **추가 기능 가져오기**
   로 표시될 수 있습니다.

설치기는 다음 항목만 이 PC의 현재 사용자에게 구성합니다.

- `%LOCALAPPDATA%\DdotExcel`의 정적 앱과 전용 Node 런타임
- 로그인할 때 서비스를 시작하는 현재 사용자 시작 프로그램
- 현재 사용자만 신뢰하는 `localhost` 자체 서명 인증서
- Excel의 현재 사용자 개발자 추가 기능 등록

서비스는 loopback 주소에만 바인딩되므로 다른 PC에서 접근할 수 없습니다.
설치 전에 TCP 3927번 포트가 비어 있어야 합니다.
Synology Drive는 ZIP을 전달하거나 보관하는 용도로 사용할 수 있습니다. 설치된
서비스와 Excel 등록 정보는 각 PC의 현재 사용자 프로필에 저장됩니다.

## 서비스 제어

일반 PowerShell에서 다음 명령을 사용할 수 있습니다.

```powershell
.\manage.ps1 status
.\manage.ps1 start
.\manage.ps1 stop
.\manage.ps1 restart
```

상태 확인 주소는 `https://localhost:3927/health`입니다.

## 업데이트

새 ZIP을 압축 해제하고 일반 PowerShell에서 `install.ps1`을 다시 실행합니다.
기존 서비스 프로세스와 앱 파일은 땡땡엑셀 소유 범위 안에서 교체됩니다.

## 제거

일반 PowerShell에서 다음 명령을 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1
```

제거 후 Excel을 완전히 종료했다가 다시 실행합니다.

AI 대화 기능은 별도의 API 네트워크 연결과 사용자별 API 설정이 필요합니다.
회사 보안 정책이 PowerShell, 사용자 인증서 또는 Office 개발자 추가 기능을
차단하면 IT 정책 변경 없이 설치할 수 없습니다.

이 패키지는 Microsoft의 사용자별 **개발자 sideload** 채널을 사용합니다. 소수
내부 PC에서 직접 설치하는 용도이며, Microsoft가 지원하는 정식 조직 배포 경로는
Marketplace 또는 Microsoft 365 관리자 배포입니다.
