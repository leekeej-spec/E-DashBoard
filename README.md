# E-Dashboard

EAS 알람 처리용 SharePoint 리스트(`EasDashboard`, `EASDB2`)를 보여주는 PowerShell + WinForms 데스크톱 대시보드입니다. 현재는 P1 조회 버전이며, SharePoint에는 아무것도 쓰지 않습니다.

## 실행 방법

1. 아래 파일들을 같은 폴더(예: `C:\EDash\`)에 받습니다. `Code → Download ZIP`으로 받아 압축을 풀어도 됩니다.
   - `EDashboard.ps1`
   - `EDashboard.cmd`
   - `EDashboard-Demo.cmd`

   `EDashboard.ps1`은 한글이 들어 있어서 **UTF-8(BOM)** 이어야 합니다. "Download raw file"이나 ZIP으로 받고, 메모장에 복사·붙여넣기하지 마세요.
2. **처음 한 번:** [docs/FLOW-EXPORT.md](docs/FLOW-EXPORT.md)대로 데이터 내보내기 Flow를 만듭니다.
   - `share.philips.com`은 SharePoint Online이라 Windows 계정만으로는 직접 읽을 수 없습니다.
   - 대신 Power Automate가 두 리스트를 OneDrive의 `EDashboard\edash-data.json`으로 3분마다 내보내고, 대시보드는 동기화된 그 파일을 읽습니다.
3. 실행합니다.

| 파일 | 용도 |
|---|---|
| `EDashboard-Demo.cmd` | 샘플 데이터로 화면 확인 (SharePoint 연결 불필요). 갱신할 때마다 한 대씩 상태가 바뀌어서 알림·색상 변화를 볼 수 있음 |
| `EDashboard.cmd` | 실제 데이터: `%OneDriveCommercial%\EDashboard\edash-data.json` |

검은 콘솔 창은 대시보드가 실행되는 동안 열어 두세요. 콘솔 창을 닫으면 대시보드도 함께 종료됩니다.

### 옵션

`EDashboard.cmd` 뒤에 붙입니다. 예: `EDashboard.cmd -RefreshSeconds 30`

| 옵션 | 기본값 | 설명 |
|---|---|---|
| `-RefreshSeconds` | 60 | 갱신 주기 (최소 15초) |
| `-StaleHours` | 6 | 에러가 이 시간 이상 갱신되지 않으면 "AllOk 누락 의심" |
| `-DailyReset` | 18:00 | 초기화 Flow 시각 (카운트다운 표시용) |
| `-NightStartHour` / `-NightEndHour` | 22 / 8 | 주간·야간 표시 |
| `-DataFile` | OneDrive\EDashboard\edash-data.json | 데이터 파일 위치 |
| `-MaxFileAgeMinutes` | 15 | 데이터 파일이 이보다 오래되면 "데이터 오래됨" 경고 |
| `-Source Rest` | – | 온프레미스 SharePoint를 REST로 직접 읽을 때 (지금 환경에서는 사용 안 함) |
| `-SiteUrl`, `-DashboardList`, `-AlertList` | 현재 사이트·리스트 | 항목 열기 링크와 REST 모드에 사용 |
| `-CheckOnly` | – | 화면 없이 한 번 읽고 콘솔에 요약만 출력 (연결 확인용) |

## 화면 구성

- **헤더**: 데이터 상태, 데이터 생성 시각(KST, 몇 분 전인지), 주간/야간, 18:00 초기화까지 남은 시간, 새로고침(F5)
- **타일**: 전체 / 위험 / Reminder / AllOk 누락 의심 / 정상 / Teams 발송 / 미매칭. 타일을 클릭하면 그 조건으로 필터링되고, 다시 클릭하면 해제됩니다.
- **필터**: 검색(SRN·사이트·시스템·FSE), Region, FSE, 에러만 보기, CSV 내보내기, 새 위험 알림 켜기/끄기
- **목록**: 위험 → Reminder → 미매칭 → 정상 순서이고, 같은 상태 안에서는 최근 에러가 위로 옵니다.
  - 행 색: 빨강 = 위험, 주황 = Reminder, 회색 = EASDB2에만 있는 SRN
  - 보라 기울임 = AllOk 누락 의심
  - 에러 컬럼의 ● = 그 에러가 활성
- **상세** (아래 칸): 선택한 시스템의 요약과 두 리스트의 원본 값
- **행 조작**: 더블클릭하면 SharePoint 항목이 열립니다. 우클릭 메뉴로 EASDB2 항목 열기, SRN 복사를 할 수 있습니다.
- **알림**: 새로 위험이 된 시스템이 생기면 트레이 풍선 알림과 소리가 납니다. 18:00 초기화처럼 한꺼번에 복구되면 알림 한 건으로 묶습니다.

## 상태 판정 규칙

| 상태 | 조건 |
|---|---|
| 정상 | 에러 컬럼(field_6~10)이 모두 비어 있음. 또는 AllOk = 예이고 AllOk 시각이 마지막 에러 시각보다 늦음 |
| 위험 | 7·9·11·13 에러 컬럼 중 하나라도 값이 있음 (냉각수 / Coldhead) |
| Reminder | 20 에러 컬럼에만 값이 있음 |
| 미매칭 | EASDB2에만 있고 EasDashboard에는 없는 SRN |

- 에러 컬럼의 값이 비어 있거나 `No`/`false`/`0`이면 "없음"으로 봅니다.
- 두 리스트는 SRN에서 따옴표·공백을 지우고 대문자로 맞춘 값으로 연결합니다.
- 같은 SRN 행이 여러 개면 상태 칸에 `[중복]`이 붙습니다.
- 리스트에 없는 컬럼이 있으면 아래 상태줄에 "리스트에 없는 컬럼: ..."으로 표시됩니다.

## 문제가 생기면

- **연결 실패**가 나오면 아래 칸에 원인이 표시되고, 갱신 주기마다 자동으로 다시 시도합니다.
- `samples/edash-data.sample.json`은 내보내기 파일의 예시입니다. `EDashboard.cmd -DataFile samples\edash-data.sample.json`으로 실행해 볼 수 있고, 이 파일의 시각은 고정되어 있어서 "데이터 오래됨" 경고가 뜨는 것이 정상입니다.
- 환경 점검은 `tools/Test-EDashEnvironment.cmd`로 합니다. 방법은 [docs/TESTING.md](docs/TESTING.md)에 있습니다.
- 로그 위치: `%LOCALAPPDATA%\EDashboard\edashboard-yyyyMMdd.log`

## 문서

- 설계안: [docs/DESIGN.md](docs/DESIGN.md)
- 테스트 방법: [docs/TESTING.md](docs/TESTING.md)
