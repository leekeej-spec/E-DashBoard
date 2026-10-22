# E-Dashboard

EAS 알람 처리용 SharePoint 리스트(`EasDashboard`, `EASDB2`)의 현황을 보여주는 대시보드입니다.

## 현재 방식: Power Apps (권장)

**[docs/POWERAPPS.md](docs/POWERAPPS.md)** 를 따라 캔버스 앱을 만듭니다.

- 앱이 두 리스트를 **각자 자기 계정으로 직접** 읽습니다. 그래서 내보내기 Flow, OneDrive 파일, 동기화, 스크립트 실행이 모두 필요 없습니다.
- 권한은 SharePoint 사이트 권한을 그대로 따릅니다.
- 배포는 앱 공유나 Teams 채널 탭으로 합니다.
- SharePoint 커넥터는 표준 커넥터라 M365 라이선스로 충분합니다.

### 화면 구성

- **헤더**: 마지막 갱신 시각, 주간/야간, 18:00 초기화까지 남은 시간, 새로고침 (60초마다 자동)
- **타일**: 전체 / 위험 / Reminder / AllOk 누락 의심 / 정상 / Teams 발송 / 미매칭. 클릭하면 필터, 다시 클릭하면 해제
- **필터**: 검색(SRN·사이트·시스템·FSE), Region, FSE, 에러만
- **목록**: 위험 → Reminder → 정상 순서이고, 같은 상태 안에서는 최근 에러가 위로 옵니다. 에러 칸 7·9·11·13·20이 활성이면 색이 칠해집니다.
- **상세 패널**: 선택한 시스템의 요약, SharePoint 항목 열기
- **알림**: 새로 위험이 된 시스템이 생기면 앱 상단에 알림 배너가 뜹니다.

### 상태 판정 규칙

| 상태 | 조건 |
|---|---|
| 정상 | 에러 컬럼이 모두 비어 있음. 또는 AllOk = 예이고 AllOk 시각이 마지막 에러 시각보다 늦음 |
| 위험 | 7·9·11·13 에러 컬럼 중 하나라도 값이 있음 (냉각수 / Coldhead) |
| Reminder | 20 에러 컬럼에만 값이 있음 |
| AllOk 누락 의심 | 위험/Reminder인데 마지막 에러 후 6시간 이상 갱신 없음 |
| 미매칭 | EASDB2에만 있고 EasDashboard에는 없는 SRN |

- 에러 컬럼의 값이 비어 있거나 `No`/`false`/`0`이면 "없음"으로 봅니다.
- 두 리스트는 SRN에서 따옴표·공백을 지우고 대문자로 맞춘 값으로 연결합니다.
- 같은 SRN 행이 여러 개면 `[중복]`이 붙습니다.

### 수식 테스트 (개발용)

`tests/powerfx/Test-PowerAppsFormulas.ps1`(PowerShell 7)는 가이드의 수식을 Power Fx 엔진으로 가짜 SharePoint 데이터에 실행해서 문법과 상태 판정을 확인합니다.
- `-TextDates`, `-ChoiceRegion` 옵션은 가이드 12장의 컬럼 형식별 수정 사항을 검증합니다.
- 가이드의 수식을 고쳤다면 이 스크립트를 다시 돌려 보세요.

## 이전 방식 (참고용, 지금은 사용하지 않음)

Power Apps로 바꾸기 전에 만든 것들입니다. 필요 없으면 지워도 됩니다.

| 파일 | 내용 |
|---|---|
| `EDashboard.ps1`, `EDashboard.cmd`, `EDashboard-Demo.cmd` | PowerShell + WinForms 데스크톱 대시보드 |
| [docs/FLOW-EXPORT.md](docs/FLOW-EXPORT.md), `samples/` | 그 대시보드에 데이터를 넘기던 내보내기 Flow와 샘플 파일 |
| `tools/Test-EDashEnvironment.*`, [docs/TESTING.md](docs/TESTING.md) | PC·SharePoint 환경 점검 스크립트와 사용법 |
| [docs/DESIGN.md](docs/DESIGN.md) | 처음 설계안 (PowerShell 기준) |

PowerShell 방식을 그만둔 이유: `share.philips.com`은 SharePoint Online(`philips.sharepoint.com`)이라 Windows 계정만으로는 REST를 호출할 수 없습니다(HTTP 403). 그래서 Flow → 파일 → 동기화로 우회해야 했고, 단계가 많았습니다.
