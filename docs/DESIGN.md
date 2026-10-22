# E-Dashboard 설계안 (PowerShell + WinForms)

EAS 알람 처리용 SharePoint 리스트 2개(`EasDashboard`, `EASDB2`)를 읽어서 보여주는 데스크톱 대시보드입니다.
처음에는 조회만 하고, 기능은 나중에 플러그인처럼 붙여 나갈 수 있도록 구조를 잡았습니다.

---

## 0. 요약

| 항목 | 결정 |
|---|---|
| 런타임 | **Windows PowerShell 5.1**: 회사 PC에 기본 설치되어 있고, Windows 통합인증(`-UseDefaultCredentials`)과 WinForms·Charting을 그대로 사용 가능 |
| SharePoint 접근 | **REST API** (`/_api/web/GetList(...)`): 별도 모듈 설치가 필요 없고, 리스트 표시이름이 바뀌어도 URL로 찾으므로 영향 없음 |
| 화면 멈춤 방지 | 데이터 조회는 **별도 Runspace(백그라운드)**, 화면 반영은 **UI 타이머**에서 처리 |
| 확장 방식 | 화면 탭과 기능을 **플러그인(hashtable 반환 .ps1)** 으로 만들고, **이벤트 버스**로 데이터를 받음 |
| 설정 | 팀 공용 `.psd1`(주석 가능) + 개인 설정 `%APPDATA%\EDashboard\user.json` |
| 단계 | P1 조회 MVP → P2 알림·이력·정합성 → P3 조치(쓰기) → P4 분석·연계 |

---

## 1. 현재 구조 정리

```
[EAS 알림 메일] ─▶ Flow1 (메일 파싱: SRN, HFN, ErrorCode, 주/야간)
                     ├─▶ EasDashboard : SRN당 1행. 에러별 필드, ActiveErrorCount, AllOk ...
                     ├─▶ EASDB2       : SRN당 1행. AlertSent(Teams 발송 여부)
                     └─▶ Teams 알림 (크리티컬/야간/반복)

[매일 18:00]     ─▶ Flow2 (초기화)
                     ├─▶ EasDashboard : AllOk가 아닌 행 → 에러필드·StartMonitoring 비움
                     └─▶ EASDB2       : AlertSent = true → false

                          ▲ 주기적으로 REST로 읽기 (기본 60초)
                 E-Dashboard (PowerShell + WinForms)
```

대시보드에서 바로 쓸 수 있는 정보:

| 리스트 | 키 | 주요 정보 |
|---|---|---|
| EasDashboard | `SRN` | Site/System/FSE/Region, 에러 컬럼 5개(7·9·11·13·20), ActiveErrorCount, ReminderCount, AllOk/AllOkTime, LastUpdated/LastUpdatedErrorTime, StartMonitoring |
| EASDB2 | `SystemReferenceNumber` | AlertSent (그 외 컬럼은 확인 필요, 12장 참고) |

두 리스트는 **SRN = SystemReferenceNumber** 로 연결합니다. 이때 SRN을 정규화(공백·따옴표 제거, 대문자)한 `SrnKey`를 키로 씁니다. 이유는 10장에 있습니다.

---

## 2. 기술 선택

| 항목 | 선택 | 비고 |
|---|---|---|
| 런타임 | Windows PowerShell 5.1 | PS7도 WinForms는 동작하지만 `DataVisualization.Charting`이 없음 |
| 인증 | Windows 통합인증 (`-UseDefaultCredentials`) | 온프레미스 SharePoint 기준. SharePoint Online이면 PnP.PowerShell 쪽 Provider로 교체 |
| 조회 | REST `GetList(@a1)/items` + `$select` + 페이징(`odata.nextLink`) | `odata=nometadata` 사용 |
| 쓰기(P3) | REST `MERGE` + `X-RequestDigest` + `IF-MATCH` | 설정 `EnableWrite = $true`일 때만 동작 |
| UI | WinForms를 코드로 구성 (디자이너 사용 안 함) | 레이아웃은 `TableLayoutPanel`, `SplitContainer`, `Dock` |
| 비동기 | `[powershell]` + 전용 Runspace + `BeginInvoke` | UI 컨트롤은 UI 스레드에서만 건드림 |
| 로컬 저장 | CSV/JSON (`%LOCALAPPDATA%\EDashboard\`) | 외부 DLL 없음. SQLite는 필요할 때 도입 |
| 테스트 | Pester | 데이터·로직 모듈을 UI와 분리해 두면 리눅스 pwsh에서도 테스트 가능 |

---

## 3. 아키텍처

### 3.1 레이어

```
UI (WinForms, UI 스레드)
  MainForm · KPI 타일 · 그리드 · 상세패널 · 탭(내장/플러그인) · 트레이 아이콘
        ▲ 이벤트: SnapshotUpdated / StatusEvents / RefreshFailed
        │                                   ▼ 명령: Refresh, OpenItem, Ack, Reset ...
Services
  RefreshService(백그라운드 조회) · AlertWatcher(이전/현재 비교) · HistoryStore · Notifier
        ▲
Domain (순수 로직, UI 의존 없음 → Pester 테스트 대상)
  필드 정규화 · SRN 조인 · 상태 판정 · 정합성 점검 · 에러그룹 매핑 · 스냅샷 diff
        ▲
Data Provider (교체 가능)
  RestProvider(DefaultCredentials) | PnPProvider(SPO) | MockProvider(샘플 JSON, 오프라인 개발용)
```

데이터 흐름은 한 방향입니다.
1. `RefreshService`가 백그라운드에서 **Snapshot**(두 리스트 조회 → 정규화 → 조인 → 상태판정 → 정합성 점검)을 만듭니다.
2. UI 스레드가 받아서 `SnapshotUpdated` 이벤트를 발행합니다.
3. KPI·그리드·탭·플러그인은 이 이벤트를 구독해서 각자 화면을 갱신합니다.

### 3.2 폴더 구조

```
E-DashBoard/
├─ Start-EDashboard.cmd        더블클릭 실행기 (ExecutionPolicy Bypass, 콘솔 숨김)
├─ Start-EDashboard.ps1        진입점: 단일 인스턴스, 설정·모듈·플러그인 로드, 폼 실행
├─ config/
│  ├─ settings.psd1            사이트/리스트 URL, 주기, 임계값, 인증 방식 (팀 공용)
│  ├─ fields.psd1              내부명 ↔ 표시명 ↔ 그리드 열 정의
│  └─ errorcodes.psd1          에러코드 → 그룹/심각도/대시보드 컬럼
├─ src/
│  ├─ EDash.Data.psm1          Provider + 정규화 + 조인 + 상태판정 + 정합성 (UI 의존 없음)
│  ├─ EDash.Core.ps1           설정 병합, 로거, 이벤트 버스, 플러그인 로더
│  ├─ EDash.Services.ps1       RefreshService, AlertWatcher, HistoryStore, Notifier
│  └─ UI/
│     ├─ Theme.ps1             색상·글꼴 (주간/야간/키오스크)
│     ├─ MainForm.ps1          셸: 헤더, KPI, 탭, 상태바, 트레이
│     ├─ Controls.ps1          New-KpiTile, New-FilterBar, New-StatusGrid, New-DetailPanel
│     └─ Tabs/                 Systems.ps1, EAlert.ps1, Health.ps1, History.ps1, Settings.ps1
├─ plugins/                    드롭인 확장 (*.ps1 → 기능 정의 hashtable 반환)
├─ samples/                    Mock용 샘플 JSON (두 리스트)
├─ tests/                      Pester: 상태판정 · SRN 정규화 · 에러그룹 · diff
└─ docs/
```

런타임 데이터(로그, 이력, 캐시, 개인 설정)는 `%LOCALAPPDATA%\EDashboard\`에 저장합니다. 그래서 스크립트 폴더는 읽기 전용 네트워크 공유에 둬도 됩니다.

### 3.3 설정 파일

**settings.psd1** (팀 공용, 주석 가능)

```powershell
@{
    SiteUrl        = 'https://share.philips.com/sites/EASalarmhandling'
    Provider       = 'Rest'          # Rest | PnP | Mock
    Lists = @{
        Dashboard  = '/sites/EASalarmhandling/Lists/EasDashboard'
        EAlert     = '/sites/EASalarmhandling/Lists/EASDB2'
    }
    RefreshSeconds = 60
    TimeZone       = 'Korea Standard Time'
    NightStartHour = 22               # Flow1 ComposeIsNight와 동일
    NightEndHour   = 8
    DailyReset     = '18:00'          # Flow2 실행 시각 (KST)
    StaleHours     = 6                # 에러가 이 시간 이상 갱신 없으면 'AllOk 누락 의심'
    EnableWrite    = $false           # P3 쓰기 기능 스위치
}
```

**fields.psd1**: 내부명 매핑과 그리드 열 정의를 한 파일에서 관리합니다. 열을 추가할 때는 한 줄만 추가하면 됩니다.

```powershell
@{
    EasDashboard = @(
        @{ Name='SRN';                       Internal='SRN';                  Type='Text';  Header='SRN';        Width=90  }
        @{ Name='SiteName';                  Internal='field_2';              Type='Text';  Header='사이트';      Width=160 }
        @{ Name='SystemName';                Internal='field_3';              Type='Text';  Header='시스템';      Width=140 }
        @{ Name='FSE';                       Internal='field_4';              Type='Text';  Header='FSE';        Width=80  }
        @{ Name='Region';                    Internal='field_5';              Type='Text';  Header='Region';     Width=70  }
        @{ Name='Err7_SupplyWaterHigh';      Internal='field_6';              Type='Text';  Header='7 공급수';    Width=60  }
        @{ Name='Err9_ReturnWaterHigh';      Internal='field_7';              Type='Text';  Header='9 회수수';    Width=60  }
        @{ Name='Err11_DeltaTempHigh';       Internal='field_8';              Type='Text';  Header='11 ΔT';      Width=60  }
        @{ Name='Err13_ColdheadGasTempHigh'; Internal='field_9';              Type='Text';  Header='13 Coldhead'; Width=70 }
        @{ Name='Err20_Reminder';            Internal='field_10';             Type='Text';  Header='20 Rem';     Width=60  }
        @{ Name='LastUpdated';               Internal='field_14';             Type='Date';  Header='갱신';        Width=110 }
        @{ Name='AllOkTime';                 Internal='field_16';             Type='Date';  Header='AllOk 시각';  Width=110 }
        @{ Name='AllOk';                     Internal='AllOk';                Type='Bool';  Header='AllOk';      Width=50  }
        @{ Name='StartMonitoring';           Internal='StartMonitoring';      Type='Date';  Header='감시 시작';   Width=110 }
        @{ Name='ActiveErrorCount';          Internal='ActiveErrorCount';     Type='Int';   Header='활성';        Width=45  }
        @{ Name='ReminderCount';             Internal='ReminderCount';        Type='Int';   Header='Rem수';       Width=45  }
        @{ Name='LastUpdatedErrorTime';      Internal='LastUpdatedErrorTime'; Type='Date';  Header='최근 에러';   Width=110 }
    )
    EASDB2 = @(
        @{ Name='SRN';       Internal='SystemReferenceNumber'; Type='Text'; Header='SRN';      Width=90 }
        @{ Name='AlertSent'; Internal='AlertSent';             Type='Bool'; Header='알림발송'; Width=60 }
        # 나머지 컬럼은 확인 후 추가
    )
}
```

`Type`(Text/Date/Bool/Int)은 실제 컬럼 형식을 확인한 뒤 맞춥니다. `Bool` 변환기는 `true/'true'/'Yes'/1`을 모두 참으로 처리합니다. Flow가 AlertSent를 `'true'`/`'false'` **문자열**로 비교하고 있기 때문입니다.

**errorcodes.psd1**: Flow1의 `ComposeErrorGroup`과 같은 규칙을 한 곳에서 관리합니다.

```powershell
@{
    Groups = @(
        @{ Group='복구';         Codes=@(0);                    Severity='Ok'       }
        @{ Group='Reminder';     Codes=@(20);                   Severity='Warning'  }
        @{ Group='전원';         Codes=@(15);                   Severity='Critical' }
        @{ Group='냉각수';       Codes=@(7,8,9,10,11,12);       Severity='Critical' }
        @{ Group='Coldhead';     Codes=@(13,14);                Severity='Critical' }
        @{ Group='기계실온습도'; Codes=@(3,4,5,6);              Severity='Warning'  }
        @{ Group='센서고장';     Codes=@(81,82,83,84,85,86);    Severity='Warning'  }
    )
    DefaultGroup = '기타'

    # EasDashboard 에러 컬럼 ↔ 코드 (SwitchErrorCode 기준 추정, 확인 필요)
    Columns = @(
        @{ Field='Err7_SupplyWaterHigh';      Codes=@(7,8);   Group='냉각수';   Severity='Critical' }
        @{ Field='Err9_ReturnWaterHigh';      Codes=@(9,10);  Group='냉각수';   Severity='Critical' }
        @{ Field='Err11_DeltaTempHigh';       Codes=@(11,12); Group='냉각수';   Severity='Critical' }
        @{ Field='Err13_ColdheadGasTempHigh'; Codes=@(13,14); Group='Coldhead'; Severity='Critical' }
        @{ Field='Err20_Reminder';            Codes=@(20);    Group='Reminder'; Severity='Warning'  }
    )
}
```

---

## 4. 데이터 모델

### 4.1 정규화된 시스템 객체 (Snapshot 안의 한 행)

```
SrnKey, SRN, SiteName, SystemName, FSE, Region
Err7_SupplyWaterHigh, Err9_ReturnWaterHigh, Err11_DeltaTempHigh, Err13_ColdheadGasTempHigh, Err20_Reminder
ActiveErrorCount, ReminderCount, AllOk(bool), AllOkTime, LastUpdated, LastUpdatedErrorTime, StartMonitoring  (KST)
DashboardId              EasDashboard 항목 ID
EAlertId, AlertSent      EASDB2 조인 결과 (없으면 $null)
--- 계산 필드 ---
Status        Critical | Reminder | Ok | Unknown
Groups        활성 에러 그룹 목록 (냉각수, Coldhead ...)
ActiveCodes   값이 있는 에러 컬럼 목록
IsStale       에러 상태인데 LastUpdatedErrorTime이 StaleHours 이상 지남 → AllOk 누락 의심
ErrorAge      지금 − (StartMonitoring ?? LastUpdatedErrorTime)
```

**Snapshot** = `@{ TakenAt; Systems; Alerts(EASDB2 원본 정규화); Health(정합성 결과); DurationMs; Source }`

### 4.2 상태 판정 규칙 (설정 기반, Pester로 테스트)

| 순서 | 조건 | Status |
|---|---|---|
| 1 | `AllOk = true` 이고 `AllOkTime ≥ LastUpdatedErrorTime` | **Ok** |
| 2 | Severity=Critical인 에러 컬럼 중 하나라도 값 있음 | **Critical** |
| 3 | Err20_Reminder 값 있음 (또는 ReminderCount > 0) | **Reminder** |
| 4 | 에러 컬럼이 모두 비어 있음 | **Ok** |
| 5 | EASDB2에만 있고 EasDashboard에 없음 | **Unknown** |

- `IsStale`은 Status와 별개인 **표시 플래그**입니다. 예를 들어 "위험 · 7시간 갱신 없음"처럼 붙여서 보여줍니다. AllOk가 안 올 수 있다는 문제를 18:00 리셋 전에 미리 알아볼 수 있습니다.
- 1번에서 `AllOk`만 보지 않고 **시각을 비교**하는 이유가 있습니다. AllOk가 남아 있는 상태에서 새 에러가 들어온 경우에도 위험으로 판정하기 위해서입니다.

### 4.3 시간 처리

- SharePoint REST는 날짜를 UTC(`...Z`)로 줍니다. 그래서 `ConvertTo-Kst` 하나로 통일합니다. 입력은 문자열과 DateTime을 모두 받고, `[TimeZoneInfo]::ConvertTimeBySystemTimeZoneId(..., 'Korea Standard Time')`로 변환합니다.
- 이렇게 하면 PC 시간대 설정과 상관없이 KST로 표시됩니다. 주/야간 판정(22~08시)과 18:00 리셋 카운트다운도 같은 함수를 씁니다.

---

## 5. 화면 구성

### 5.1 메인 화면 와이어프레임

```
┌ 헤더 ─────────────────────────────────────────────────────────────────────
│ E-Dashboard   ● 연결됨 · 갱신 14:32:10 KST · [주간] · 18:00 리셋까지 3:27:50   [⟳ F5] [⚙]
├ KPI 타일 (클릭하면 해당 조건으로 필터) ───────────────────────────────────
│ [전체 152] [위험 4] [Reminder 7] [AllOk 누락의심 2] [정상 139] [Teams 발송 5]
├ 탭 ───────────────────────────────────────────────────────────────────────
│ (시스템) (E-Alert) (정합성 점검) (이력) (설정)  + 플러그인 탭들
│
│ 검색 [__________]  Region [전체▼]  FSE [전체▼]  그룹 [전체▼]  [v] 에러만  [열 설정]
│ ┌──────────────────────────────────────────────────────────────────────
│ │ 상태 │ SRN   │ 사이트   │ 시스템  │ FSE │ Rgn │ 7 │ 9 │ 11 │ 13 │ 20 │ 활성 │ Rem │ 최근 에러 │ 경과  │ 발송
│ │ 위험 │ 12345 │ OO병원   │ MR 1.5T │ ... │ 서울 │ ● │   │    │    │    │  1   │  0  │ 14:05    │ 27분  │ ●
│ │ Rem  │ 23456 │ △△의원  │ MR 3T   │ ... │ 부산 │   │   │    │    │ ●  │  0   │  3  │ 09:40    │ 4시간 │
│ │ ...
│ ├ 상세 (선택 행) ──────────────────────────────────────────────────────
│ │ [EasDashboard 원본] [EASDB2 원본] [이 SRN 이벤트 이력]
│ └──────────────────────────────────────────────────────────────────────
├ 상태바 ───────────────────────────────────────────────────────────────────
│ 152건 · 조회 0.8s · 다음 갱신 45s · 로그 열기
```

### 5.2 WinForms 컨트롤 구성

| 영역 | 컨트롤 | 메모 |
|---|---|---|
| 전체 레이아웃 | `TableLayoutPanel` (행: 헤더/KPI/탭/상태바) | Dock 순서 문제를 피하기 위해 사용 |
| KPI | `FlowLayoutPanel` + 타일(Panel + Label 2개) | 클릭 이벤트로 필터 적용 |
| 시스템 탭 | `SplitContainer`(가로 분할) → 위: 필터바 + `DataGridView`, 아래: 상세 `TabControl` | |
| 그리드 | `DataGridView` ← `DataTable`(PrimaryKey=SrnKey) ← `DataView.RowFilter` | 갱신할 때 행 단위로 병합해서 깜빡임과 선택·스크롤 초기화를 막음 |
| 색상 | `CellFormatting` 이벤트에서 Status별 배경색 지정 | `DoubleBuffered`를 리플렉션으로 켬 |
| 행 메뉴 | `ContextMenuStrip` | SharePoint에서 열기, SRN 복사, (P3) 확인 처리, 수동 리셋 |
| 트레이 | `NotifyIcon` + 풍선 알림 | 최소화하면 트레이로 |
| 타이머 | 갱신 60s / 결과 확인 500ms / 시계 1s | 모두 `System.Windows.Forms.Timer` (UI 스레드) |

### 5.3 탭별 기능

| 탭 | 내용 |
|---|---|
| **시스템** | 위 그리드. 정렬, 필터, 열 숨김·순서·너비 저장. 더블클릭하면 SharePoint `DispForm.aspx?ID=`을 엶 |
| **E-Alert** | EASDB2 목록과 대시보드 조인 결과. AlertSent=true 항목, 오늘 Teams 발송 이력 |
| **정합성 점검** | 중복 SRN, 리스트 간 미매칭, `AlertSent=true`인데 정상인 항목, `ActiveErrorCount`와 실제 에러 컬럼 수 불일치, AllOk 누락 의심 |
| **이력** | 로컬 이벤트 로그(에러 발생/복구/리셋)와 날짜별 스냅샷 조회, CSV 내보내기 |
| **설정** | 갱신 주기, 알림 옵션, 내 Region/FSE 필터, 테마, Provider 연결 테스트 |

---

## 6. 기능 카탈로그 (단계별)

| 분류 | 기능 | 단계 | 필요조건 |
|---|---|---|---|
| 조회 | KPI 타일 (클릭하면 필터) | P1 | – |
| 조회 | 시스템 그리드 + 상태 색상 + 정렬 | P1 | – |
| 조회 | 검색·필터 (Region/FSE/그룹/상태/에러만) | P1 | – |
| 조회 | 상세 패널 (두 리스트 원본 + 경과 시간) | P1 | – |
| 조회 | 자동 새로고침, 연결 상태, F5 수동 갱신, 실패 시 재시도(백오프) | P1 | – |
| 조회 | 주/야간 표시, 18:00 리셋 카운트다운, "오늘 리셋 대상 N건" | P1 | – |
| 조회 | 항목 더블클릭 → SharePoint에서 열기 | P1 | – |
| 운영 | 단일 인스턴스, 로그 파일, 창 위치·열 설정 저장 | P1 | – |
| 개발 | Mock Provider (샘플 JSON으로 오프라인 실행) | P1 | – |
| 알림 | 변화 감지 이벤트 (신규 위험, 복구, Teams 발송, AllOk 누락 의심) | P2 | – |
| 알림 | 트레이 풍선 알림, 야간 크리티컬 경보음 | P2 | – |
| 알림 | 18:00 일괄 리셋 인식 (복구 알림을 여러 건 보내지 않고 "일괄 리셋 1건"으로 기록) | P2 | – |
| 알림 | 알림 범위 설정 (내 Region / 내 FSE만) | P2 | – |
| 이력 | 이벤트 로그 (CSV, 일자별 파일) | P2 | 대시보드 실행 중에만 기록됨 |
| 이력 | 17:55 리셋 직전 스냅샷 저장 → 일일 요약 | P2 | 〃 |
| 이력 | CSV / Excel 내보내기 | P2 | – |
| 점검 | 정합성 점검 탭 | P2 | – |
| 화면 | 키오스크(관제 모니터) 모드: 전체화면, 큰 글꼴, 에러만, 다크 테마 | P2 | – |
| 조치 | 확인(Ack) 처리 + 메모 | P3 | EasDashboard에 컬럼 추가 (AckBy/AckTime/Memo) |
| 조치 | SRN 단위 수동 리셋 (Flow2와 같은 필드) | P3 | `EnableWrite`, 편집 권한 |
| 조치 | AlertSent 수동 해제 | P3 | 〃 |
| 조치 | FSE / Region 수정 | P3 | 〃 |
| 조치 | 감사 로그 (누가·언제·무엇을) | P3 | – |
| 연계 | FSE에게 메일 (Outlook COM / mailto, 본문 자동 작성) | P3 | – |
| 연계 | Teams 채팅 딥링크 | P3 | – |
| 분석 | 차트: 일별·그룹별 건수, Top N 시스템, 시간대 분포 | P4 | 이력 데이터 |
| 분석 | SharePoint 이력 리스트 연동 (10장 4번) | P4 | Flow1에 액션 1개 추가 |
| 연계 | 알림 수신자 리스트(GetAlertMail, email1~5) 조회·편집 | P4 | 세 번째 리스트 |
| 배포 | 코드 서명 / ps2exe 패키징 / 자동 업데이트(공유폴더 버전 확인) | P4 | – |

---

## 7. 확장 구조 (플러그인 + 이벤트 버스)

### 7.1 이벤트 버스

UI 스레드에서만 발행합니다. 그래서 구독자(플러그인)는 컨트롤을 바로 건드려도 안전합니다.

| 이벤트 | 데이터 | 발행 시점 |
|---|---|---|
| `SnapshotUpdated` | Snapshot | 갱신이 성공할 때마다 |
| `StatusEvents` | 이벤트 배열 (`BecameCritical`, `BecameReminder`, `Recovered`, `AlertSent`, `StaleDetected`, `DailyReset`, `NewSystem`) | 이전 스냅샷과 비교해서 변화가 있을 때 |
| `RefreshFailed` / `ConnectionRestored` | 예외 / 시각 | 연결 상태가 바뀔 때 |
| `SelectionChanged` | 선택된 시스템 | 그리드에서 다른 행을 선택할 때 |
| `NightModeChanged` | bool | 22:00, 08:00 |

### 7.2 플러그인 계약

`plugins\*.ps1` 파일이 아래 hashtable을 **반환**하면 로더가 등록합니다. 내장 탭(시스템/E-Alert/정합성/이력)도 같은 방식으로 만들어서 구조를 하나로 통일합니다.

```powershell
@{
    Id        = 'night-siren'                 # 고유 ID (설정에서 켜기/끄기)
    Name      = '야간 크리티컬 경보음'
    Order     = 100                           # 탭/메뉴 순서

    Init      = { param($ctx) }               # 1회 실행 (선택)
    CreateTab = $null                         # { param($ctx) ... return [Windows.Forms.Control] } (선택)
    RowMenu   = @()                           # @{ Text='...'; Action={ param($ctx, $system) } } (선택)

    On = @{
        StatusEvents = {
            param($ctx, $events)
            if (-not $ctx.IsNight()) { return }
            if ($events | Where-Object { $_.Type -eq 'BecameCritical' }) {
                [System.Media.SystemSounds]::Exclamation.Play()
                $ctx.Notify('야간 크리티컬', ($events.SRN -join ', '), 'Error')
            }
        }
    }
}
```

플러그인에 넘기는 `$ctx`:

| 멤버 | 설명 |
|---|---|
| `Config`, `Theme` | 병합된 설정, 색상·글꼴 |
| `Snapshot()` / `Selected()` | 최신 스냅샷, 현재 선택 행 |
| `Notify(title, text, level)` | 트레이 알림 |
| `Log(msg, level)` | 로그 |
| `IsNight()`, `NowKst()` | 시간 유틸 |
| `OpenItem(system)` | SharePoint에서 열기 |
| `Subscribe(event, sb)` / `Publish(event, data)` | 이벤트 버스 |
| `Sp` | 데이터 접근. `EnableWrite = $false`이면 쓰기 메서드가 예외를 던짐 |

새 기능은 보통 **plugins 폴더에 파일 하나만 추가**해서 붙입니다. 코어 수정이 필요한 경우는 Snapshot에 새 데이터 소스(예: 세 번째 리스트)를 넣을 때 정도입니다.

---

## 8. 핵심 코드 스케치 (PS 5.1 문법)

### 8.1 리스트 전체 조회 (페이징)

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Invoke-EDashSpGet {
    param(
        [Parameter(Mandatory)] [string] $SiteUrl,
        [Parameter(Mandatory)] [string] $ListUrl,   # '/sites/EASalarmhandling/Lists/EasDashboard'
        [string[]] $Select,
        [string]   $Filter
    )
    $query = "@a1='$ListUrl'&`$top=5000"
    if ($Select) { $query += '&$select=' + ($Select -join ',') }
    if ($Filter) { $query += '&$filter=' + [uri]::EscapeDataString($Filter) }
    $url     = "$SiteUrl/_api/web/GetList(@a1)/items?$query"
    $headers = @{ Accept = 'application/json;odata=nometadata' }

    $items = New-Object System.Collections.Generic.List[object]
    while ($url) {
        $resp = Invoke-RestMethod -Uri $url -Headers $headers -UseDefaultCredentials -UseBasicParsing
        foreach ($v in $resp.value) { $items.Add($v) }
        $url = $resp.'odata.nextLink'
    }
    $items.ToArray()
}
```

- `$select`에는 `fields.psd1`의 `Internal` 목록과 `ID`, `Modified`를 자동으로 넣습니다.
- FSE(`field_4`)가 **사용자(Person) 컬럼**이면 `$select=field_4/Title&$expand=field_4`가 필요합니다. 형식 확인이 필요합니다.

### 8.2 백그라운드 갱신 (UI 멈춤 방지)

```powershell
# 초기화: 데이터 모듈을 미리 로드한 전용 Runspace
$iss = [initialsessionstate]::CreateDefault()
$iss.ImportPSModule("$Root\src\EDash.Data.psm1")
$script:Rs = [runspacefactory]::CreateRunspace($iss)
$script:Rs.Open()
$script:Job = $null

function Start-EDashRefresh {
    if ($script:Job) { return }                      # 진행 중이면 무시 (중복 조회 방지)
    $ps = [powershell]::Create()
    $ps.Runspace = $script:Rs
    [void]$ps.AddCommand('Get-EDashSnapshot').AddParameter('Config', $script:Config)
    $script:Job = @{ PS = $ps; Handle = $ps.BeginInvoke() }
}

# 500ms UI 타이머: 완료 여부만 확인하고, 결과 반영은 UI 스레드에서 처리
$pollTimer.Add_Tick({
    if (-not $script:Job -or -not $script:Job.Handle.IsCompleted) { return }
    $job = $script:Job; $script:Job = $null
    try {
        $snap = $job.PS.EndInvoke($job.Handle) | Select-Object -Last 1
        Publish-EDashEvent 'SnapshotUpdated' $snap
    } catch {
        Publish-EDashEvent 'RefreshFailed' $_
    } finally {
        $job.PS.Dispose()
    }
})
```

`Get-EDashSnapshot`은 두 리스트 조회 → 정규화 → SrnKey 조인 → 상태판정 → 정합성 점검까지 하는 **순수 함수**입니다. UI와 관계가 없어서 Pester로 테스트하고, Mock Provider로 오프라인에서 돌릴 수 있습니다.

### 8.3 그리드 갱신 (선택·스크롤 유지)과 색상

```powershell
function Update-EDashGrid($grid, [System.Data.DataTable]$table, $snapshot) {
    $selKey = $null
    if ($grid.CurrentRow) { $selKey = $grid.CurrentRow.Cells['SrnKey'].Value }
    $firstRow = $grid.FirstDisplayedScrollingRowIndex

    $table.BeginLoadData()
    $seen = @{}
    foreach ($s in $snapshot.Systems) {
        [void]$table.LoadDataRow(@($s.SrnKey, $s.Status, $s.SRN, $s.SiteName <# ... 열 순서대로 #>), $true)
        $seen[$s.SrnKey] = $true
    }
    foreach ($r in @($table.Rows)) { if (-not $seen[$r['SrnKey']]) { $r.Delete() } }
    $table.AcceptChanges()
    $table.EndLoadData()

    # 선택·스크롤 복원 (생략: SrnKey로 행 찾아서 CurrentCell 지정, FirstDisplayedScrollingRowIndex 복원)
}

$grid.Add_CellFormatting({
    param($s, $e)
    if ($e.RowIndex -lt 0) { return }
    $row = $s.Rows[$e.RowIndex].DataBoundItem.Row
    switch ($row['Status']) {
        'Critical' { $e.CellStyle.BackColor = $Theme.CriticalBack; $e.CellStyle.ForeColor = $Theme.CriticalFore }
        'Reminder' { $e.CellStyle.BackColor = $Theme.WarningBack }
        'Unknown'  { $e.CellStyle.BackColor = $Theme.UnknownBack }
    }
    if ($row['IsStale']) { $e.CellStyle.Font = $Theme.ItalicFont }
})
```

### 8.4 쓰기 (P3, `EnableWrite`일 때만)

```powershell
function Set-EDashSpItem {
    param($SiteUrl, $ListUrl, [int]$Id, [hashtable]$Fields, [string]$ETag = '*')
    $h = @{ Accept = 'application/json;odata=nometadata' }
    $digest = (Invoke-RestMethod "$SiteUrl/_api/contextinfo" -Method Post -Headers $h -UseDefaultCredentials).FormDigestValue
    $type   = (Invoke-RestMethod "$SiteUrl/_api/web/GetList(@a1)?@a1='$ListUrl'&`$select=ListItemEntityTypeFullName" `
                 -Headers $h -UseDefaultCredentials).ListItemEntityTypeFullName

    $body = (@{ __metadata = @{ type = $type } } + $Fields) | ConvertTo-Json -Compress
    Invoke-RestMethod "$SiteUrl/_api/web/GetList(@a1)/items($Id)?@a1='$ListUrl'" -Method Post -UseDefaultCredentials `
        -ContentType 'application/json;odata=verbose;charset=utf-8' `
        -Headers ($h + @{ 'X-RequestDigest' = $digest; 'IF-MATCH' = $ETag; 'X-HTTP-Method' = 'MERGE' }) `
        -Body ([Text.Encoding]::UTF8.GetBytes($body))      # PS 5.1 한글 깨짐 방지: 바이트로 전송
}
```

- `$Fields`의 키는 **내부명**(`field_6` 등)이어야 합니다. 표시명은 fields.psd1로 변환합니다.
- `IF-MATCH`에 `*` 대신 조회 시점의 ETag를 넣으면, 그 사이 Flow가 바꾼 항목을 덮어쓰지 않습니다(낙관적 동시성). ETag는 `odata=minimalmetadata`로 조회하면 `odata.etag`로 받을 수 있습니다.
- 모든 쓰기는 확인 대화상자를 거치고 감사 로그에 남깁니다.

---

## 9. 운영·배포 시 주의사항

| 항목 | 내용 |
|---|---|
| **인코딩** | `.ps1`/`.psd1`은 반드시 **UTF-8 with BOM**으로 저장해야 합니다. PS 5.1은 BOM이 없으면 한글이 깨집니다. JSON은 `Get-Content -Raw -Encoding UTF8`로 읽습니다 |
| **실행** | `Start-EDashboard.cmd`: `start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0Start-EDashboard.ps1"` |
| **실행 정책** | GPO(MachinePolicy)로 막혀 있으면 Bypass가 통하지 않습니다. 이 경우 코드 서명이 필요합니다 |
| **단일 인스턴스** | `System.Threading.Mutex`(`Global\EDashboard`)로 중복 실행 방지. 이미 실행 중이면 기존 창을 활성화 |
| **DPI** | `SetProcessDPIAware`(user32) + `AutoScaleMode = Dpi`. 한글은 `Malgun Gothic` 9~10pt |
| **오류 처리** | 조회 실패 시 마지막 스냅샷을 유지하고 헤더를 "● 연결 끊김 (마지막 성공 14:31)"으로 표시. 재시도 간격은 60s → 120s → 최대 300s |
| **호출 부하** | 60초마다 2회 GET, `$select`로 필요한 컬럼만. 수백 건 규모면 전체 조회로 충분하고, 커지면 `Modified gt` 증분 조회로 바꿈 |
| **로그** | `%LOCALAPPDATA%\EDashboard\logs\yyyyMMdd.log`, 7일 보관 |
| **보안** | 자격증명을 저장하지 않음 (Windows 통합인증). 쓰기 기능은 설정 스위치와 확인 대화상자를 거치고 감사 로그에 남김 |

---

## 10. Flow 쪽 참고 (대시보드 설계에 영향 있는 부분)

1. **중복 SRN 행 가능성**: "When a new email arrives (V3)"는 기본적으로 메일마다 병렬 실행됩니다. 같은 SRN 메일 2통이 거의 동시에 오면 둘 다 `GetEasDashboard` 결과가 0건이라서 `CreateEasDashboardItem`이 2번 실행될 수 있습니다. **트리거 설정 → 동시성 제어 켜기 → 병렬 처리 수 1**을 권장합니다. 대시보드의 정합성 탭에서도 중복을 잡아냅니다.
2. **SRN 정규화 불일치**: `GetEasDashboard`는 `replace(ComposeSRN, '"', '')`로 따옴표를 지우지만, `GetEAlertMain`은 `ComposeSRN`을 그대로 씁니다. SRN에 따옴표가 섞이면 두 리스트의 키가 서로 달라집니다. `ComposeSRN` 단계에서 한 번에 정규화하는 것을 권장합니다. 대시보드는 `SrnKey`로 조인해서 이 경우도 흡수합니다.
3. `ComposeSRNCut`의 `indexOf(outputs('ComposeSRNFrom'), '')`는 붙여넣는 과정에서 HTML 태그가 빠진 것으로 보입니다. 실제로 빈 문자열이면 SRN이 항상 빈 값이 되니, 원본 식만 한 번 확인해 주세요.
4. **이력 보존**: 18:00 리셋이 상태를 지우기 때문에 "언제, 어떤 에러가, 얼마나 지속됐는지"가 남지 않습니다. 대시보드의 로컬 이력은 PC가 켜져 있는 동안만 쌓입니다. 완전한 통계가 필요하면 Flow1에 "이력 리스트에 1행 추가" 액션 하나만 넣으면 되고, 대시보드는 그 리스트를 읽기만 하면 됩니다 (P4 옵션).
5. **AlertSent 형식**: Flow는 `'true'`/`'false'` 문자열로 비교합니다. 이 컬럼을 Yes/No로 바꾸면 Flow 식이 깨집니다. 대시보드는 두 형식을 모두 처리하도록 만듭니다.

---

## 11. 로드맵

| 단계 | 범위 | 완료 기준 |
|---|---|---|
| **P0 환경 확인** | 12장 확인 사항, REST 테스트 | 두 리스트 JSON 샘플 확보 → `samples/`에 (민감정보 마스킹) |
| **P1 조회 MVP** | 셸, Data 모듈, Mock/Rest Provider, KPI, 그리드, 필터, 상세, 자동 갱신, 설정 파일, Pester | 실제 리스트로 60초 갱신, UI 멈춤 없음 |
| **P2 모니터링** | AlertWatcher, 트레이 알림, 야간 경보, 이력 CSV, 스냅샷, 내보내기, 정합성 탭, 키오스크 모드 | 신규 위험 발생 시 1분 이내 알림 |
| **P3 조치** | 쓰기 계층, Ack/메모, 수동 리셋, AlertSent 해제, 메일/Teams 연계, 감사 로그 | 권한 있는 사용자만, 확인 후 반영 |
| **P4 분석·연계** | 차트, 이력 리스트 연동, 수신자 리스트, 패키징/자동 업데이트 | – |

---

## 12. 착수 전 확인 사항

1. **SharePoint 종류와 인증 확인.** 회사 PC의 PowerShell 5.1에서 아래를 실행해서 결과를 공유해 주세요(이름 등은 마스킹).
   - 성공하면 온프레미스 + 통합인증 방식으로 확정합니다.
   - 401/403이 나면 SharePoint Online(PnP) 방식으로 갑니다.
   - 결과에 실제 필드 형식(텍스트/날짜/Yes-No/사용자)도 같이 나와서 아래 2~4번이 함께 풀립니다.

   ```powershell
   $PSVersionTable.PSVersion
   $site = 'https://share.philips.com/sites/EASalarmhandling'
   foreach ($l in 'EasDashboard','EASDB2') {
       Invoke-RestMethod "$site/_api/web/GetList(@a1)/items?@a1='/sites/EASalarmhandling/Lists/$l'&`$top=2" `
           -UseDefaultCredentials -UseBasicParsing -Headers @{ Accept = 'application/json;odata=nometadata' } |
           ConvertTo-Json -Depth 5
   }
   ```

2. **EASDB2 컬럼 목록(내부명).** 지금은 `SystemReferenceNumber`, `AlertSent`만 알고 있습니다.
3. **에러 컬럼(field_6 ~ field_10)의 값 형식**(시각? 'Yes'? 횟수?)과, `SwitchErrorCode`에서 **8·10·12·14 코드가 어느 컬럼을 갱신하는지**.
4. **case 0(복구)** 에서 AllOk/AllOkTime/에러 컬럼/ActiveErrorCount가 어떻게 바뀌는지.
5. Flow2의 **`GetDashboardNotOk` 필터 식**, 그리고 `ResetDashboardItem`이 비우는 필드 목록.
6. **사용 형태**: 개인 PC 1대인지, 관제 모니터인지, FSE 여러 명에게 배포하는지. 그리고 쓰기(P3) 기능이 필요한지.
7. 회사 PC의 **스크립트 실행 정책**(GPO 차단 여부).
