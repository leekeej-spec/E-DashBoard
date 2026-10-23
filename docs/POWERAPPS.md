# E-Dashboard Power Apps 버전 만들기

SharePoint 리스트 `EasDashboard`, `EASDB2`(표시 이름 **EasTeamsAlarm**)를 **Power Apps 캔버스 앱**이 직접 읽어서 보여주는 방식입니다.

## 0. 무엇이 달라지나

| | PowerShell 방식 (이전) | Power Apps 방식 (지금) |
|---|---|---|
| 인증 | 토큰 문제 → 내보내기 Flow + OneDrive 파일로 우회 | 각자 **자기 계정**으로 앱을 열면 끝 (SharePoint 권한 그대로) |
| 필요한 것 | Flow, 파일, 동기화, 스크립트 실행 | 앱 하나 |
| 배포 | 파일 복사, 실행 정책 | 앱 **공유** 또는 Teams 탭 |
| 라이선스 | – | SharePoint 커넥터는 **표준** 커넥터라 M365 라이선스로 충분 (프리미엄 불필요) |
| 데이터 지연 | 최대 4~5분 | **60초** (앱이 직접 읽음) |

필요 없어지는 것: 데이터 내보내기 Flow, `edash-data.json`, OneDrive/문서 라이브러리 동기화, `EDashboard.ps1`

---

## 1. 시작 전에: 컬럼 형식 확인 (2분)

Power Apps 수식은 컬럼 형식(텍스트/날짜/선택 등)에 따라 문법이 다릅니다. 아래 수식은 **가장 흔한 형식**을 기준으로 작성했습니다.

- **형식에 상관없이 동작하는 컬럼:** 에러 컬럼 5개, AllOk, AlertSent, 개수 컬럼 2개
- **형식을 확인해야 하는 컬럼:** 아래 표의 컬럼만

확인 방법: `EasDashboard` 리스트 → 톱니(설정) → **목록 설정** → "열" 표의 **형식**을 봅니다.

| 컬럼 (표시 이름) | 기준으로 잡은 형식 | 다르면 |
|---|---|---|
| SiteName, SystemName, FSE, REGION | 한 줄 텍스트 | 선택(Choice)·사용자(Person)면 [12장](#12-컬럼-형식이-다를-때) |
| LastUpdated, AllOkTime, StartMonitoring, LastUpdatedErrorTime | 날짜 및 시간 | 한 줄 텍스트면 [12장](#12-컬럼-형식이-다를-때) |

형식이 달라도 걱정할 필요는 없습니다. 앱 편집기가 해당 줄에 빨간 밑줄을 그어 주고, 12장의 한 줄만 바꾸면 됩니다.

---

## 2. 앱 만들기

1. https://make.powerapps.com → **+ 만들기 → 빈 앱 → 빈 캔버스 앱**
   - 이름: `E-Dashboard`
   - 형식: **태블릿**
2. **데이터 추가**(왼쪽 원통 아이콘) → **SharePoint** → 연결 선택 → 사이트 `EASalarmhandling`(`https://share.philips.com/sites/EASalarmhandling`) → **EasDashboard**, **EasTeamsAlarm** 체크 (URL 이름은 `EASDB2`) → 연결
3. **설정 → 일반 → 데이터 행 제한**: `500` → **`2000`**
   - 두 리스트를 통째로 읽기 때문에 필요합니다. 리스트가 2000건을 넘으면 앱이 경고합니다.
4. **클래식 컨트롤**을 씁니다 → [2-1](#2-1-클래식-컨트롤-고르기)
5. 컨트롤을 넣을 때마다 **이름을 바꿉니다** → [2-2](#2-2-컨트롤-이름-바꾸기)

### 2-1. 클래식 컨트롤 고르기

Power Apps에는 같은 이름의 컨트롤이 두 종류 있습니다. **클래식**과 **최신(모던)** 입니다. 둘은 속성 이름이 다릅니다. 예를 들어 검색창의 입력값이 클래식에서는 `.Text`, 최신에서는 `.Value`입니다. 이 가이드의 수식은 **클래식** 기준이라, 최신 컨트롤을 넣으면 빨간 밑줄이 생깁니다.

두 종류가 갈리는 컨트롤은 아래 5개뿐입니다. 갤러리, HTML 텍스트, 타이머, 사각형은 한 종류만 있어서 신경 쓰지 않아도 됩니다.

| 이 가이드의 종류 | 골라야 할 것 (클래식) | 고르면 안 되는 것 (최신) | 클래식인지 확인하는 방법 |
|---|---|---|---|
| 레이블 | **텍스트 레이블** (Text label) | 텍스트 (Text) | 속성 목록에 `Color`가 있음 (최신은 `FontColor`) |
| 단추 | **단추** (Button) | 단추 (Button) | 속성 목록에 `Fill`이 있음 |
| 텍스트 입력 | **텍스트 입력** (Text input) | 텍스트 입력 (Text input) | 속성 목록에 `Default`가 있음 (최신은 `Value`) |
| 콤보 상자 | **콤보 상자** (Combo box) | 콤보 상자 (Combobox) | 속성 목록에 `InputTextPlaceholder`가 있음 |
| 확인란 | **확인란** (Check box) | 확인란 (Checkbox) | 속성 목록에 `Text`가 있음 (최신은 `Label`) |

**방법 A (가장 쉬움): 최신 컨트롤을 끄기.** 이렇게 하면 삽입 메뉴에 클래식만 남습니다.
1. 위쪽 메뉴의 **설정**(톱니) → **업데이트** → **새로 만들기** 탭
2. **최신 컨트롤 및 테마** (Modern controls and themes) → **끔**
3. 설정 창을 닫습니다.
   - 이 항목이 보이지 않는 버전도 있습니다. 그때는 방법 B를 쓰세요.

**방법 B: 삽입 메뉴에서 "클래식" 그룹에서 고르기**
1. 왼쪽 세로 막대의 **+ (삽입)** 를 누릅니다.
2. 목록 아래쪽 **클래식** (Classic) 그룹을 펼칩니다.
3. 거기서 텍스트 레이블, 단추, 텍스트 입력, 콤보 상자, 확인란을 고릅니다.
   - 위쪽 "인기" 목록에 같은 이름이 있으면 그건 대개 최신 컨트롤입니다.

**확인하는 법:** 컨트롤을 선택하고 수식 입력줄 왼쪽의 **속성 드롭다운**을 엽니다. 위 표의 "확인하는 방법"에 적힌 속성이 있으면 클래식입니다. 잘못 넣었으면 Delete로 지우고 다시 넣으세요.

### 2-2. 컨트롤 이름 바꾸기

컨트롤을 넣으면 `Label1`, `Button1` 같은 이름이 자동으로 붙습니다. 수식은 다른 컨트롤을 **이름으로** 가리킵니다. 예를 들어 목록 필터는 `txtSearch.Text`로 검색창 내용을 읽습니다. 그래서 이름이 다르면 수식이 그 컨트롤을 찾지 못합니다.

**반드시 바꿔야 하는 이름은 5개입니다.** 다른 수식이 이 이름을 직접 쓰기 때문입니다.

| 이름 | 종류 | 이 이름을 쓰는 곳 |
|---|---|---|
| **`btnLoad`** | 단추 | 타이머, 화면 OnVisible (`Select(btnLoad)`) |
| **`txtSearch`** | 텍스트 입력 | 목록 필터, 필터 초기화 |
| **`cbRegion`** | 콤보 상자 | 목록 필터, 필터 초기화 |
| **`cbFse`** | 콤보 상자 | 목록 필터, 필터 초기화 |
| **`chkErrors`** | 확인란 | 목록 필터, 필터 초기화 |

나머지(`lblTitle`, `galSys`, `rectBar` 등)는 수식에서 참조하지 않으므로 기본 이름 그대로 둬도 동작합니다. 다만 나중에 찾기 쉽도록 가이드 이름으로 바꾸는 것을 권장합니다.

**바꾸는 방법 (셋 중 편한 것):**

1. **트리 보기에서 더블클릭**
   - 왼쪽 세로 막대 맨 위의 **트리 보기** (Tree view, 층 모양 아이콘) 를 엽니다.
   - 화면(`Screen1`) 아래에 컨트롤 목록이 보입니다. 바꿀 이름(예: `TextInput1`)을 **더블클릭**하면 글자를 고칠 수 있게 됩니다.
   - `txtSearch`를 입력하고 **Enter**를 누릅니다.
2. **트리 보기의 … 메뉴**
   - 컨트롤 이름 오른쪽의 **…** (추가 옵션) → **이름 바꾸기** (Rename) → 입력 → Enter
3. **오른쪽 속성 창 맨 위**
   - 컨트롤을 선택하면 오른쪽 속성 창 맨 위에 이름(예: `TextInput1`)이 보입니다.
   - 그 이름을 클릭해서 고치고 Enter를 누릅니다.

**주의할 점:**
- **대소문자까지 똑같이** 입력하세요. `btnLoad`와 `BtnLoad`는 다른 이름입니다.
- 이름은 앱 전체에서 하나만 있어야 합니다. 이미 쓰는 이름이면 Power Apps가 `txtSearch_1`처럼 바꿔 버리니, 트리 보기에서 확인하세요.
- **갤러리 안 컨트롤:** 트리 보기에서 갤러리 이름 왼쪽의 **>** 를 눌러 펼치면 그 안의 컨트롤(`Label2` 등)이 보입니다. 같은 방법으로 바꿉니다.
- **순서:** 컨트롤을 넣으면 **바로 이름부터 바꾸고**, 그다음에 수식을 넣는 것이 가장 편합니다.
  - 수식을 먼저 넣어도 괜찮습니다. 이름을 맞추는 순간 빨간 밑줄이 사라집니다.

### 2-3. 속성에 수식·값 넣는 법 (참고)

- **수식:** 컨트롤을 선택 → 수식 입력줄 왼쪽 **속성 드롭다운**에서 속성(예: `OnSelect`, `Items`, `Text`)을 고름 → 오른쪽 **fx** 칸에 가이드의 수식을 붙여넣기
- **위치·크기:** 오른쪽 속성 창의 **위치**(X, Y)와 **크기**(너비, 높이) 칸에 숫자를 넣습니다. 수식 입력줄에서 `X`, `Y`, `Width`, `Height` 속성을 골라 넣어도 됩니다.
- **화면·앱 속성:** 트리 보기 맨 위의 **App** 또는 **Screen1**을 선택하면 `OnStart`, `OnVisible` 같은 속성이 속성 드롭다운에 나옵니다.

---

## 3. App.OnStart

왼쪽 트리 뷰에서 **App** 선택 → 속성 `OnStart`:

<!-- fx:App.OnStart -->
```powerfx
Set(gStaleHours, 6);      // 에러가 이 시간 이상 갱신되지 않으면 "AllOk 누락 의심"
Set(gFilter, "All");
Set(gLoaded, false);
Set(gNow, Now());
Set(gLoadedAt, Now());
Set(gSelId, Blank())
```

---

## 4. 화면 배치

화면 크기는 1366 × 768(태블릿 기본)입니다. 아래 표대로 넣고 X, Y, Width, Height를 맞춥니다.

| 이름 (**굵은 것은 필수**) | 종류 | X | Y | Width | Height | 비고 |
|---|---|---|---|---|---|---|
| `rectHeader` | 사각형 | 0 | 0 | 1366 | 56 | Fill `RGBA(30, 45, 70, 1)` |
| `lblTitle` | 레이블 | 16 | 8 | 240 | 40 | Text `"E-Dashboard"`, Color `Color.White`, Size `20`, FontWeight `FontWeight.Bold` |
| `lblHeader` | 레이블 | 260 | 8 | 940 | 40 | Color `Color.White`, Size `12` (8장 수식) |
| **`btnLoad`** | 단추 | 1216 | 10 | 134 | 36 | Text `"새로고침"` (5장 수식) |
| `tmrRefresh` | 타이머 | 0 | 0 | 10 | 10 | (5장) |
| `galKpi` | 가로 갤러리 (빈) | 8 | 62 | 1350 | 92 | (6장) |
| **`txtSearch`** | 텍스트 입력 | 16 | 162 | 230 | 36 | Default `""`, HintText `"SRN · 사이트 · 시스템 · FSE 검색"` |
| **`cbRegion`** | 콤보 상자 | 256 | 162 | 170 | 36 | (7장) |
| **`cbFse`** | 콤보 상자 | 436 | 162 | 190 | 36 | (7장) |
| **`chkErrors`** | 확인란 | 640 | 162 | 110 | 36 | Text `"에러만"` |
| `btnClear` | 단추 | 760 | 162 | 120 | 36 | Text `"필터 초기화"` (7장) |
| `galSys` | 세로 갤러리 (빈) | 8 | 206 | 880 | 554 | (9장) |
| `galUnmatched` | 세로 갤러리 (빈) | 8 | 206 | 880 | 554 | (11장) |
| `rectDetail` | 사각형 | 896 | 206 | 462 | 554 | Fill `RGBA(246, 247, 249, 1)` |
| `htmlDetail` | HTML 텍스트 | 904 | 214 | 446 | 480 | (10장) |
| `btnOpen` | 단추 | 904 | 706 | 446 | 44 | Text `"SharePoint에서 열기"` (10장) |

---

## 5. 데이터 읽기: `btnLoad`, 타이머, 화면

### btnLoad.OnSelect

두 리스트를 읽어서 상태를 계산하는 핵심 수식입니다. 그대로 붙여넣으세요.

<!-- fx:btnLoad.OnSelect -->
```powerfx
// 1) 직전 위험 목록 (신규 위험 알림용)
ClearCollect(colPrevCrit, ShowColumns(Filter(colSys, Status = "위험"), SrnKey));

// 2) 최신 데이터
Refresh(EasDashboard);
Refresh(EasTeamsAlarm);
Set(gNow, Now());

// 3) EASDB2: SRN 정규화(따옴표·공백 제거, 대문자) + Teams 발송 여부
ClearCollect(
    colAlert,
    AddColumns(
        EasTeamsAlarm,
        AKey, Upper(Substitute(Substitute(Substitute(Trim(SystemReferenceNumber), """", ""), " ", ""), Char(10), "")),
        ASent, Lower(Trim("" & AlertSent)) in ["true", "yes", "y", "1"]
    )
);

// 4) EasDashboard: 컬럼 정리 + 에러 판정 + EASDB2 연결
//    (컬럼 형식이 다르면 12장을 보고 SiteTxt ~ StartT 줄만 바꾸면 됩니다)
ClearCollect(
    colStage,
    AddColumns(
        AddColumns(
            EasDashboard,
            SrnKey, Upper(Substitute(Substitute(Substitute(Trim(SRN), """", ""), " ", ""), Char(10), "")),
            SrnText, Trim(Substitute(SRN, """", "")),
            SiteTxt, SiteName,
            SysTxt, SystemName,
            FseTxt, FSE,
            RegionTxt, REGION,
            E7, !(Lower(Trim("" & Err7_SupplyWaterHigh)) in ["", "no", "false", "0", "-", "ok"]),
            E9, !(Lower(Trim("" & Err9_ReturnWaterHigh)) in ["", "no", "false", "0", "-", "ok"]),
            E11, !(Lower(Trim("" & Err11_DeltaTempHigh)) in ["", "no", "false", "0", "-", "ok"]),
            E13, !(Lower(Trim("" & Err13_ColdheadGasTempHigh)) in ["", "no", "false", "0", "-", "ok"]),
            E20, !(Lower(Trim("" & Err20_Reminder)) in ["", "no", "false", "0", "-", "ok"]),
            ActiveCnt, Value("" & ActiveErrorCount),
            RemCnt, Value("" & ReminderCount),
            AllOkB, Lower(Trim("" & AllOk)) in ["true", "yes", "y", "1", "예"],
            AllOkT, AllOkTime,
            ErrT, Coalesce(LastUpdatedErrorTime, LastUpdated),
            StartT, StartMonitoring
        ),
        HasCrit, E7 || E9 || E11 || E13,
        ActiveN, If(E7, 1, 0) + If(E9, 1, 0) + If(E11, 1, 0) + If(E13, 1, 0) + If(E20, 1, 0),
        Groups, If(E7 || E9 || E11, "냉각수", "") & If((E7 || E9 || E11) && E13, ", ", "") & If(E13, "Coldhead", ""),
        Sent, LookUp(colAlert, AKey = SrnKey, ASent),
        HasAlert, !IsBlank(LookUp(colAlert, AKey = SrnKey))
    )
);

// 5) 상태 판정 → 화면에서 쓰는 최종 목록
ClearCollect(
    colSys,
    ForAll(
        colStage As s,
        With(
            {
                st: If(
                    s.ActiveN = 0, "정상",
                    s.AllOkB && !IsBlank(s.AllOkT) && (IsBlank(s.ErrT) || s.AllOkT >= s.ErrT), "정상",
                    s.AllOkB && IsBlank(s.AllOkT) && IsBlank(s.ErrT), "정상",
                    s.HasCrit, "위험",
                    "Reminder"
                )
            },
            {
                ID: s.ID,
                SrnKey: s.SrnKey,
                SrnText: s.SrnText,
                SiteTxt: s.SiteTxt,
                SysTxt: s.SysTxt,
                FseTxt: s.FseTxt,
                RegionTxt: s.RegionTxt,
                E7: s.E7, E9: s.E9, E11: s.E11, E13: s.E13, E20: s.E20,
                ActiveCnt: Coalesce(s.ActiveCnt, s.ActiveN),
                RemCnt: s.RemCnt,
                AllOkB: s.AllOkB,
                AllOkT: s.AllOkT,
                ErrT: s.ErrT,
                Since: Coalesce(s.StartT, s.ErrT),
                Groups: s.Groups,
                Sent: Coalesce(s.Sent, false),
                HasAlert: s.HasAlert,
                Status: st,
                Stale: st <> "정상" && !IsBlank(s.ErrT) && DateDiff(s.ErrT, gNow, TimeUnit.Hours) >= gStaleHours,
                Dup: s.SrnKey <> "" && CountIf(colStage, SrnKey = s.SrnKey) > 1,
                SortKey: Switch(st, "위험", 0, "Reminder", 1, 3) * 100000000 - If(IsBlank(s.ErrT), 0, DateDiff(DateTime(2000, 1, 1, 0, 0, 0), s.ErrT, TimeUnit.Minutes))
            }
        )
    )
);

// 6) EASDB2에만 있는 SRN
ClearCollect(colUnmatched, Filter(colAlert, !(AKey in ShowColumns(colStage, SrnKey))));

// 7) 알림: 새로 위험이 된 시스템 / 행 수 제한
If(
    gLoaded,
    With(
        {nc: Filter(colSys, Status = "위험" && !(SrnKey in colPrevCrit))},
        If(CountRows(nc) > 0, Notify("신규 위험 " & CountRows(nc) & "건: " & Concat(FirstN(nc, 5), SrnText & " " & SiteTxt, ", "), NotificationType.Warning, 15000))
    )
);
If(CountRows(colStage) >= 2000 || CountRows(colAlert) >= 2000, Notify("리스트가 2000건 이상이라 일부만 읽었습니다. 관리자에게 알려 주세요.", NotificationType.Error));
Set(gLoaded, true);
Set(gLoadedAt, Now())
```

> 편집기에서 `EasDashboard`, `EasTeamsAlarm` 아래에 노란 삼각형(위임 경고)이 보입니다. 데이터 행 제한(2000) 안에서는 정상 동작하므로 무시해도 됩니다.

### tmrRefresh (60초마다 자동 새로고침)

| 속성 | 값 |
|---|---|
| Duration | `60000` |
| Repeat | `true` |
| AutoStart | `true` |
| Visible | `false` |
| OnTimerEnd | `Select(btnLoad)` |

### Screen1.OnVisible (앱을 열 때 바로 읽기)

```powerfx
Select(btnLoad)
```

---

## 6. 숫자 타일: `galKpi`

가로 갤러리(빈)입니다.

| 속성 | 값 |
|---|---|
| TemplateSize | `188` |
| TemplatePadding | `6` |
| TemplateFill | `If(ThisItem.K = gFilter && gFilter <> "All", RGBA(225, 236, 250, 1), Color.White)` |
| OnSelect | `Set(gFilter, If(gFilter = ThisItem.K, "All", ThisItem.K))` |

**Items:**

<!-- fx:galKpi.Items -->
```powerfx
Table(
    {K: "All",      L: "전체",            C: CountRows(colSys),                    Col: RGBA(120, 130, 145, 1)},
    {K: "Critical", L: "위험",            C: CountIf(colSys, Status = "위험"),     Col: RGBA(210, 40, 40, 1)},
    {K: "Reminder", L: "Reminder",        C: CountIf(colSys, Status = "Reminder"), Col: RGBA(235, 150, 20, 1)},
    {K: "Stale",    L: "AllOk 누락 의심", C: CountIf(colSys, Stale),               Col: RGBA(130, 70, 180, 1)},
    {K: "Ok",       L: "정상",            C: CountIf(colSys, Status = "정상"),     Col: RGBA(40, 160, 80, 1)},
    {K: "Sent",     L: "Teams 발송",      C: CountIf(colSys, Sent),                Col: RGBA(50, 110, 200, 1)},
    {K: "Unknown",  L: "미매칭",          C: CountRows(colUnmatched),              Col: RGBA(150, 150, 165, 1)}
)
```

갤러리 **첫 번째 칸 안에** 아래 컨트롤 3개를 넣습니다.

| 이름 | 종류 | X | Y | W | H | 속성 |
|---|---|---|---|---|---|---|
| `rectKpiBar` | 사각형 | 0 | 0 | 6 | `Parent.TemplateHeight` | Fill `ThisItem.Col` |
| `lblKpiName` | 레이블 | 16 | 4 | 160 | 26 | Text `ThisItem.L`, Color `RGBA(110, 110, 110, 1)` |
| `lblKpiCount` | 레이블 | 16 | 30 | 160 | 44 | Text `ThisItem.C`, Size `24`, FontWeight `FontWeight.Bold` |

타일을 클릭하면 목록이 그 조건으로 걸러지고, 같은 타일을 다시 클릭하면 해제됩니다.

---

## 7. 필터

| 컨트롤 | 속성 | 값 |
|---|---|---|
| `cbRegion` | Items | `Distinct(Filter(colSys, !IsBlank(RegionTxt)), RegionTxt)` |
| | SelectMultiple | `false` |
| | InputTextPlaceholder | `"Region (전체)"` |
| `cbFse` | Items | `Distinct(Filter(colSys, !IsBlank(FseTxt)), FseTxt)` |
| | SelectMultiple | `false` |
| | InputTextPlaceholder | `"FSE (전체)"` |
| `btnClear` | OnSelect | `Reset(txtSearch); Reset(cbRegion); Reset(cbFse); Reset(chkErrors); Set(gFilter, "All")` |

콤보 상자의 DisplayFields·SearchFields는 자동으로 `["Value"]`가 됩니다.

---

## 8. 헤더: `lblHeader.Text`

<!-- fx:lblHeader.Text -->
```powerfx
"갱신 " & Text(gLoadedAt, "hh:mm:ss") & "    |    " &
If(Hour(gNow) >= 22 || Hour(gNow) < 8, "야간", "주간") & "    |    18:00 초기화까지 " &
With(
    {reset: DateAdd(DateTime(Year(gNow), Month(gNow), Day(gNow), 0, 0, 0), 18, TimeUnit.Hours)},
    With(
        {m: DateDiff(gNow, If(gNow >= reset, DateAdd(reset, 1, TimeUnit.Days), reset), TimeUnit.Minutes)},
        RoundDown(m / 60, 0) & "시간 " & Mod(m, 60) & "분"
    )
)
```

---

## 9. 시스템 목록: `galSys`

세로 갤러리(빈)입니다.

| 속성 | 값 |
|---|---|
| TemplateSize | `64` |
| Visible | `gFilter <> "Unknown"` |
| OnSelect | `Set(gSelId, ThisItem.ID)` |
| TemplateFill | (아래) |

**Items:**

<!-- fx:galSys.Items -->
```powerfx
Sort(
    Filter(
        colSys,
        Switch(
            gFilter,
            "Critical", Status = "위험",
            "Reminder", Status = "Reminder",
            "Stale", Stale,
            "Ok", Status = "정상",
            "Sent", Sent,
            true
        ) &&
        (!chkErrors.Value || Status <> "정상") &&
        (IsEmpty(cbRegion.SelectedItems) || RegionTxt = cbRegion.Selected.Value) &&
        (IsEmpty(cbFse.SelectedItems) || FseTxt = cbFse.Selected.Value) &&
        (IsBlank(txtSearch.Text) || txtSearch.Text in SrnText || txtSearch.Text in SiteTxt || txtSearch.Text in SysTxt || txtSearch.Text in FseTxt)
    ),
    SortKey
)
```

정렬 순서는 위험 → Reminder → 정상이고, 같은 상태 안에서는 최근 에러가 위로 옵니다.

**TemplateFill:**

<!-- fx:galSys.TemplateFill -->
```powerfx
If(
    ThisItem.ID = gSelId, RGBA(220, 232, 250, 1),
    Switch(ThisItem.Status, "위험", RGBA(255, 238, 238, 1), "Reminder", RGBA(255, 247, 228, 1), Color.White)
)
```

갤러리 **첫 번째 칸 안에** 아래 컨트롤을 넣습니다.

| 이름 | 종류 | X | Y | W | H | Text |
|---|---|---|---|---|---|---|
| `rectBar` | 사각형 | 0 | 0 | 6 | `Parent.TemplateHeight` | – |
| `lblStatus` | 레이블 | 14 | 6 | 130 | 24 | (아래) |
| `lblSrn` | 레이블 | 14 | 32 | 130 | 24 | `ThisItem.SrnText & If(ThisItem.Dup, " [중복]", "")` |
| `lblSite` | 레이블 | 150 | 6 | 300 | 24 | `ThisItem.SiteTxt` (FontWeight Bold) |
| `lblSys` | 레이블 | 150 | 32 | 300 | 24 | `ThisItem.SysTxt & "  ·  " & ThisItem.FseTxt & "  ·  " & ThisItem.RegionTxt` |
| `lblE7` | 레이블 | 460 | 18 | 44 | 28 | `"7"` |
| `lblE9` | 레이블 | 508 | 18 | 44 | 28 | `"9"` |
| `lblE11` | 레이블 | 556 | 18 | 44 | 28 | `"11"` |
| `lblE13` | 레이블 | 604 | 18 | 44 | 28 | `"13"` |
| `lblE20` | 레이블 | 652 | 18 | 44 | 28 | `"20"` |
| `lblLast` | 레이블 | 708 | 6 | 165 | 24 | (아래) |
| `lblAge` | 레이블 | 708 | 32 | 165 | 24 | (아래) |

- `rectBar.Fill`:

<!-- fx:rectBar.Fill -->
```powerfx
Switch(ThisItem.Status, "위험", RGBA(210, 40, 40, 1), "Reminder", RGBA(235, 150, 20, 1), RGBA(40, 160, 80, 1))
```

- `lblStatus.Text`:

<!-- fx:lblStatus.Text -->
```powerfx
ThisItem.Status & If(ThisItem.Stale, " · 갱신없음", "")
```

- `lblStatus.Color`:

<!-- fx:lblStatus.Color -->
```powerfx
If(ThisItem.Stale, RGBA(130, 70, 180, 1), RGBA(30, 30, 30, 1))
```

  FontWeight는 `FontWeight.Bold`입니다.

- 에러 칸 5개 공통: Align `Align.Center`.
  - `lblE7`의 **Fill**과 **Color**를 아래처럼 넣고, 나머지는 `ThisItem.E7` 부분만 각각 `E9` / `E11` / `E13` / `E20`으로 바꿉니다.
  - `lblE20`은 Reminder라서 빨강 대신 주황 `RGBA(235, 150, 20, 1)`을 씁니다.

<!-- fx:lblE7.Fill -->
```powerfx
If(ThisItem.E7, RGBA(190, 0, 0, 1), RGBA(235, 235, 235, 1))
```

<!-- fx:lblE7.Color -->
```powerfx
If(ThisItem.E7, Color.White, RGBA(160, 160, 160, 1))
```

- `lblLast.Text`:

<!-- fx:lblLast.Text -->
```powerfx
If(IsBlank(ThisItem.ErrT), "", "최근 " & Text(ThisItem.ErrT, "mm-dd hh:mm"))
```

- `lblAge.Text`: 에러가 시작된 뒤 지난 시간 + Teams 발송 표시

<!-- fx:lblAge.Text -->
```powerfx
If(
    ThisItem.Status = "정상" || IsBlank(ThisItem.Since), "",
    With(
        {m: DateDiff(ThisItem.Since, gNow, TimeUnit.Minutes)},
        If(
            m < 60, m & "분",
            m < 1440, RoundDown(m / 60, 0) & "시간 " & Mod(m, 60) & "분",
            RoundDown(m / 1440, 0) & "일 " & RoundDown(Mod(m, 1440) / 60, 0) & "시간"
        ) & " 경과"
    )
) & If(ThisItem.Sent, "  · Teams", "")
```

---

## 10. 상세 패널

`htmlDetail.HtmlText`:

<!-- fx:htmlDetail.HtmlText -->
```powerfx
With(
    {s: LookUp(colSys, ID = gSelId)},
    If(
        IsBlank(s.ID),
        "<div style='color:#888;font-size:14px'>목록에서 시스템을 선택하세요.</div>",
        "<div style='font-size:14px;line-height:1.7'>" &
        "<div style='font-size:20px;font-weight:bold'>" & s.SrnText & "</div>" &
        "<b>상태</b>: " & s.Status & If(s.Stale, " (AllOk 누락 의심: " & gStaleHours & "시간 이상 갱신 없음)", "") & If(s.Dup, " [같은 SRN 행 중복]", "") & "<br>" &
        "<b>사이트</b>: " & s.SiteTxt & "<br>" &
        "<b>시스템</b>: " & s.SysTxt & "<br>" &
        "<b>FSE / Region</b>: " & s.FseTxt & " / " & s.RegionTxt & "<br>" &
        "<b>활성 에러</b>: " & Coalesce(Concat(Filter(Table({n: "7 공급수", on: s.E7}, {n: "9 회수수", on: s.E9}, {n: "11 ΔT", on: s.E11}, {n: "13 Coldhead", on: s.E13}, {n: "20 Reminder", on: s.E20}), on), n, ", "), "-") & "<br>" &
        "<b>에러 시작</b>: " & If(IsBlank(s.Since), "-", Text(s.Since, "yyyy-mm-dd hh:mm")) & "<br>" &
        "<b>최근 에러</b>: " & If(IsBlank(s.ErrT), "-", Text(s.ErrT, "yyyy-mm-dd hh:mm")) & "<br>" &
        "<b>AllOk</b>: " & If(s.AllOkB, "예", "아니요") & " (" & If(IsBlank(s.AllOkT), "-", Text(s.AllOkT, "yyyy-mm-dd hh:mm")) & ")<br>" &
        "<b>활성 에러 수 / Reminder 수</b>: " & s.ActiveCnt & " / " & If(IsBlank(s.RemCnt), "-", Text(s.RemCnt)) & "<br>" &
        "<b>Teams 알림 (EASDB2)</b>: " & If(!s.HasAlert, "EASDB2 항목 없음", s.Sent, "발송됨", "미발송") &
        "</div>"
    )
)
```

`btnOpen`:

| 속성 | 값 |
|---|---|
| Visible | `!IsBlank(gSelId)` |
| OnSelect | `Launch("https://share.philips.com/sites/EASalarmhandling/Lists/EasDashboard/DispForm.aspx?ID=" & gSelId)` |

---

## 11. 미매칭 목록: `galUnmatched`

EASDB2에만 있고 EasDashboard에는 없는 SRN을 보여주는 목록입니다. 세로 갤러리(빈)입니다.

| 속성 | 값 |
|---|---|
| Items | `colUnmatched` |
| Visible | `gFilter = "Unknown"` |
| TemplateSize | `44` |

칸 안에 레이블 하나를 넣습니다: X `14`, Y `8`, W `840`, H `28`

<!-- fx:lblUnmatched.Text -->
```powerfx
ThisItem.SystemReferenceNumber & If(ThisItem.ASent, "    · Teams 발송", "") & "    (EASDB2 ID " & ThisItem.ID & ")"
```

---

## 12. 컬럼 형식이 다를 때

5장 `btnLoad.OnSelect`의 4)번 블록에서 **해당 줄만** 바꿉니다.

| 컬럼 형식 | 원래 줄 | 바꿀 줄 |
|---|---|---|
| REGION이 **선택(Choice)** | `RegionTxt, REGION,` | `RegionTxt, REGION.Value,` |
| FSE가 **사용자(Person)** | `FseTxt, FSE,` | `FseTxt, FSE.DisplayName,` |
| FSE가 **선택** | `FseTxt, FSE,` | `FseTxt, FSE.Value,` |
| SiteName / SystemName이 **선택** | `SiteTxt, SiteName,` | `SiteTxt, SiteName.Value,` (SystemName도 같은 방식) |
| AllOkTime이 **텍스트** | `AllOkT, AllOkTime,` | `AllOkT, DateTimeValue(AllOkTime),` |
| LastUpdated / LastUpdatedErrorTime이 **텍스트** | `ErrT, Coalesce(LastUpdatedErrorTime, LastUpdated),` | `ErrT, DateTimeValue(Coalesce(LastUpdatedErrorTime, LastUpdated)),` |
| StartMonitoring이 **텍스트** | `StartT, StartMonitoring` | `StartT, DateTimeValue(StartMonitoring)` |

수식에서 컬럼 이름에 빨간 밑줄이 생기면 표시 이름이 다른 경우입니다.
- 해당 위치에 `ThisRecord.`를 입력하면 자동 완성 목록이 나오니, 거기서 실제 이름을 고르세요.
- 이름에 공백이 있으면 `'Site Name'`처럼 작은따옴표로 감쌉니다.

---

## 13. 저장 · 게시 · 공유

1. **저장**(Ctrl+S) → **게시**
2. **공유:** make.powerapps.com → 앱 → **공유** → 함께 볼 사람이나 그룹을 입력합니다(공동 소유자는 체크하지 않음).
   - 받는 사람은 EASalarmhandling 사이트에 **읽기 권한**만 있으면 됩니다. 팀 구성원이면 이미 있습니다.
   - 처음 열 때 "SharePoint 연결 허용" 창이 한 번 나오면 **허용**을 누릅니다.
3. **Teams에서 열기 (권장):** EAS alarm handling 팀의 채널 → 탭 추가(**+**) → **Power Apps** → `E-Dashboard` 선택
   - 채널 탭에서 바로 볼 수 있습니다.

---

## 14. 문제 해결

| 증상 | 원인 / 조치 |
|---|---|
| 수식에 빨간 밑줄 | 12장 (컬럼 형식이나 표시 이름이 다름) |
| `colSys`, `gFilter` 등에 빨간 밑줄 | 수식을 넣는 순서 때문일 수 있습니다. App.OnStart와 btnLoad.OnSelect를 먼저 넣고, 앱을 한 번 실행(F5)하면 사라집니다 |
| 목록이 비어 있음 | 미리 보기(F5)에서 `새로고침`을 한 번 누르세요. 편집 화면에서는 Screen.OnVisible이 자동으로 실행되지 않습니다 |
| 시스템 수가 500 또는 2000에서 멈춤 | 2장 3번의 데이터 행 제한 확인 |
| 시각이 9시간 어긋남 | 그 날짜 컬럼이 텍스트 형식입니다 → 12장 `DateTimeValue` 줄로 바꾸기 |
| 모든 시스템이 "정상" | 에러 컬럼 이름(`Err7_SupplyWaterHigh` 등)을 확인하세요. 4)번 블록의 E7~E20 줄에서 쓰는 이름입니다 |
| 다른 사람이 "권한 없음" | 그 사람의 SharePoint 사이트 권한 확인 (앱 공유와 별개) |
