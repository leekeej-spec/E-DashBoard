# E-Dashboard 모바일 앱 (휴대폰 세로)

Power Apps는 한 앱 안의 모든 화면이 같은 크기를 씁니다. 그래서 태블릿(1366 × 768)용 E-Dashboard에 화면을 추가해도 휴대폰에서는 작게 보입니다. 휴대폰용은 **휴대폰 형식의 앱을 따로** 만듭니다. 데이터(리스트)와 수식은 같습니다.

화면은 3개입니다.

| 화면 | 내용 |
|---|---|
| `scrMain` | 헤더, 숫자 타일(옆으로 밀기), 검색, 장비 목록 |
| `scrDetail` | 장비 상세, All OK, SharePoint 열기, 엔지니어 알림 |
| `scrNotify` | 엔지니어 선택, 메시지, Teams·메일 보내기 |

---

## 1. 앱 만들기

1. make.powerapps.com → **+ 만들기 → 빈 앱 → 빈 캔버스 앱**
   - 이름: `E-Dashboard Mobile`
   - 형식: **휴대폰**
2. **⚙ 설정 → 표시**
   - 앱 레이아웃: 휴대폰
   - 방향: 세로
   - 가로 세로 비율 고정: **켬**
3. **⚙ 설정 → 일반 → 데이터 행 제한**: **2000**
4. **⚙ 설정 → 업데이트 → 최신 컨트롤 및 테마**: **끔** (클래식 컨트롤 사용)
5. **데이터 추가:**
   - SharePoint `EASalarmhandling` 사이트의 **EasDashboard**, **EasTeamsAlarm**, **FseEmail**
   - **Office 365 Outlook** (메일 보내기를 쓸 경우)
6. **화면 3개 만들기**
   - 기본 화면 `Screen1`의 이름을 **`scrMain`** 으로 바꿉니다.
   - **새 화면 → 빈 화면**을 두 번 만들어 이름을 **`scrDetail`**, **`scrNotify`** 로 바꿉니다.
   - 붙여넣을 YAML이 다른 화면 이름을 참조하므로, **YAML보다 화면을 먼저** 만들어 두세요.

## 2. 앱·화면 속성 (직접 입력)

**App.OnStart**
```
Set(gStaleHours, 6);
Set(gFilter, "All");
Set(gLoaded, false);
Set(gNeedReload, false);
Set(gConfirm, false);
Set(gNow, Now());
Set(gLoadedAt, Now());
Set(gSelId, Blank())
```

**scrMain.OnVisible**: 처음 열 때나 All OK 처리 후 목록으로 돌아올 때 다시 읽습니다.
```
If(!gLoaded || gNeedReload, Set(gNeedReload, false); Select(btnLoad))
```

**scrNotify.OnVisible**
```
Reset(ddEngineer); Reset(txtMsg)
```

**자동 새로고침 (선택):** `scrMain`에 **+ 삽입 → 타이머**를 넣고 아래 속성을 줍니다.
- Duration `60000`
- Repeat `true`
- AutoStart `true`
- Visible `false`
- OnTimerEnd `Select(btnLoad)`

---

## 3. scrMain YAML

트리 보기에서 **scrMain 우클릭 → 붙여넣기**:

```yaml
- rectHeader:
    Control: Rectangle@2.3.0
    Properties:
      Fill: =RGBA(30, 45, 70, 1)
      Height: =96
      Width: =640
      X: =0
      Y: =0
- lblTitle:
    Control: Label@2.5.1
    Properties:
      Color: =Color.White
      FontWeight: =FontWeight.Bold
      Size: =22
      Text: ="E-Dashboard"
      Height: =44
      Width: =430
      X: =20
      Y: =8
- lblHeader:
    Control: Label@2.5.1
    Properties:
      Color: =RGBA(220, 225, 235, 1)
      Size: =12
      Height: =36
      Width: =440
      X: =20
      Y: =52
      Text: |-
        ="갱신 " & Text(gLoadedAt, "hh:mm") & "  |  " &
        If(Hour(gNow) >= 22 || Hour(gNow) < 8, "야간", "주간") & "  |  18:00까지 " &
        With(
            {reset: DateAdd(DateTime(Year(gNow), Month(gNow), Day(gNow), 0, 0, 0), 18, TimeUnit.Hours)},
            With(
                {m: DateDiff(gNow, If(gNow >= reset, DateAdd(reset, 1, TimeUnit.Days), reset), TimeUnit.Minutes)},
                RoundDown(m / 60, 0) & "시간 " & Mod(m, 60) & "분"
            )
        )
- btnLoad:
    Control: Classic/Button@2.2.0
    Properties:
      Fill: =RGBA(0, 120, 212, 1)
      Color: =Color.White
      Size: =15
      Text: ="새로고침"
      Height: =56
      Width: =150
      X: =470
      Y: =20
      OnSelect: |-
        =ClearCollect(colPrevCrit, ShowColumns(Filter(colSys, Status = "위험"), SrnKey));
        Refresh(EasDashboard);
        Refresh(EasTeamsAlarm);
        Set(gNow, Now());
        ClearCollect(
            colAlert,
            AddColumns(
                EasTeamsAlarm,
                AKey, Upper(Substitute(Substitute(Substitute(Trim(SystemReferenceNumber), """", ""), " ", ""), Char(10), "")),
                ASent, Lower(Trim("" & AlertSent)) in ["true", "yes", "y", "1"]
            )
        );
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
                    AllOkT, DateTimeValue(AllOkTime, "en-US"),
                    ErrT, DateTimeValue(Coalesce(LastUpdatedErrorTime, LastUpdated), "en-US"),
                    StartT, DateTimeValue(StartMonitoring, "en-US")
                ),
                HasCrit, E7 || E9 || E11 || E13,
                ActiveN, If(E7, 1, 0) + If(E9, 1, 0) + If(E11, 1, 0) + If(E13, 1, 0) + If(E20, 1, 0),
                Groups, If(E7 || E9 || E11, "냉각수", "") & If((E7 || E9 || E11) && E13, ", ", "") & If(E13, "Coldhead", ""),
                Sent, LookUp(colAlert, AKey = SrnKey, ASent),
                HasAlert, !IsBlank(LookUp(colAlert, AKey = SrnKey))
            )
        );
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
        If(
            gLoaded,
            With(
                {nc: Filter(colSys, Status = "위험" && !(SrnKey in colPrevCrit))},
                If(CountRows(nc) > 0, Notify("신규 위험 " & CountRows(nc) & "건: " & Concat(FirstN(nc, 5), SrnText & " " & SiteTxt, ", "), NotificationType.Warning, 15000))
            )
        );
        Set(gLoaded, true);
        Set(gLoadedAt, Now())
- galKpi:
    Control: Gallery@2.15.0
    Variant: BrowseLayout_Horizontal_TwoTextOneImageVariant_ver5.0
    Properties:
      BorderColor: =RGBA(245, 245, 245, 1)
      Height: =100
      Width: =640
      X: =0
      Y: =100
      TemplatePadding: =6
      TemplateSize: =150
      OnSelect: =Set(gFilter, If(gFilter = ThisItem.K, "All", ThisItem.K))
      TemplateFill: =If(ThisItem.K = gFilter && gFilter <> "All", RGBA(225, 236, 250, 1), Color.White)
      Items: |-
        =Table(
            {K: "All",      L: "전체",       C: CountRows(colSys),                    Col: RGBA(120, 130, 145, 1)},
            {K: "Critical", L: "위험",       C: CountIf(colSys, Status = "위험"),     Col: RGBA(210, 40, 40, 1)},
            {K: "Reminder", L: "Reminder",   C: CountIf(colSys, Status = "Reminder"), Col: RGBA(235, 150, 20, 1)},
            {K: "Stale",    L: "AllOk 누락", C: CountIf(colSys, Stale),               Col: RGBA(130, 70, 180, 1)},
            {K: "Sent",     L: "Teams 발송", C: CountIf(colSys, Sent),                Col: RGBA(50, 110, 200, 1)},
            {K: "Ok",       L: "정상",       C: CountIf(colSys, Status = "정상"),     Col: RGBA(40, 160, 80, 1)}
        )
    Children:
      - rectKpiBar:
          Control: Rectangle@2.3.0
          Properties:
            Fill: =ThisItem.Col
            Height: =Parent.TemplateHeight
            OnSelect: =Select(Parent)
            Width: =6
            X: =0
            Y: =0
      - lblKpiName:
          Control: Label@2.5.1
          Properties:
            Color: =RGBA(110, 110, 110, 1)
            Size: =12
            Height: =30
            OnSelect: =Select(Parent)
            Text: =ThisItem.L
            Width: =130
            X: =14
            Y: =4
      - lblKpiCount:
          Control: Label@2.5.1
          Properties:
            FontWeight: =FontWeight.Bold
            Size: =24
            Height: =48
            OnSelect: =Select(Parent)
            Text: =ThisItem.C
            Width: =130
            X: =14
            Y: =34
- txtSearch:
    Control: Classic/TextInput@2.3.2
    Properties:
      Default: =""
      HintText: ="SRN · 사이트 · 시스템 · FSE 검색"
      Size: =14
      Height: =52
      Width: =600
      X: =20
      Y: =208
- galSys:
    Control: Gallery@2.15.0
    Variant: BrowseLayout_Vertical_TwoTextOneImageVariant_ver5.0
    Properties:
      BorderColor: =RGBA(230, 230, 230, 1)
      Height: =868
      Width: =640
      X: =0
      Y: =268
      TemplatePadding: =0
      TemplateSize: =116
      OnSelect: =Set(gSelId, ThisItem.ID); Set(gConfirm, false); Navigate(scrDetail, ScreenTransition.Cover)
      TemplateFill: |-
        =Switch(ThisItem.Status, "위험", RGBA(255, 238, 238, 1), "Reminder", RGBA(255, 247, 228, 1), Color.White)
      Items: |-
        =Sort(
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
                (IsBlank(Trim(txtSearch.Text)) || Trim(txtSearch.Text) in SrnText || Trim(txtSearch.Text) in SiteTxt || Trim(txtSearch.Text) in SysTxt || Trim(txtSearch.Text) in FseTxt)
            ),
            SortKey
        )
    Children:
      - rectBar:
          Control: Rectangle@2.3.0
          Properties:
            Fill: =Switch(ThisItem.Status, "위험", RGBA(210, 40, 40, 1), "Reminder", RGBA(235, 150, 20, 1), RGBA(40, 160, 80, 1))
            Height: =Parent.TemplateHeight
            OnSelect: =Select(Parent)
            Width: =8
            X: =0
            Y: =0
      - lblStatus:
          Control: Label@2.5.1
          Properties:
            Color: =If(ThisItem.Stale, RGBA(130, 70, 180, 1), RGBA(30, 30, 30, 1))
            FontWeight: =FontWeight.Bold
            Size: =14
            Height: =32
            OnSelect: =Select(Parent)
            Text: =ThisItem.Status & If(ThisItem.Stale, " · 갱신없음", "")
            Width: =210
            X: =18
            Y: =6
      - lblSrn:
          Control: Label@2.5.1
          Properties:
            Size: =14
            Height: =32
            OnSelect: =Select(Parent)
            Text: =ThisItem.SrnText & If(ThisItem.Dup, " [중복]", "")
            Width: =180
            X: =230
            Y: =6
      - lblLast:
          Control: Label@2.5.1
          Properties:
            Align: =Align.Right
            Color: =RGBA(90, 90, 90, 1)
            Size: =12
            Height: =32
            OnSelect: =Select(Parent)
            Text: =If(IsBlank(ThisItem.ErrT), "", Text(ThisItem.ErrT, "mm-dd hh:mm"))
            Width: =200
            X: =420
            Y: =6
      - lblSite:
          Control: Label@2.5.1
          Properties:
            FontWeight: =FontWeight.Bold
            Size: =14
            Height: =32
            OnSelect: =Select(Parent)
            Text: =ThisItem.SiteTxt
            Width: =600
            X: =18
            Y: =40
      - lblSys:
          Control: Label@2.5.1
          Properties:
            Color: =RGBA(90, 90, 90, 1)
            Size: =12
            Height: =32
            OnSelect: =Select(Parent)
            Text: =ThisItem.SysTxt & "  ·  " & ThisItem.FseTxt
            Width: =375
            X: =18
            Y: =76
      - lblE7:
          Control: Label@2.5.1
          Properties:
            Align: =Align.Center
            Color: =If(ThisItem.E7, Color.White, RGBA(160, 160, 160, 1))
            Fill: =If(ThisItem.E7, RGBA(190, 0, 0, 1), RGBA(235, 235, 235, 1))
            FontWeight: =FontWeight.Bold
            Size: =11
            Height: =30
            OnSelect: =Select(Parent)
            Text: ="7"
            Width: =40
            X: =400
            Y: =78
      - lblE9:
          Control: Label@2.5.1
          Properties:
            Align: =Align.Center
            Color: =If(ThisItem.E9, Color.White, RGBA(160, 160, 160, 1))
            Fill: =If(ThisItem.E9, RGBA(190, 0, 0, 1), RGBA(235, 235, 235, 1))
            FontWeight: =FontWeight.Bold
            Size: =11
            Height: =30
            OnSelect: =Select(Parent)
            Text: ="9"
            Width: =40
            X: =444
            Y: =78
      - lblE11:
          Control: Label@2.5.1
          Properties:
            Align: =Align.Center
            Color: =If(ThisItem.E11, Color.White, RGBA(160, 160, 160, 1))
            Fill: =If(ThisItem.E11, RGBA(190, 0, 0, 1), RGBA(235, 235, 235, 1))
            FontWeight: =FontWeight.Bold
            Size: =11
            Height: =30
            OnSelect: =Select(Parent)
            Text: ="11"
            Width: =40
            X: =488
            Y: =78
      - lblE13:
          Control: Label@2.5.1
          Properties:
            Align: =Align.Center
            Color: =If(ThisItem.E13, Color.White, RGBA(160, 160, 160, 1))
            Fill: =If(ThisItem.E13, RGBA(190, 0, 0, 1), RGBA(235, 235, 235, 1))
            FontWeight: =FontWeight.Bold
            Size: =11
            Height: =30
            OnSelect: =Select(Parent)
            Text: ="13"
            Width: =40
            X: =532
            Y: =78
      - lblE20:
          Control: Label@2.5.1
          Properties:
            Align: =Align.Center
            Color: =If(ThisItem.E20, Color.White, RGBA(160, 160, 160, 1))
            Fill: =If(ThisItem.E20, RGBA(235, 150, 20, 1), RGBA(235, 235, 235, 1))
            FontWeight: =FontWeight.Bold
            Size: =11
            Height: =30
            OnSelect: =Select(Parent)
            Text: ="20"
            Width: =40
            X: =576
            Y: =78
```

---

## 4. scrDetail YAML

트리 보기에서 **scrDetail 우클릭 → 붙여넣기**:

```yaml
- rectHeaderD:
    Control: Rectangle@2.3.0
    Properties:
      Fill: =RGBA(30, 45, 70, 1)
      Height: =96
      Width: =640
      X: =0
      Y: =0
- btnBackD:
    Control: Classic/Button@2.2.0
    Properties:
      Fill: =RGBA(30, 45, 70, 1)
      Color: =Color.White
      Size: =16
      Text: ="< 목록"
      OnSelect: =Navigate(scrMain, ScreenTransition.UnCover)
      Height: =56
      Width: =140
      X: =10
      Y: =20
- lblTitleD:
    Control: Label@2.5.1
    Properties:
      Color: =Color.White
      FontWeight: =FontWeight.Bold
      Size: =20
      Text: =LookUp(colSys, ID = gSelId).SrnText & "  " & LookUp(colSys, ID = gSelId).Status
      Height: =56
      Width: =460
      X: =165
      Y: =20
- htmlDetail:
    Control: HtmlViewer@2.1.0
    Properties:
      Height: =560
      Width: =600
      X: =20
      Y: =110
      HtmlText: |-
        =With(
            {s: LookUp(colSys, ID = gSelId)},
            If(
                IsBlank(s.ID),
                "<div style='color:#888;font-size:16px'>목록에서 장비를 선택하세요.</div>",
                "<div style='font-size:16px;line-height:1.8'>" &
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
- btnAllOk:
    Control: Classic/Button@2.2.0
    Properties:
      Fill: =RGBA(40, 160, 80, 1)
      Size: =16
      OnSelect: =Set(gConfirm, true)
      Text: ="All OK 처리 (수동 초기화)"
      Visible: =!IsBlank(gSelId) && !gConfirm && LookUp(colSys, ID = gSelId).Status <> "정상"
      Height: =60
      Width: =600
      X: =20
      Y: =690
- btnConfirm:
    Control: Classic/Button@2.2.0
    Properties:
      Fill: =RGBA(210, 40, 40, 1)
      Size: =16
      Text: |-
        ="확인: " & LookUp(colSys, ID = gSelId).SrnText & " 초기화"
      Visible: =gConfirm
      Height: =60
      Width: =390
      X: =20
      Y: =690
      OnSelect: |-
        =IfError(
            Patch(
                EasDashboard,
                LookUp(EasDashboard, ID = gSelId),
                {
                    Err7_SupplyWaterHigh: false,
                    Err9_ReturnWaterHigh: false,
                    Err11_DeltaTempHigh: false,
                    Err13_ColdheadGasTempHigh: false,
                    Err20_Reminder: false,
                    AllOk: true,
                    AllOkTime: Text(Now(), "mm/dd/yyyy hh:mm", "en-US"),
                    StartMonitoring: "",
                    ActiveErrorCount: 0,
                    ReminderCount: "0"
                }
            );
            With(
                {al: LookUp(colAlert, AKey = LookUp(colSys, ID = gSelId).SrnKey)},
                If(!IsBlank(al), Patch(EasTeamsAlarm, LookUp(EasTeamsAlarm, ID = al.ID), {AlertSent: "false"}))
            );
            Notify("All OK 처리 완료: " & LookUp(colSys, ID = gSelId).SrnText, NotificationType.Success);
            Set(gConfirm, false);
            Set(gNeedReload, true);
            Navigate(scrMain, ScreenTransition.UnCover),
            Notify("All OK 처리 실패: " & FirstError.Message, NotificationType.Error);
            Set(gConfirm, false)
        )
- btnCancel:
    Control: Classic/Button@2.2.0
    Properties:
      Fill: =RGBA(150, 150, 150, 1)
      Size: =16
      OnSelect: =Set(gConfirm, false)
      Text: ="취소"
      Visible: =gConfirm
      Height: =60
      Width: =200
      X: =420
      Y: =690
- btnOpen:
    Control: Classic/Button@2.2.0
    Properties:
      Size: =16
      OnSelect: =Launch("https://share.philips.com/sites/EASalarmhandling/Lists/EasDashboard/DispForm.aspx?ID=" & gSelId)
      Text: ="SharePoint에서 열기"
      Visible: =!IsBlank(gSelId)
      Height: =60
      Width: =600
      X: =20
      Y: =770
- btnNotify:
    Control: Classic/Button@2.2.0
    Properties:
      Fill: =RGBA(98, 100, 167, 1)
      Size: =16
      OnSelect: =Navigate(scrNotify, ScreenTransition.Cover)
      Text: ="엔지니어에게 알림"
      Visible: =!IsBlank(gSelId)
      Height: =60
      Width: =600
      X: =20
      Y: =850
```

---

## 5. scrNotify YAML

트리 보기에서 **scrNotify 우클릭 → 붙여넣기**:

```yaml
- rectHeaderN:
    Control: Rectangle@2.3.0
    Properties:
      Fill: =RGBA(30, 45, 70, 1)
      Height: =96
      Width: =640
      X: =0
      Y: =0
- btnBackN:
    Control: Classic/Button@2.2.0
    Properties:
      Fill: =RGBA(30, 45, 70, 1)
      Color: =Color.White
      Size: =16
      Text: ="< 상세"
      OnSelect: =Navigate(scrDetail, ScreenTransition.UnCover)
      Height: =56
      Width: =140
      X: =10
      Y: =20
- lblTitleN:
    Control: Label@2.5.1
    Properties:
      Color: =Color.White
      FontWeight: =FontWeight.Bold
      Size: =20
      Text: ="엔지니어 알림"
      Height: =56
      Width: =460
      X: =165
      Y: =20
- lblSubN:
    Control: Label@2.5.1
    Properties:
      FontWeight: =FontWeight.Bold
      Size: =15
      Text: =LookUp(colSys, ID = gSelId).SrnText & "  " & LookUp(colSys, ID = gSelId).SiteTxt
      Height: =56
      Width: =600
      X: =20
      Y: =108
- lblEngineer:
    Control: Label@2.5.1
    Properties:
      Size: =13
      Text: |-
        ="받는 사람   " & LookUp(FseEmail, FseName = ddEngineer.Selected.Value).Email
      Height: =32
      Width: =600
      X: =20
      Y: =172
- ddEngineer:
    Control: Classic/DropDown@2.3.1
    Properties:
      Default: =LookUp(colSys, ID = gSelId).FseTxt
      Items: =Sort(Distinct(FseEmail, FseName), Value)
      Size: =15
      Height: =56
      Width: =600
      X: =20
      Y: =208
- lblMsg:
    Control: Label@2.5.1
    Properties:
      Size: =13
      Text: ="메시지 (수정 가능)"
      Height: =32
      Width: =600
      X: =20
      Y: =280
- txtMsg:
    Control: Classic/TextInput@2.3.2
    Properties:
      Mode: =TextMode.MultiLine
      Size: =14
      Height: =400
      Width: =600
      X: =20
      Y: =316
      Default: |-
        =With(
            {s: LookUp(colSys, ID = gSelId)},
            "[EAS 알람] " & s.Status & If(s.Groups <> "", " (" & s.Groups & ")", "") & Char(10) &
            "SRN: " & s.SrnText & Char(10) &
            "사이트: " & s.SiteTxt & Char(10) &
            "시스템: " & s.SysTxt & Char(10) &
            "활성 에러: " & Coalesce(Concat(Filter(Table({n: "7 공급수", on: s.E7}, {n: "9 회수수", on: s.E9}, {n: "11 ΔT", on: s.E11}, {n: "13 Coldhead", on: s.E13}, {n: "20 Reminder", on: s.E20}), on), n, ", "), "-") & Char(10) &
            "에러 시작: " & If(IsBlank(s.Since), "-", Text(s.Since, "yyyy-mm-dd hh:mm")) & Char(10) &
            "확인 부탁드립니다."
        )
- btnSendTeams:
    Control: Classic/Button@2.2.0
    Properties:
      DisplayMode: =If(IsBlank(ddEngineer.Selected.Value), DisplayMode.Disabled, DisplayMode.Edit)
      Fill: =RGBA(98, 100, 167, 1)
      Size: =16
      Text: ="Teams 채팅 열기"
      Height: =60
      Width: =600
      X: =20
      Y: =740
      OnSelect: |-
        =Launch(
            "https://teams.microsoft.com/l/chat/0/0?users=" & Trim(LookUp(FseEmail, FseName = ddEngineer.Selected.Value).Email) &
            "&message=" & EncodeUrl(txtMsg.Text)
        )
- btnSendMail:
    Control: Classic/Button@2.2.0
    Properties:
      DisplayMode: =If(IsBlank(ddEngineer.Selected.Value), DisplayMode.Disabled, DisplayMode.Edit)
      Fill: =RGBA(0, 120, 212, 1)
      Size: =16
      Text: ="메일 보내기"
      Height: =60
      Width: =600
      X: =20
      Y: =820
      OnSelect: |-
        =IfError(
            Office365Outlook.SendEmailV2(
                Trim(LookUp(FseEmail, FseName = ddEngineer.Selected.Value).Email),
                "[EAS 알람] " & LookUp(colSys, ID = gSelId).SrnText & " " & LookUp(colSys, ID = gSelId).SiteTxt & " - " & LookUp(colSys, ID = gSelId).Status,
                Substitute(txtMsg.Text, Char(10), "<br>")
            );
            Notify("메일 보냄: " & ddEngineer.Selected.Value, NotificationType.Success);
            Navigate(scrDetail, ScreenTransition.UnCover),
            Notify("메일 실패: " & FirstError.Message, NotificationType.Error)
        )
```

---

## 6. 확인과 게시

1. **App → OnStart 실행** → scrMain에서 **F5**
2. **목록 화면:** 숫자 타일을 옆으로 밀어 보고, 타일을 눌러 필터가 되는지, 검색이 되는지 확인합니다.
3. **상세 화면:** 장비를 누르면 상세 화면으로 넘어가고, **< 목록**으로 돌아오는지 확인합니다.
4. **All OK:** 처리하면 목록으로 돌아오면서 다시 읽고, 그 장비가 정상으로 바뀌는지 확인합니다.
5. **엔지니어 알림:** 엔지니어가 자동 선택되는지, Teams·메일 보내기가 되는지 확인합니다.
6. **저장 → 게시 → 공유**: 태블릿 앱과 같은 사람·그룹에게 공유합니다.
7. **휴대폰에서 열기:** Power Apps 모바일 앱 또는 Teams 모바일 → **E-Dashboard Mobile**

## 7. 붙여넣기 오류가 날 때

| 증상 | 조치 |
|---|---|
| 버전 오류 (`DropDown`, `TextInput`, `HtmlViewer` 등) | 그 컨트롤을 직접 하나 넣고 **코드 보기**에서 `Control:` 줄을 확인한 뒤, YAML의 버전을 그 값으로 바꿔 다시 붙여넣기. 태블릿 앱에서 성공한 버전이 있으면 그것을 쓰면 됩니다 |
| `scrDetail`, `scrNotify`에 빨간 밑줄 | 그 이름의 화면이 아직 없음 → 1장 6번 |
| `PA1001 ... invalid mapping` | 한 줄 값 안에 `: `가 있음 → `|-` 블록 형식으로 |
| 이름에 `_1`이 붙음 | 같은 이름 컨트롤이 이미 있음 → 기존 것 삭제 후 다시 붙여넣기 |
