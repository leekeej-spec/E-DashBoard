<#
.SYNOPSIS
    E-Dashboard: EAS 알람 SharePoint 리스트(EasDashboard, EASDB2) 모니터 (P1 조회 버전)

.DESCRIPTION
    데이터 원본은 두 가지입니다.
      File (기본): Power Automate가 두 리스트를 OneDrive의 JSON 파일로 내보내고(docs/FLOW-EXPORT.md),
                   OneDrive 동기화로 PC에 내려온 파일을 읽습니다. SharePoint Online은 Windows 계정만으로
                   REST 접근이 안 되기 때문에 이 방식을 씁니다.
      Rest       : SharePoint REST를 Windows 통합인증으로 직접 호출 (온프레미스 SharePoint용).

    두 리스트를 주기적으로 읽어 SRN 단위로 합치고, 상태(위험/Reminder/정상/미매칭)와
    'AllOk 누락 의심'(에러가 StaleHours 이상 갱신되지 않음)을 판정해 보여줍니다.
    SharePoint에는 아무것도 쓰지 않습니다 (읽기 전용).

    이 파일은 UTF-8(BOM)으로 저장되어 있어야 Windows PowerShell 5.1에서 한글이 깨지지 않습니다.
    GitHub에서 "Download raw file" 또는 ZIP으로 받으세요 (메모장 복사·붙여넣기 금지).

.EXAMPLE
    EDashboard.cmd                실제 데이터 (OneDrive\EDashboard\edash-data.json)
    EDashboard.cmd -DataFile D:\x\edash-data.json   다른 위치의 데이터 파일
    EDashboard.cmd -Demo          샘플 데이터로 화면 확인 (SharePoint 연결 없이)
    EDashboard.cmd -CheckOnly     화면 없이 한 번 읽고 요약만 출력
#>
[CmdletBinding()]
param(
    [string] $SiteUrl        = 'https://share.philips.com/sites/EASalarmhandling',
    [string] $DashboardList  = 'EasDashboard',
    [string] $AlertList      = 'EASDB2',
    [int]    $RefreshSeconds = 60,
    [int]    $StaleHours     = 6,
    [int]    $NightStartHour = 22,
    [int]    $NightEndHour   = 8,
    [string] $DailyReset     = '18:00',
    [ValidateSet('File', 'Rest')]
    [string] $Source         = 'File',
    [string] $DataFile,
    [int]    $MaxFileAgeMinutes = 15,
    [switch] $Demo,
    [switch] $CheckOnly
)

$ErrorActionPreference = 'Stop'

if (-not $DataFile) {
    # OneDriveCommercial = 회사 OneDrive(OneDrive - Philips) 동기화 폴더
    $oneDrive = @($env:OneDriveCommercial, $env:OneDrive, $env:USERPROFILE, $HOME) | Where-Object { $_ } | Select-Object -First 1
    $DataFile = Join-Path (Join-Path $oneDrive 'EDashboard') 'edash-data.json'
}

$Config = @{
    SiteUrl        = $SiteUrl.TrimEnd('/')
    DashboardList  = $DashboardList
    AlertList      = $AlertList
    RefreshSeconds = [Math]::Max(15, $RefreshSeconds)
    StaleHours     = $StaleHours
    NightStartHour = $NightStartHour
    NightEndHour   = $NightEndHour
    DailyReset     = $DailyReset
    Source         = $Source
    DataFile       = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DataFile)
    MaxFileAgeMinutes = $MaxFileAgeMinutes
    Demo           = [bool]$Demo

    # EasDashboard 내부명 매핑 (사용자 컬럼 매핑표 기준)
    DashboardFields = @(
        @{ Name = 'SRN';                  Internal = 'SRN' }
        @{ Name = 'SiteName';             Internal = 'field_2' }
        @{ Name = 'SystemName';           Internal = 'field_3' }
        @{ Name = 'FSE';                  Internal = 'field_4' }
        @{ Name = 'Region';               Internal = 'field_5' }
        @{ Name = 'LastUpdated';          Internal = 'field_14' }
        @{ Name = 'AllOkTime';            Internal = 'field_16' }
        @{ Name = 'AllOk';                Internal = 'AllOk' }
        @{ Name = 'StartMonitoring';      Internal = 'StartMonitoring' }
        @{ Name = 'ActiveErrorCount';     Internal = 'ActiveErrorCount' }
        @{ Name = 'ReminderCount';        Internal = 'ReminderCount' }
        @{ Name = 'LastUpdatedErrorTime'; Internal = 'LastUpdatedErrorTime' }
    )

    # 에러 컬럼 (Flow1 ComposeErrorGroup 기준). Critical = 위험으로 판정
    ErrorColumns = @(
        @{ Name = 'Err7';  Internal = 'field_6';  Header = '7 공급수';    Group = '냉각수';   Critical = $true }
        @{ Name = 'Err9';  Internal = 'field_7';  Header = '9 회수수';    Group = '냉각수';   Critical = $true }
        @{ Name = 'Err11'; Internal = 'field_8';  Header = '11 ΔT';       Group = '냉각수';   Critical = $true }
        @{ Name = 'Err13'; Internal = 'field_9';  Header = '13 Coldhead'; Group = 'Coldhead'; Critical = $true }
        @{ Name = 'Err20'; Internal = 'field_10'; Header = '20 Rem';      Group = 'Reminder'; Critical = $false }
    )
}

# ---------------------------------------------------------------------------------------------
# 데이터 계층: 화면과 무관한 순수 로직. 백그라운드 Runspace와 메인 스레드 양쪽에 같은 코드를 로드한다.
# ---------------------------------------------------------------------------------------------
$DataCode = {
    $ErrorActionPreference = 'Stop'
    $ProgressPreference    = 'SilentlyContinue'
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
    try { [Net.WebRequest]::DefaultWebProxy.Credentials = [Net.CredentialCache]::DefaultNetworkCredentials } catch { }

    if ($null -eq $global:EDashFieldCache) { $global:EDashFieldCache = @{} }
    if ($null -eq $global:EDashAccept)     { $global:EDashAccept = 'application/json;odata=nometadata' }
    # 한국은 서머타임이 없으므로 UTC+9 고정 시간대로 충분하다 (PC 시간대 설정과 무관).
    if ($null -eq $global:EDashKst)        { $global:EDashKst = [TimeZoneInfo]::CreateCustomTimeZone('KST', [TimeSpan]::FromHours(9), 'KST', 'KST') }

    function Get-EDashKstNow {
        [TimeZoneInfo]::ConvertTime([DateTimeOffset]::UtcNow, $global:EDashKst).DateTime
    }

    function ConvertTo-EDashKst {
        param($Value)
        if ($null -eq $Value) { return $null }
        try {
            if ($Value -is [datetime]) {
                $dto = [DateTimeOffset]$Value
            } else {
                $s = ([string]$Value).Trim()
                if (-not $s) { return $null }
                $dto = [DateTimeOffset]::Parse($s, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal)
            }
            return [TimeZoneInfo]::ConvertTime($dto, $global:EDashKst).DateTime
        } catch {
            return $null
        }
    }

    function ConvertTo-EDashText {
        param($Value)
        if ($null -eq $Value)      { return '' }
        if ($Value -is [string])   { return $Value }
        if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
        if ($Value -is [array])    { return (@($Value | ForEach-Object { ConvertTo-EDashText $_ }) -join ', ') }
        if ($Value -is [System.Management.Automation.PSCustomObject]) {
            if ($Value.PSObject.Properties['results'])     { return ConvertTo-EDashText @($Value.results) }
            if ($Value.PSObject.Properties['Value'])       { return [string]$Value.Value }        # Power Automate 선택/조회 컬럼
            if ($Value.PSObject.Properties['DisplayName']) { return [string]$Value.DisplayName }  # Power Automate 사용자 컬럼
            if ($Value.PSObject.Properties['Label'])       { return [string]$Value.Label }
            if ($Value.PSObject.Properties['Title'])       { return [string]$Value.Title }
            return '{...}'
        }
        return [string]$Value
    }

    function ConvertTo-EDashBool {
        param($Value)
        if ($null -eq $Value)    { return $false }
        if ($Value -is [bool])   { return $Value }
        if ($Value -is [string]) { return ($Value.Trim() -in @('true', 'yes', 'y', '1', '예')) }
        try { return ([double]$Value -ne 0) } catch { return $false }
    }

    function ConvertTo-EDashInt {
        param($Value)
        if ($null -eq $Value) { return $null }
        try { return [int][double]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture) } catch { return $null }
    }

    # 에러 컬럼 값이 '활성'인지: 값이 있고 No/false/0 같은 '없음' 표시가 아니면 활성.
    function Test-EDashActive {
        param($Value)
        if ($null -eq $Value)      { return $false }
        if ($Value -is [bool])     { return $Value }
        if ($Value -is [datetime]) { return $true }
        $s = (ConvertTo-EDashText $Value).Trim()
        if (-not $s) { return $false }
        return ($s -notin @('no', 'false', '0', 'n', '-', 'ok', 'off', '아니요'))
    }

    function Get-EDashSrnKey {
        param([string]$Srn)
        if (-not $Srn) { return '' }
        return ($Srn -replace '["\s]', '').ToUpperInvariant()
    }

    function Invoke-EDashSpGet {
        param([string]$Url)
        $req = @{ Uri = $Url; Headers = @{ Accept = $global:EDashAccept }; UseDefaultCredentials = $true; UseBasicParsing = $true }
        if ($PSVersionTable.PSVersion.Major -ge 6 -and $Url -like 'http:*') { $req.AllowUnencryptedAuthentication = $true }
        try {
            $resp = Invoke-WebRequest @req
        } catch {
            $err    = $_
            $detail = $null
            if ($err.ErrorDetails -and $err.ErrorDetails.Message) {
                $detail = $err.ErrorDetails.Message
                try {
                    $json = $detail | ConvertFrom-Json
                    if ($json.'odata.error') { $detail = $json.'odata.error'.message.value }
                    elseif ($json.error)     { $detail = $json.error.message.value }
                } catch { }
            }
            # 구버전 SharePoint 2013은 odata=nometadata를 거부한다 -> verbose로 재시도
            if ($global:EDashAccept -like '*nometadata*' -and "$detail" -match 'ACCEPT') {
                $global:EDashAccept = 'application/json;odata=verbose'
                return Invoke-EDashSpGet $Url
            }
            $status = $null
            try { $status = [int]$err.Exception.Response.StatusCode } catch { }
            $msg = 'SharePoint 요청 실패'
            if ($status) { $msg += ' (HTTP ' + $status + ')' }
            if ($detail) { $msg += ': ' + ($detail -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim() }
            else         { $msg += ': ' + $err.Exception.Message }
            if ($msg.Length -gt 400) { $msg = $msg.Substring(0, 400) + '...' }
            throw $msg
        }
        $ctype = [string](@($resp.Headers['Content-Type']) -join ',')
        if ($ctype -and $ctype -notmatch 'json') { throw 'SharePoint가 JSON이 아닌 응답을 보냈습니다 (로그인 페이지로 이동했을 수 있음).' }
        $text = [Text.Encoding]::UTF8.GetString($resp.RawContentStream.ToArray())
        # 리스트 항목에는 "Id"와 "ID"가 함께 오는데, PowerShell JSON 파서는 대소문자만 다른 키를 거부한다.
        $text = $text -creplace '"Id":', '"Id_":'
        $obj = $text | ConvertFrom-Json
        if ($obj.PSObject.Properties['d']) { return $obj.d }
        return $obj
    }

    function Get-EDashSpCollection {
        param([string]$Url)
        $all = New-Object System.Collections.Generic.List[object]
        while ($Url) {
            $page = Invoke-EDashSpGet $Url
            if ($page.PSObject.Properties['results']) {
                foreach ($v in $page.results) { $all.Add($v) }       # odata=verbose
                $Url = $page.__next
            } else {
                foreach ($v in $page.value) { $all.Add($v) }         # odata=nometadata
                $Url = $page.'odata.nextLink'
            }
        }
        return , $all.ToArray()
    }

    # 리스트의 컬럼 정보를 읽어(캐시) 실제로 있는 컬럼만 $select 해서 항목을 가져온다.
    function Get-EDashList {
        param($Config, [string]$ListName, [string[]]$Wanted, [switch]$AllCustom)
        $site = $Config.SiteUrl
        $path = ([uri]$site).AbsolutePath.TrimEnd('/') + '/Lists/' + $ListName
        $a1   = "@a1='" + ($path -replace ' ', '%20') + "'"
        $base = $site + '/_api/web/GetList(@a1)'
        try {
            if (-not $global:EDashFieldCache.ContainsKey($ListName)) {
                $global:EDashFieldCache[$ListName] = Get-EDashSpCollection ($base + '/fields?' + $a1 + '&$filter=' + [uri]::EscapeDataString('Hidden eq false') +
                                                                           '&$select=Title,InternalName,TypeAsString,FromBaseType')
            }
            $fields  = $global:EDashFieldCache[$ListName]
            $names   = @($fields | ForEach-Object { $_.InternalName })
            $select  = New-Object System.Collections.Generic.List[string]
            $columns = New-Object System.Collections.Generic.List[object]
            foreach ($f in $fields) {
                if ($f.TypeAsString -in @('Computed', 'Attachments')) { continue }
                if (-not (($Wanted -contains $f.InternalName) -or ($AllCustom -and -not $f.FromBaseType))) { continue }
                $sel = $f.InternalName
                if ($sel.StartsWith('_')) { $sel = 'OData_' + $sel }
                if ($f.TypeAsString -in @('User', 'UserMulti', 'Lookup', 'LookupMulti')) { $sel += 'Id' }
                $select.Add($sel)
                $columns.Add([pscustomobject]@{ Select = $sel; Internal = $f.InternalName; Title = $f.Title; Type = $f.TypeAsString })
            }
            $items = Get-EDashSpCollection ($base + '/items?' + $a1 + '&$top=500&$select=' + ((@('ID', 'Modified') + $select.ToArray()) -join ','))
        } catch {
            [void]$global:EDashFieldCache.Remove($ListName)
            throw ('[{0}] {1}' -f $ListName, $_.Exception.Message)
        }
        return [pscustomobject]@{
            Items   = $items
            Columns = $columns.ToArray()
            Missing = @($Wanted | Where-Object { $names -notcontains $_ })
        }
    }

    # ----- File 원본: Power Automate가 OneDrive에 내보낸 JSON -----
    function Read-EDashJsonFile {
        param([string]$Path)
        if (-not (Test-Path -LiteralPath $Path)) { throw ('데이터 파일이 없습니다: {0}' -f $Path) }
        # OneDrive가 쓰는 중이어도 읽을 수 있게 공유 모드로 연다.
        $fs = [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        try {
            $reader = New-Object System.IO.StreamReader($fs, [Text.Encoding]::UTF8, $true)
            $text = $reader.ReadToEnd()
        } finally {
            $fs.Dispose()
        }
        if (-not $text.Trim()) { throw '데이터 파일이 비어 있습니다 (Flow가 아직 실행되지 않았거나 OneDrive 동기화 중).' }
        $text = $text -creplace '"Id":', '"Id_":'
        try { return ($text | ConvertFrom-Json) }
        catch { throw ('데이터 파일을 읽지 못했습니다 (동기화 중이면 다음 갱신 때 다시 읽습니다): ' + $_.Exception.Message) }
    }

    # 파일의 항목 속성으로 상세 패널용 컬럼 목록을 만든다 (Power Automate 메타데이터 속성은 제외).
    function Get-EDashFileColumns {
        param($Items, $Config)
        $friendly = @{}
        foreach ($x in @($Config.DashboardFields) + @($Config.ErrorColumns)) { $friendly[$x.Internal] = $(if ($x.Header) { $x.Header } else { $x.Name }) }
        $names = New-Object System.Collections.Generic.List[string]
        foreach ($it in @($Items | Select-Object -First 50)) {
            foreach ($p in $it.PSObject.Properties) {
                $n = $p.Name
                if ($n -match '^[@{]' -or $n.Contains('#') -or $n -in @('ItemInternalId', 'Id_', 'ID')) { continue }
                if (-not $names.Contains($n)) { $names.Add($n) }
            }
        }
        foreach ($n in $names) {
            $sample = $null
            foreach ($it in $Items) { $v = $it.$n; if ($null -ne $v -and "$v" -ne '') { $sample = $v; break } }
            $isDate = ($sample -is [datetime]) -or ("$sample" -match '^\d{4}-\d\d-\d\dT\d\d:\d\d')
            [pscustomobject]@{
                Select   = $n
                Internal = $n
                Title    = $(if ($friendly.ContainsKey($n)) { $friendly[$n] } else { $n })
                Type     = $(if ($isDate) { 'DateTime' } else { 'Text' })
            }
        }
    }

    function Get-EDashFileData {
        param($Config)
        $doc = Read-EDashJsonFile $Config.DataFile
        if (-not $doc.PSObject.Properties['dashboard']) { throw '데이터 파일 형식이 다릅니다 ("dashboard" 항목 없음). docs/FLOW-EXPORT.md의 식을 확인하세요.' }
        $dash   = @($doc.dashboard | Where-Object { $_ })
        $alerts = @($doc.alerts | Where-Object { $_ })
        return @{
            Dash        = [pscustomobject]@{ Items = $dash;   Columns = @(Get-EDashFileColumns $dash $Config);   Missing = @() }
            Alert       = [pscustomobject]@{ Items = $alerts; Columns = @(Get-EDashFileColumns $alerts $Config); Missing = @() }
            GeneratedAt = ConvertTo-EDashKst $doc.generatedAt
        }
    }

    # ----- DEMO 데이터 (SharePoint 없이 화면 확인용) -----
    function Set-EDashDemoError {
        param($Item, [string]$Field, [datetime]$AtKst)
        $iso = $AtKst.AddHours(-9).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
        $Item.$Field              = $AtKst.ToString('yyyy-MM-dd HH:mm')
        $Item.AllOk               = 'No'
        $Item.LastUpdatedErrorTime = $iso
        $Item.field_14            = $iso
        if (-not $Item.StartMonitoring) { $Item.StartMonitoring = $iso }
        $Item.ActiveErrorCount    = @('field_6', 'field_7', 'field_8', 'field_9' | Where-Object { $Item.$_ }).Count
        if ($Field -eq 'field_10') { $Item.ReminderCount = [int]$Item.ReminderCount + 1 }
        foreach ($a in $global:EDashDemo.Alerts) {
            if ($a.SystemReferenceNumber -eq $Item.SRN -and $Field -ne 'field_10') { $a.AlertSent = 'true' }
        }
    }

    function Clear-EDashDemoError {
        param($Item, [datetime]$AtKst)
        foreach ($f in 'field_6', 'field_7', 'field_8', 'field_9', 'field_10') { $Item.$f = $null }
        $Item.AllOk            = 'Yes'
        $Item.field_16         = $AtKst.AddHours(-9).ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
        $Item.StartMonitoring  = $null
        $Item.ActiveErrorCount = 0
        $Item.ReminderCount    = 0
        foreach ($a in $global:EDashDemo.Alerts) { if ($a.SystemReferenceNumber -eq $Item.SRN) { $a.AlertSent = 'false' } }
    }

    function Get-EDashDemoData {
        param($Config)
        $now = Get-EDashKstNow
        if (-not $global:EDashDemo) {
            $names   = '가람', '나래', '다온', '라온', '마루', '바다', '새봄', '아름', '자람', '하늘', '한빛', '푸른', '온누리', '미래'
            $regions = '서울', '경기', '인천', '부산', '대구', '광주', '대전'
            $models  = 'Ingenia 1.5T', 'Ingenia 3.0T', 'Achieva 1.5T', 'Prodiva 1.5T', 'Elition 3.0T', 'Ingenia Ambition'
            $dash    = New-Object System.Collections.Generic.List[object]
            $alerts  = New-Object System.Collections.Generic.List[object]
            for ($i = 1; $i -le 28; $i++) {
                $srn = [string](502000 + $i * 37)
                $dash.Add([pscustomobject][ordered]@{
                    ID = $i; SRN = $srn
                    field_2 = $names[$i % $names.Count] + $(if ($i % 3) { '병원' } else { '의원' })
                    field_3 = $models[$i % $models.Count]; field_4 = 'FSE-' + [char](65 + $i % 6); field_5 = $regions[$i % $regions.Count]
                    field_6 = $null; field_7 = $null; field_8 = $null; field_9 = $null; field_10 = $null
                    field_14 = $null; field_16 = $null; AllOk = 'Yes'; StartMonitoring = $null
                    ActiveErrorCount = 0; ReminderCount = 0; LastUpdatedErrorTime = $null
                })
                if ($i -ne 5) { $alerts.Add([pscustomobject]@{ ID = $i; SystemReferenceNumber = $srn; AlertSent = 'false' }) }
            }
            $alerts.Add([pscustomobject]@{ ID = 99; SystemReferenceNumber = '599999'; AlertSent = 'false' })
            $global:EDashDemo = @{ Rnd = (New-Object System.Random 20261022); Dash = $dash; Alerts = $alerts }
            Set-EDashDemoError $dash[2]  'field_6'  $now.AddMinutes(-27)
            Set-EDashDemoError $dash[7]  'field_9'  $now.AddMinutes(-95)
            Set-EDashDemoError $dash[11] 'field_7'  $now.AddHours(-8)        # 오래된 에러 -> AllOk 누락 의심
            Set-EDashDemoError $dash[15] 'field_10' $now.AddHours(-3)
            Set-EDashDemoError $dash[20] 'field_10' $now.AddHours(-5)
            Set-EDashDemoError $dash[23] 'field_8'  $now.AddMinutes(-12)
        } else {
            # 갱신할 때마다 한 대의 상태를 바꿔서 알림·색상 변화를 볼 수 있게 한다.
            $d    = $global:EDashDemo
            $item = $d.Dash[$d.Rnd.Next($d.Dash.Count)]
            $roll = $d.Rnd.Next(10)
            if ($roll -lt 4)     { Set-EDashDemoError $item (@('field_6', 'field_7', 'field_8', 'field_9')[$d.Rnd.Next(4)]) $now }
            elseif ($roll -lt 6) { Set-EDashDemoError $item 'field_10' $now }
            else                 { Clear-EDashDemoError $item $now }
        }
        $dashCols = @(@($Config.DashboardFields) + @($Config.ErrorColumns) | ForEach-Object {
            [pscustomobject]@{ Select = $_.Internal; Internal = $_.Internal; Title = $(if ($_.Header) { $_.Header } else { $_.Name }); Type = 'Text' }
        })
        $alertCols = @(
            [pscustomobject]@{ Select = 'SystemReferenceNumber'; Internal = 'SystemReferenceNumber'; Title = 'SystemReferenceNumber'; Type = 'Text' }
            [pscustomobject]@{ Select = 'AlertSent'; Internal = 'AlertSent'; Title = 'AlertSent'; Type = 'Text' }
        )
        return @{
            Dash  = [pscustomobject]@{ Items = $global:EDashDemo.Dash.ToArray(); Columns = $dashCols; Missing = @() }
            Alert = [pscustomobject]@{ Items = $global:EDashDemo.Alerts.ToArray(); Columns = $alertCols; Missing = @() }
        }
    }

    # ----- 스냅샷: 두 리스트 조회 -> 정규화 -> SRN 조인 -> 상태 판정 -----
    function Get-EDashSnapshot {
        param($Config)
        $sw = [Diagnostics.Stopwatch]::StartNew()
        if ($Config.Demo) {
            $data = Get-EDashDemoData $Config
        } elseif ($Config.Source -eq 'File') {
            $data = Get-EDashFileData $Config
        } else {
            $wanted = @($Config.DashboardFields | ForEach-Object { $_.Internal }) + @($Config.ErrorColumns | ForEach-Object { $_.Internal })
            $data = @{
                Dash  = Get-EDashList $Config $Config.DashboardList $wanted
                Alert = Get-EDashList $Config $Config.AlertList @('SystemReferenceNumber', 'AlertSent') -AllCustom
            }
        }
        $now = Get-EDashKstNow
        $f = @{}
        foreach ($x in $Config.DashboardFields) { $f[$x.Name] = $x.Internal }

        $alertByKey = @{}
        foreach ($a in $data.Alert.Items) {
            $k = Get-EDashSrnKey (ConvertTo-EDashText $a.SystemReferenceNumber)
            if ($k -and -not $alertByKey.ContainsKey($k)) { $alertByKey[$k] = $a }
        }
        $keyCount = @{}
        foreach ($it in $data.Dash.Items) {
            $k = Get-EDashSrnKey (ConvertTo-EDashText $it.($f['SRN']))
            if ($k) { $keyCount[$k] = 1 + [int]$keyCount[$k] }
        }

        $systems = New-Object System.Collections.Generic.List[object]
        $used    = @{}
        $seen    = @{}
        foreach ($it in $data.Dash.Items) {
            $srnRaw = ConvertTo-EDashText $it.($f['SRN'])
            $srnKey = Get-EDashSrnKey $srnRaw
            $key    = if ($srnKey) { $srnKey } else { 'ID-' + $it.ID }
            if ($seen.ContainsKey($key)) { $key = $key + '#' + $it.ID }
            $seen[$key] = $true

            $alert = $null
            if ($srnKey -and $alertByKey.ContainsKey($srnKey)) { $alert = $alertByKey[$srnKey]; $used[$srnKey] = $true }

            $errs   = [ordered]@{}
            $groups = New-Object System.Collections.Generic.List[string]
            $hasCrit = $false; $hasRem = $false; $active = 0
            foreach ($c in $Config.ErrorColumns) {
                $on = Test-EDashActive $it.($c.Internal)
                $errs[$c.Name] = $on
                if (-not $on) { continue }
                $active++
                if ($c.Critical) {
                    $hasCrit = $true
                    if (-not $groups.Contains($c.Group)) { $groups.Add($c.Group) }
                } else {
                    $hasRem = $true
                }
            }

            $allOk     = ConvertTo-EDashBool $it.($f['AllOk'])
            $allOkTime = ConvertTo-EDashKst $it.($f['AllOkTime'])
            $lastErr   = ConvertTo-EDashKst $it.($f['LastUpdatedErrorTime'])
            $lastUpd   = ConvertTo-EDashKst $it.($f['LastUpdated'])
            $startMon  = ConvertTo-EDashKst $it.($f['StartMonitoring'])
            $errTime   = if ($lastErr) { $lastErr } else { $lastUpd }

            # 상태 판정: AllOk가 마지막 에러보다 나중이면 정상, 아니면 에러 컬럼 기준
            if ($active -eq 0)                                                    { $status = 'Ok' }
            elseif ($allOk -and $allOkTime -and (-not $errTime -or $allOkTime -ge $errTime)) { $status = 'Ok' }
            elseif ($allOk -and -not $allOkTime -and -not $errTime)               { $status = 'Ok' }
            elseif ($hasCrit)                                                     { $status = 'Critical' }
            else                                                                  { $status = 'Reminder' }

            $stale = ($status -ne 'Ok') -and $errTime -and (($now - $errTime).TotalHours -ge $Config.StaleHours)
            $since = if ($startMon) { $startMon } else { $errTime }

            $systems.Add([pscustomobject]@{
                Key              = $key
                SrnKey           = $srnKey
                SRN              = ($srnRaw -replace '"', '').Trim()
                DashId           = $it.ID
                AlertId          = $(if ($alert) { $alert.ID } else { $null })
                Site             = ConvertTo-EDashText $it.($f['SiteName'])
                System           = ConvertTo-EDashText $it.($f['SystemName'])
                FSE              = ConvertTo-EDashText $it.($f['FSE'])
                Region           = ConvertTo-EDashText $it.($f['Region'])
                Errors           = $errs
                Groups           = ($groups -join ', ')
                ActiveCount      = $active
                ActiveErrorCount = ConvertTo-EDashInt $it.($f['ActiveErrorCount'])
                ReminderCount    = ConvertTo-EDashInt $it.($f['ReminderCount'])
                AllOk            = $allOk
                AllOkTime        = $allOkTime
                LastUpdated      = $lastUpd
                LastErr          = $errTime
                Since            = $since
                AlertSent        = $(if ($alert) { ConvertTo-EDashBool $alert.AlertSent } else { $false })
                Status           = $status
                Stale            = [bool]$stale
                Duplicate        = ($srnKey -and $keyCount[$srnKey] -gt 1)
                DashItem         = $it
                AlertItem        = $alert
            })
        }

        # EASDB2에만 있는 SRN
        foreach ($k in @($alertByKey.Keys)) {
            if ($used.ContainsKey($k)) { continue }
            $a = $alertByKey[$k]
            $systems.Add([pscustomobject]@{
                Key = 'A-' + $k; SrnKey = $k; SRN = ((ConvertTo-EDashText $a.SystemReferenceNumber) -replace '"', '').Trim()
                DashId = $null; AlertId = $a.ID; Site = ''; System = ''; FSE = ''; Region = ''
                Errors = [ordered]@{}; Groups = ''; ActiveCount = 0; ActiveErrorCount = $null; ReminderCount = $null
                AllOk = $false; AllOkTime = $null; LastUpdated = $null; LastErr = $null; Since = $null
                AlertSent = (ConvertTo-EDashBool $a.AlertSent); Status = 'Unknown'; Stale = $false; Duplicate = $false
                DashItem = $null; AlertItem = $a
            })
        }

        $missing = @(@($data.Dash.Missing | ForEach-Object { '{0}.{1}' -f $Config.DashboardList, $_ }) +
                     @($data.Alert.Missing | ForEach-Object { '{0}.{1}' -f $Config.AlertList, $_ }))
        $sw.Stop()
        return [pscustomobject]@{
            TakenAt      = $now
            Systems      = $systems.ToArray()
            DashColumns  = $data.Dash.Columns
            AlertColumns = $data.Alert.Columns
            Missing      = $missing
            DurationMs   = $sw.ElapsedMilliseconds
            Demo         = [bool]$Config.Demo
            Source       = $(if ($Config.Demo) { 'Demo' } else { $Config.Source })
            GeneratedAt  = $(if ($Config.Demo) { $now } else { $data.GeneratedAt })
            Accept       = $(if ($Config.Demo) { 'sample data' } elseif ($Config.Source -eq 'Rest') { $global:EDashAccept } else { $Config.DataFile })
        }
    }
}

# 백그라운드 Runspace: 데이터 코드를 한 번 로드해 두고, 갱신마다 Get-EDashSnapshot만 비동기 실행한다.
function New-EDashWorker {
    $rs = [runspacefactory]::CreateRunspace()
    $rs.Open()
    $init = [powershell]::Create()
    $init.Runspace = $rs
    [void]$init.AddScript($DataCode.ToString()).Invoke()
    $init.Dispose()
    return $rs
}

function Start-EDashJob {
    param($Runspace, $Config)
    $ps = [powershell]::Create()
    $ps.Runspace = $Runspace
    [void]$ps.AddCommand('Get-EDashSnapshot').AddParameter('Config', $Config)
    return @{ PS = $ps; Handle = $ps.BeginInvoke() }
}

function Complete-EDashJob {
    param($Job)
    try {
        $out = $Job.PS.EndInvoke($Job.Handle)
        $snap = $out | Select-Object -Last 1
        if (-not $snap) {
            $first = $Job.PS.Streams.Error | Select-Object -First 1
            if ($first) { throw $first.Exception.Message }
            throw '데이터를 받지 못했습니다.'
        }
        return $snap
    } catch {
        $ex = $_.Exception
        while ($ex.InnerException) { $ex = $ex.InnerException }
        throw $ex.Message
    } finally {
        $Job.PS.Dispose()
    }
}

. $DataCode

# ---------------------------------------------------------------------------------------------
# -CheckOnly: 화면 없이 한 번 읽어서 요약 출력 (연결 확인용)
# ---------------------------------------------------------------------------------------------
if ($CheckOnly) {
    $what = if ($Demo) { '[DEMO]' } elseif ($Source -eq 'File') { $Config.DataFile } else { $Config.SiteUrl }
    Write-Host ('E-Dashboard 연결 확인  {0}' -f $what)
    $rs = New-EDashWorker
    try {
        $job = Start-EDashJob $rs $Config
        while (-not $job.Handle.IsCompleted) { Start-Sleep -Milliseconds 100 }
        $snap = Complete-EDashJob $job
    } catch {
        Write-Host ('실패: ' + $_.Exception.Message) -ForegroundColor Red
        exit 1
    } finally {
        $rs.Dispose()
    }
    $sys = @($snap.Systems)
    Write-Host ('{0}건 / 조회 {1}ms / {2}' -f $sys.Count, $snap.DurationMs, $snap.Accept) -ForegroundColor Green
    if ($snap.Source -eq 'File') {
        if ($snap.GeneratedAt) { Write-Host ('데이터 생성 {0:yyyy-MM-dd HH:mm:ss} KST ({1:n0}분 전)' -f $snap.GeneratedAt, ($snap.TakenAt - $snap.GeneratedAt).TotalMinutes) }
        else { Write-Host '데이터 생성 시각(generatedAt)이 파일에 없습니다.' -ForegroundColor Yellow }
    }
    foreach ($g in ($sys | Group-Object Status)) { Write-Host ('  {0,-9} {1}' -f $g.Name, $g.Count) }
    Write-Host ('  AllOk 누락 의심 {0} / Teams 발송 {1} / 중복 SRN 행 {2}' -f @($sys | Where-Object Stale).Count, @($sys | Where-Object AlertSent).Count, @($sys | Where-Object Duplicate).Count)
    if ($snap.Missing.Count) { Write-Host ('없는 컬럼: ' + ($snap.Missing -join ', ')) -ForegroundColor Yellow }
    $rank = @{ Critical = 0; Reminder = 1; Unknown = 2; Ok = 3 }
    $sys | Sort-Object @{ Expression = { $rank[$_.Status] } }, @{ Expression = { $_.LastErr }; Descending = $true } |
        Select-Object -First 15 Status, Stale, SRN, Region, FSE, Groups, LastErr, AlertSent |
        Format-Table -AutoSize | Out-String -Width 200 | Write-Host
    exit 0
}

# ---------------------------------------------------------------------------------------------
# 화면 (WinForms)
# ---------------------------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

$script:Mutex = New-Object System.Threading.Mutex($false, ('Local\EDashboard' + $(if ($Demo) { '-Demo' } else { '' })))
if (-not $script:Mutex.WaitOne(0)) {
    [void][System.Windows.Forms.MessageBox]::Show('E-Dashboard가 이미 실행 중입니다.', 'E-Dashboard')
    exit 0
}

$script:Config   = $Config
$script:LogDir   = Join-Path $(if ($env:LOCALAPPDATA) { $env:LOCALAPPDATA } else { [IO.Path]::GetTempPath() }) 'EDashboard'
$script:Snapshot = $null
$script:SystemsByKey   = @{}
$script:Job            = $null
$script:TileFilter     = 'All'
$script:SuppressFilter = $false
$script:NextRefreshAt  = Get-Date
$script:ErrorColumnNames = @($Config.ErrorColumns | ForEach-Object { $_.Name })
$script:StatusRank = @{ Critical = 0; Reminder = 1; Unknown = 2; Ok = 3 }

function Write-EDashLog {
    param([string]$Message)
    try {
        if (-not (Test-Path -LiteralPath $script:LogDir)) { [void](New-Item -ItemType Directory -Path $script:LogDir) }
        $file = Join-Path $script:LogDir ('edashboard-{0:yyyyMMdd}.log' -f (Get-Date))
        Add-Content -LiteralPath $file -Value ('{0:HH:mm:ss} {1}' -f (Get-Date), $Message) -Encoding UTF8
    } catch { }
}

function New-EDashColor { param([int]$R, [int]$G, [int]$B) [System.Drawing.Color]::FromArgb($R, $G, $B) }

$Theme = @{
    Font       = New-Object System.Drawing.Font('Malgun Gothic', 9)
    Bold       = New-Object System.Drawing.Font('Malgun Gothic', 9, [System.Drawing.FontStyle]::Bold)
    Italic     = New-Object System.Drawing.Font('Malgun Gothic', 9, [System.Drawing.FontStyle]::Italic)
    Title      = New-Object System.Drawing.Font('Malgun Gothic', 13, [System.Drawing.FontStyle]::Bold)
    Big        = New-Object System.Drawing.Font('Malgun Gothic', 18, [System.Drawing.FontStyle]::Bold)
    Mono       = New-Object System.Drawing.Font('Consolas', 9.5)
    HeaderBack = New-EDashColor 30 45 70
    HeaderFore = [System.Drawing.Color]::White
    Muted      = New-EDashColor 110 110 110
    Good       = New-EDashColor 120 220 140
    Bad        = New-EDashColor 255 140 140
    ErrorDot   = New-EDashColor 190 0 0
    StaleFore  = New-EDashColor 120 60 160
    TileOn     = New-EDashColor 225 236 250
}
$script:StatusBack = @{
    Critical = New-EDashColor 255 222 222
    Reminder = New-EDashColor 255 241 208
    Unknown  = New-EDashColor 232 232 240
}

function Format-EDashAge {
    param([TimeSpan]$Span)
    if ($Span.TotalMinutes -lt 1) { return '방금' }
    if ($Span.TotalHours -lt 1)   { return '{0}분' -f [int][Math]::Floor($Span.TotalMinutes) }
    if ($Span.TotalDays -lt 1)    { return '{0}시간 {1}분' -f [int][Math]::Floor($Span.TotalHours), $Span.Minutes }
    return '{0}일 {1}시간' -f [int][Math]::Floor($Span.TotalDays), $Span.Hours
}

function Format-EDashTime {
    param($Value)
    if ($Value -is [datetime]) { return $Value.ToString('yyyy-MM-dd HH:mm') }
    return '-'
}

function Get-EDashStatusText {
    param($System)
    switch ($System.Status) {
        'Critical' { $t = '위험 ' + $System.Groups }
        'Reminder' { $t = 'Reminder' }
        'Unknown'  { $t = '미매칭 (EASDB2만)' }
        default    { $t = '정상' }
    }
    if ($System.Stale)     { $t += ' · 갱신없음' }
    if ($System.Duplicate) { $t += ' [중복]' }
    return $t
}

function New-EDashLabel {
    param([string]$Text, $Font, $Color, [int]$Top = 12)
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $Text
    $l.AutoSize = $true
    $l.Margin = New-Object System.Windows.Forms.Padding(8, $Top, 8, 0)
    if ($Font)  { $l.Font = $Font }
    if ($Color) { $l.ForeColor = $Color }
    return $l
}

# ----- 폼 -----
$form = New-Object System.Windows.Forms.Form
$form.Text = 'E-Dashboard' + $(if ($Demo) { '   [DEMO - 샘플 데이터]' } elseif ($Source -eq 'File') { '   ' + $Config.DataFile } else { '   ' + $Config.SiteUrl })
$form.Size = New-Object System.Drawing.Size(1420, 860)
$form.MinimumSize = New-Object System.Drawing.Size(900, 500)
$form.StartPosition = 'CenterScreen'
$form.Font = $Theme.Font
$form.KeyPreview = $true
$script:Form = $form

$layout = New-Object System.Windows.Forms.TableLayoutPanel
$layout.Dock = 'Fill'
$layout.ColumnCount = 1
$layout.RowCount = 4
[void]$layout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
foreach ($h in 46, 86, 38) { [void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $h))) }
[void]$layout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))

# 헤더
$header = New-Object System.Windows.Forms.FlowLayoutPanel
$header.Dock = 'Fill'
$header.Margin = New-Object System.Windows.Forms.Padding(0)
$header.BackColor = $Theme.HeaderBack
$header.WrapContents = $false
$header.Controls.Add((New-EDashLabel 'E-Dashboard' $Theme.Title $Theme.HeaderFore 9))
$script:LblConn     = New-EDashLabel '● 연결 중...' $Theme.Bold $Theme.HeaderFore 14
$script:LblUpdated  = New-EDashLabel '' $null $Theme.HeaderFore 14
$script:LblDayNight = New-EDashLabel '' $Theme.Bold $Theme.HeaderFore 14
$script:LblReset    = New-EDashLabel '' $null $Theme.HeaderFore 14
$btnRefresh = New-Object System.Windows.Forms.Button
$btnRefresh.Text = '새로고침 (F5)'
$btnRefresh.AutoSize = $true
$btnRefresh.BackColor = [System.Drawing.SystemColors]::Control
$btnRefresh.Margin = New-Object System.Windows.Forms.Padding(16, 9, 8, 0)
$header.Controls.AddRange(@($script:LblConn, $script:LblUpdated, $script:LblDayNight, $script:LblReset, $btnRefresh))

# KPI 타일
$kpi = New-Object System.Windows.Forms.FlowLayoutPanel
$kpi.Dock = 'Fill'
$kpi.Margin = New-Object System.Windows.Forms.Padding(0)
$kpi.BackColor = New-EDashColor 244 246 249
$kpi.WrapContents = $false
$script:Tiles = @{}

function New-EDashTile {
    param([string]$Key, [string]$Caption, $Accent)
    $p = New-Object System.Windows.Forms.Panel
    $p.Size = New-Object System.Drawing.Size(168, 70)
    $p.Margin = New-Object System.Windows.Forms.Padding(8, 8, 0, 8)
    $p.BackColor = [System.Drawing.Color]::White
    $p.BorderStyle = 'FixedSingle'
    $p.Cursor = [System.Windows.Forms.Cursors]::Hand
    $bar = New-Object System.Windows.Forms.Panel
    $bar.Dock = 'Left'
    $bar.Width = 6
    $bar.BackColor = $Accent
    $cap = New-Object System.Windows.Forms.Label
    $cap.Text = $Caption
    $cap.AutoSize = $true
    $cap.ForeColor = $Theme.Muted
    $cap.Location = New-Object System.Drawing.Point(14, 6)
    $val = New-Object System.Windows.Forms.Label
    $val.Text = '-'
    $val.AutoSize = $true
    $val.Font = $Theme.Big
    $val.Location = New-Object System.Drawing.Point(12, 26)
    $p.Controls.AddRange(@($cap, $val, $bar))
    foreach ($c in @($p, $cap, $val, $bar)) {
        $c.Tag = $Key
        $c.Add_Click({ Set-EDashTileFilter ([string]$this.Tag) })
    }
    $script:Tiles[$Key] = @{ Panel = $p; Value = $val }
    return $p
}

$kpi.Controls.Add((New-EDashTile 'All'      '전체'               (New-EDashColor 120 130 145)))
$kpi.Controls.Add((New-EDashTile 'Critical' '위험'               (New-EDashColor 210 40 40)))
$kpi.Controls.Add((New-EDashTile 'Reminder' 'Reminder'           (New-EDashColor 235 150 20)))
$kpi.Controls.Add((New-EDashTile 'Stale'    'AllOk 누락 의심'    (New-EDashColor 130 70 180)))
$kpi.Controls.Add((New-EDashTile 'Ok'       '정상'               (New-EDashColor 40 160 80)))
$kpi.Controls.Add((New-EDashTile 'Sent'     'Teams 발송'         (New-EDashColor 50 110 200)))
$kpi.Controls.Add((New-EDashTile 'Unknown'  '미매칭 (EASDB2만)'  (New-EDashColor 150 150 165)))

# 필터 바
$filterBar = New-Object System.Windows.Forms.FlowLayoutPanel
$filterBar.Dock = 'Fill'
$filterBar.Margin = New-Object System.Windows.Forms.Padding(0)
$filterBar.WrapContents = $false

function New-EDashFilterLabel { param([string]$Text) New-EDashLabel $Text $null $null 10 }

$script:TxtSearch = New-Object System.Windows.Forms.TextBox
$script:TxtSearch.Width = 180
$script:TxtSearch.Margin = New-Object System.Windows.Forms.Padding(3, 7, 10, 0)
$script:CboRegion = New-Object System.Windows.Forms.ComboBox
$script:CboRegion.DropDownStyle = 'DropDownList'
$script:CboRegion.Width = 110
$script:CboRegion.Margin = New-Object System.Windows.Forms.Padding(3, 7, 10, 0)
[void]$script:CboRegion.Items.Add('(전체)')
$script:CboRegion.SelectedIndex = 0
$script:CboFse = New-Object System.Windows.Forms.ComboBox
$script:CboFse.DropDownStyle = 'DropDownList'
$script:CboFse.Width = 130
$script:CboFse.Margin = New-Object System.Windows.Forms.Padding(3, 7, 10, 0)
[void]$script:CboFse.Items.Add('(전체)')
$script:CboFse.SelectedIndex = 0
$script:ChkErrors = New-Object System.Windows.Forms.CheckBox
$script:ChkErrors.Text = '에러만'
$script:ChkErrors.AutoSize = $true
$script:ChkErrors.Margin = New-Object System.Windows.Forms.Padding(6, 9, 10, 0)
$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = '필터 초기화'
$btnClear.AutoSize = $true
$btnClear.Margin = New-Object System.Windows.Forms.Padding(3, 6, 3, 0)
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'CSV 내보내기'
$btnExport.AutoSize = $true
$btnExport.Margin = New-Object System.Windows.Forms.Padding(3, 6, 16, 0)
$script:ChkNotify = New-Object System.Windows.Forms.CheckBox
$script:ChkNotify.Text = '새 위험 알림 (트레이)'
$script:ChkNotify.AutoSize = $true
$script:ChkNotify.Checked = $true
$script:ChkNotify.Margin = New-Object System.Windows.Forms.Padding(6, 9, 10, 0)
$filterBar.Controls.AddRange(@(
    (New-EDashFilterLabel '검색'), $script:TxtSearch,
    (New-EDashFilterLabel 'Region'), $script:CboRegion,
    (New-EDashFilterLabel 'FSE'), $script:CboFse,
    $script:ChkErrors, $btnClear, $btnExport, $script:ChkNotify))

# 그리드 + 상세
$split = New-Object System.Windows.Forms.SplitContainer
$split.Dock = 'Fill'
$split.Orientation = 'Horizontal'
$split.Margin = New-Object System.Windows.Forms.Padding(0)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.AllowUserToResizeRows = $false
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.RowHeadersVisible = $false
$grid.AutoGenerateColumns = $true
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.BorderStyle = 'None'
$grid.ColumnHeadersHeightSizeMode = 'AutoSize'
$grid.RowTemplate.Height = 24
$grid.DefaultCellStyle.SelectionBackColor = New-EDashColor 51 102 170
$grid.DefaultCellStyle.SelectionForeColor = [System.Drawing.Color]::White
try { $grid.GetType().GetProperty('DoubleBuffered', [Reflection.BindingFlags]'NonPublic,Instance').SetValue($grid, $true, $null) } catch { }
$script:Grid = $grid

$script:Detail = New-Object System.Windows.Forms.TextBox
$script:Detail.Dock = 'Fill'
$script:Detail.Multiline = $true
$script:Detail.ReadOnly = $true
$script:Detail.ScrollBars = 'Vertical'
$script:Detail.Font = $Theme.Mono
$script:Detail.BackColor = New-EDashColor 250 250 250

$split.Panel1.Controls.Add($grid)
$split.Panel2.Controls.Add($script:Detail)

$status = New-Object System.Windows.Forms.StatusStrip
$script:StatusCount = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:StatusMsg   = New-Object System.Windows.Forms.ToolStripStatusLabel
$script:StatusMsg.Spring = $true
$script:StatusMsg.TextAlign = 'MiddleLeft'
$script:StatusNext  = New-Object System.Windows.Forms.ToolStripStatusLabel
[void]$status.Items.AddRange(@($script:StatusCount, $script:StatusMsg, $script:StatusNext))

$layout.Controls.Add($header, 0, 0)
$layout.Controls.Add($kpi, 0, 1)
$layout.Controls.Add($filterBar, 0, 2)
$layout.Controls.Add($split, 0, 3)
$form.Controls.Add($layout)
$form.Controls.Add($status)
$layout.BringToFront()

# ----- 데이터 테이블 (그리드 바인딩) -----
$script:Table = New-Object System.Data.DataTable 'Systems'
$colDefs = @(
    , @('Key', [string]); , @('StatusCode', [string]); , @('Rank', [int]); , @('Stale', [bool]); , @('Sent', [bool])
    , @('StatusText', [string]); , @('SRN', [string]); , @('Site', [string]); , @('System', [string]); , @('FSE', [string]); , @('Region', [string])
)
foreach ($c in $Config.ErrorColumns) { $colDefs += , @($c.Name, [string]) }
$colDefs += , @('ActiveCount', [int])
$colDefs += , @('ReminderCount', [int])
$colDefs += , @('LastErr', [datetime])
$colDefs += , @('Age', [string])
$colDefs += , @('AllOkText', [string])
$colDefs += , @('SentText', [string])
foreach ($d in $colDefs) { [void]$script:Table.Columns.Add($d[0], $d[1]) }
$script:Table.Columns['Stale'].DefaultValue = $false
$script:Table.Columns['Sent'].DefaultValue = $false
$script:Table.PrimaryKey = @($script:Table.Columns['Key'])
$script:View = New-Object System.Data.DataView($script:Table)
$script:View.Sort = 'Rank ASC, LastErr DESC'

$script:Headers = [ordered]@{ StatusText = '상태'; SRN = 'SRN'; Site = '사이트'; System = '시스템'; FSE = 'FSE'; Region = 'Region' }
foreach ($c in $Config.ErrorColumns) { $script:Headers[$c.Name] = $c.Header }
$script:Headers['ActiveCount']   = '활성'
$script:Headers['ReminderCount'] = 'Rem수'
$script:Headers['LastErr']       = '최근 에러'
$script:Headers['Age']           = '경과'
$script:Headers['AllOkText']     = 'AllOk'
$script:Headers['SentText']      = 'Teams'
$script:Widths = @{ StatusText = 150; SRN = 90; Site = 170; System = 130; FSE = 80; Region = 70; ActiveCount = 45; ReminderCount = 50; LastErr = 95; Age = 95; AllOkText = 50; SentText = 55 }

function Set-EDashCell {
    param($Row, [string]$Name, $Value)
    if ($null -eq $Value) { $Value = [DBNull]::Value }
    if (-not $Row[$Name].Equals($Value)) { $Row[$Name] = $Value }
}

# ----- 화면 갱신 -----
function Get-EDashSelectedSystem {
    $r = $script:Grid.CurrentRow
    if (-not $r -or -not $r.DataBoundItem) { return $null }
    return $script:SystemsByKey[[string]$r.DataBoundItem.Row['Key']]
}

function Select-EDashKey {
    param([string]$Key)
    foreach ($r in $script:Grid.Rows) {
        if ([string]$r.DataBoundItem.Row['Key'] -eq $Key) { $script:Grid.CurrentCell = $r.Cells['SRN']; return }
    }
}

function Update-EDashDetail {
    $s = Get-EDashSelectedSystem
    if (-not $s) { $script:Detail.Text = ''; return }
    $now = Get-EDashKstNow
    $sb  = New-Object System.Text.StringBuilder
    $add = { param($k, $v) [void]$sb.AppendLine(('  {0,-26} {1}' -f $k, $v)) }
    [void]$sb.AppendLine('[요약]')
    & $add '상태' (Get-EDashStatusText $s)
    & $add 'SRN' $s.SRN
    & $add '사이트 / 시스템' ('{0} / {1}' -f $s.Site, $s.System)
    & $add 'FSE / Region' ('{0} / {1}' -f $s.FSE, $s.Region)
    $on = @($script:Config.ErrorColumns | Where-Object { $s.Errors[$_.Name] } | ForEach-Object { $_.Header })
    & $add '활성 에러' $(if ($on.Count) { $on -join ', ' } else { '-' })
    & $add '에러 시작 / 최근' ('{0} / {1}' -f (Format-EDashTime $s.Since), (Format-EDashTime $s.LastErr))
    if ($s.Status -ne 'Ok' -and $s.Since) { & $add '경과' (Format-EDashAge ($now - $s.Since)) }
    if ($s.Stale) { & $add 'AllOk 누락 의심' ('마지막 에러 후 {0}시간 이상 갱신 없음' -f $script:Config.StaleHours) }
    & $add 'AllOk / AllOk 시각' ('{0} / {1}' -f $(if ($s.AllOk) { '예' } else { '아니요' }), (Format-EDashTime $s.AllOkTime))
    & $add 'Teams 알림 (EASDB2)' $(if ($s.AlertItem) { if ($s.AlertSent) { '발송됨' } else { '미발송' } } else { 'EASDB2 항목 없음' })
    if ($s.Duplicate) { & $add '중복' '같은 SRN이 EasDashboard에 여러 행 있음' }

    foreach ($part in @(
        @{ Title = $script:Config.DashboardList; Item = $s.DashItem; Columns = $script:Snapshot.DashColumns; Id = $s.DashId },
        @{ Title = $script:Config.AlertList;     Item = $s.AlertItem; Columns = $script:Snapshot.AlertColumns; Id = $s.AlertId })) {
        [void]$sb.AppendLine()
        if (-not $part.Item) { [void]$sb.AppendLine(('[{0}] 항목 없음' -f $part.Title)); continue }
        [void]$sb.AppendLine(('[{0}  ID {1}]' -f $part.Title, $part.Id))
        foreach ($c in $part.Columns) {
            $v = $part.Item.($c.Select)
            if ($c.Type -eq 'DateTime') { $text = Format-EDashTime (ConvertTo-EDashKst $v) } else { $text = ConvertTo-EDashText $v }
            & $add ('{0} ({1})' -f $c.Title, $c.Internal) $text
        }
    }
    $script:Detail.Text = $sb.ToString()
}

function Update-EDashCombo {
    param($Combo, [string[]]$Values)
    $current = [string]$Combo.SelectedItem
    $Combo.BeginUpdate()
    $Combo.Items.Clear()
    [void]$Combo.Items.Add('(전체)')
    foreach ($v in $Values) { [void]$Combo.Items.Add($v) }
    $Combo.EndUpdate()
    $idx = $Combo.Items.IndexOf($current)
    if ($idx -lt 0) { $idx = 0 }
    $Combo.SelectedIndex = $idx
}

function Protect-EDashLike {
    param([string]$Text)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        if ('*%[]'.IndexOf($ch) -ge 0) { [void]$sb.Append('[').Append($ch).Append(']') }
        elseif ($ch -eq [char]"'")     { [void]$sb.Append("''") }
        else                           { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function Update-EDashFilter {
    if ($script:SuppressFilter) { return }
    $parts = New-Object System.Collections.Generic.List[string]
    switch ($script:TileFilter) {
        'Critical' { $parts.Add("[StatusCode] = 'Critical'") }
        'Reminder' { $parts.Add("[StatusCode] = 'Reminder'") }
        'Ok'       { $parts.Add("[StatusCode] = 'Ok'") }
        'Unknown'  { $parts.Add("[StatusCode] = 'Unknown'") }
        'Stale'    { $parts.Add('[Stale] = true') }
        'Sent'     { $parts.Add('[Sent] = true') }
    }
    if ($script:ChkErrors.Checked) { $parts.Add("[StatusCode] <> 'Ok'") }
    if ($script:CboRegion.SelectedIndex -gt 0) { $parts.Add("[Region] = '" + ([string]$script:CboRegion.SelectedItem).Replace("'", "''") + "'") }
    if ($script:CboFse.SelectedIndex -gt 0)    { $parts.Add("[FSE] = '" + ([string]$script:CboFse.SelectedItem).Replace("'", "''") + "'") }
    $q = $script:TxtSearch.Text.Trim()
    if ($q) {
        $like = Protect-EDashLike $q
        $parts.Add(("([SRN] LIKE '%{0}%' OR [Site] LIKE '%{0}%' OR [System] LIKE '%{0}%' OR [FSE] LIKE '%{0}%')" -f $like))
    }
    $script:View.RowFilter = ($parts -join ' AND ')
    $script:StatusCount.Text = '표시 {0} / 전체 {1}' -f $script:View.Count, $script:Table.Rows.Count
}

function Set-EDashTileFilter {
    param([string]$Key)
    if ($script:TileFilter -eq $Key -and $Key -ne 'All') { $Key = 'All' }
    $script:TileFilter = $Key
    foreach ($k in $script:Tiles.Keys) {
        $script:Tiles[$k].Panel.BackColor = $(if ($k -eq $Key -and $Key -ne 'All') { $Theme.TileOn } else { [System.Drawing.Color]::White })
    }
    Update-EDashFilter
}

function Show-EDashSnapshot {
    param($Snap)
    $prev = $script:Snapshot
    $script:Snapshot = $Snap
    $script:SystemsByKey = @{}
    foreach ($s in $Snap.Systems) { $script:SystemsByKey[$s.Key] = $s }

    $selKey = $null
    $cur = $script:Grid.CurrentRow
    if ($cur -and $cur.DataBoundItem) { $selKey = [string]$cur.DataBoundItem.Row['Key'] }
    $firstRow = $script:Grid.FirstDisplayedScrollingRowIndex

    # 행 단위 병합: 바뀐 값만 반영해서 깜빡임과 선택·스크롤 초기화를 줄인다.
    $t    = $script:Table
    $keep = @{}
    foreach ($s in $Snap.Systems) {
        $row   = $t.Rows.Find($s.Key)
        $isNew = ($null -eq $row)
        if ($isNew) { $row = $t.NewRow(); $row['Key'] = $s.Key }
        Set-EDashCell $row 'StatusCode' $s.Status
        Set-EDashCell $row 'Rank' ([int]$script:StatusRank[$s.Status])
        Set-EDashCell $row 'Stale' ([bool]$s.Stale)
        Set-EDashCell $row 'Sent' ([bool]$s.AlertSent)
        Set-EDashCell $row 'StatusText' (Get-EDashStatusText $s)
        Set-EDashCell $row 'SRN' $s.SRN
        Set-EDashCell $row 'Site' $s.Site
        Set-EDashCell $row 'System' $s.System
        Set-EDashCell $row 'FSE' $s.FSE
        Set-EDashCell $row 'Region' $s.Region
        foreach ($c in $script:Config.ErrorColumns) { Set-EDashCell $row $c.Name $(if ($s.Errors[$c.Name]) { '●' } else { '' }) }
        Set-EDashCell $row 'ActiveCount' $(if ($null -ne $s.ActiveErrorCount) { [int]$s.ActiveErrorCount } elseif ($s.DashItem) { [int]$s.ActiveCount } else { $null })
        Set-EDashCell $row 'ReminderCount' $(if ($null -ne $s.ReminderCount) { [int]$s.ReminderCount } else { $null })
        Set-EDashCell $row 'LastErr' $s.LastErr
        Set-EDashCell $row 'Age' $(if ($s.Status -ne 'Ok' -and $s.Since) { Format-EDashAge ($Snap.TakenAt - $s.Since) } else { '' })
        Set-EDashCell $row 'AllOkText' $(if ($s.AllOk) { 'Y' } else { '' })
        Set-EDashCell $row 'SentText' $(if ($s.AlertSent) { '발송' } else { '' })
        if ($isNew) { $t.Rows.Add($row) }
        $keep[$s.Key] = $true
    }
    foreach ($r in @($t.Rows)) { if (-not $keep.ContainsKey([string]$r['Key'])) { $r.Delete() } }
    $t.AcceptChanges()

    if ($selKey) { Select-EDashKey $selKey }
    if ($firstRow -ge 0 -and $firstRow -lt $script:Grid.RowCount) { try { $script:Grid.FirstDisplayedScrollingRowIndex = $firstRow } catch { } }

    $sys = @($Snap.Systems)
    $script:SuppressFilter = $true
    Update-EDashCombo $script:CboRegion @($sys | ForEach-Object { $_.Region } | Where-Object { $_ } | Sort-Object -Unique)
    Update-EDashCombo $script:CboFse    @($sys | ForEach-Object { $_.FSE }    | Where-Object { $_ } | Sort-Object -Unique)
    $script:SuppressFilter = $false
    Update-EDashFilter

    $counts = @{
        All      = $sys.Count
        Critical = @($sys | Where-Object { $_.Status -eq 'Critical' }).Count
        Reminder = @($sys | Where-Object { $_.Status -eq 'Reminder' }).Count
        Ok       = @($sys | Where-Object { $_.Status -eq 'Ok' }).Count
        Unknown  = @($sys | Where-Object { $_.Status -eq 'Unknown' }).Count
        Stale    = @($sys | Where-Object { $_.Stale }).Count
        Sent     = @($sys | Where-Object { $_.AlertSent }).Count
    }
    foreach ($k in $counts.Keys) { if ($script:Tiles.ContainsKey($k)) { $script:Tiles[$k].Value.Text = [string]$counts[$k] } }

    $messages = New-Object System.Collections.Generic.List[string]
    if ($Snap.Missing.Count) { $messages.Add('리스트에 없는 컬럼: ' + ($Snap.Missing -join ', ')) }
    $script:LblConn.ForeColor = $Theme.Good
    if ($Snap.Source -eq 'Demo') {
        $script:LblConn.Text    = '● DEMO 데이터'
        $script:LblUpdated.Text = '갱신 {0:HH:mm:ss} KST' -f $Snap.TakenAt
    } elseif ($Snap.Source -eq 'File') {
        $script:LblConn.Text = '● 파일 데이터'
        if ($Snap.GeneratedAt) {
            $age = $Snap.TakenAt - $Snap.GeneratedAt
            $script:LblUpdated.Text = '데이터 {0:HH:mm:ss} KST ({1} 전) · 읽음 {2:HH:mm:ss}' -f $Snap.GeneratedAt, (Format-EDashAge $age), $Snap.TakenAt
            if ($age.TotalMinutes -gt $script:Config.MaxFileAgeMinutes) {
                $script:LblConn.Text      = '● 데이터 오래됨'
                $script:LblConn.ForeColor = $Theme.Bad
                $messages.Add(('데이터가 {0} 전 것입니다. Power Automate 내보내기 Flow의 실행 기록과 OneDrive 동기화 상태를 확인하세요.' -f (Format-EDashAge $age)))
            }
        } else {
            $script:LblUpdated.Text = '읽음 {0:HH:mm:ss} KST (데이터 생성 시각 없음)' -f $Snap.TakenAt
        }
    } else {
        $script:LblConn.Text    = '● 연결됨'
        $script:LblUpdated.Text = '갱신 {0:HH:mm:ss} KST (조회 {1:n1}초)' -f $Snap.TakenAt, ($Snap.DurationMs / 1000)
    }
    $script:StatusMsg.Text = $messages -join '  |  '

    # 알림: 새로 위험이 된 시스템 / 복구된 시스템
    if ($prev -and $script:ChkNotify.Checked) {
        $prevStatus = @{}
        foreach ($p in $prev.Systems) { $prevStatus[$p.Key] = $p.Status }
        $newCrit   = @($sys | Where-Object { $_.Status -eq 'Critical' -and $prevStatus[$_.Key] -ne 'Critical' })
        $recovered = @($sys | Where-Object { $_.Status -eq 'Ok' -and $prevStatus[$_.Key] -in @('Critical', 'Reminder') })
        if ($newCrit.Count) {
            $text = (@($newCrit | Select-Object -First 5 | ForEach-Object { '{0} {1} ({2})' -f $_.SRN, $_.Site, $_.Groups })) -join "`n"
            if ($newCrit.Count -gt 5) { $text += "`n... 외 {0}건" -f ($newCrit.Count - 5) }
            if ($text.Length -gt 250) { $text = $text.Substring(0, 250) }
            $script:Tray.ShowBalloonTip(10000, ('E-Dashboard: 신규 위험 {0}건' -f $newCrit.Count), $text, [System.Windows.Forms.ToolTipIcon]::Warning)
            [System.Media.SystemSounds]::Exclamation.Play()
            Write-EDashLog ('신규 위험: ' + (@($newCrit | ForEach-Object { $_.SRN }) -join ', '))
        } elseif ($recovered.Count -gt 5) {
            $script:Tray.ShowBalloonTip(5000, 'E-Dashboard', ('일괄 복구 {0}건 (18:00 초기화일 수 있음)' -f $recovered.Count), [System.Windows.Forms.ToolTipIcon]::Info)
            Write-EDashLog ('일괄 복구 {0}건' -f $recovered.Count)
        } elseif ($recovered.Count -gt 0) {
            $text = (@($recovered | ForEach-Object { '{0} {1}' -f $_.SRN, $_.Site })) -join "`n"
            $script:Tray.ShowBalloonTip(5000, ('E-Dashboard: 복구 {0}건' -f $recovered.Count), $text, [System.Windows.Forms.ToolTipIcon]::Info)
            Write-EDashLog ('복구: ' + (@($recovered | ForEach-Object { $_.SRN }) -join ', '))
        }
    }
    Update-EDashDetail
}

function Show-EDashError {
    param([string]$Message)
    $script:LblConn.Text = '● 연결 실패'
    $script:LblConn.ForeColor = $Theme.Bad
    $script:StatusMsg.Text = $(if ($Message.Length -gt 180) { $Message.Substring(0, 180) + '...' } else { $Message })
    Write-EDashLog ('ERROR ' + $Message)
    if (-not $script:Snapshot) {
        $script:Detail.Text = (@(
            '데이터를 가져오지 못했습니다.', '',
            "  $Message", '',
            '확인할 것'
        ) + $(if ($script:Config.Source -eq 'File') { @(
            ('  - 데이터 파일 경로: {0}' -f $script:Config.DataFile),
            '  - Power Automate 내보내기 Flow를 만들고 실행했는지 (docs\FLOW-EXPORT.md)',
            '  - OneDrive 동기화가 켜져 있는지 (작업 표시줄의 구름 아이콘)',
            '  - 파일이 다른 곳에 있으면 EDashboard.cmd -DataFile "경로"'
        ) } else { @(
            '  - 사내망 또는 VPN에 연결되어 있는지',
            '  - 브라우저에서 EasDashboard 리스트가 열리는지',
            '  - tools\Test-EDashEnvironment.cmd 로 환경 점검'
        ) }) + @(
            '  - 화면만 먼저 보려면 EDashboard-Demo.cmd 실행', '',
            ('{0}초 후 자동으로 다시 시도합니다. (F5: 바로 재시도)' -f $script:Config.RefreshSeconds)
        ) -join "`r`n")
    }
}

function Start-EDashRefresh {
    if ($script:Job) { return }
    $script:Job = Start-EDashJob $script:Runspace $script:Config
    $script:NextRefreshAt = (Get-Date).AddSeconds($script:Config.RefreshSeconds)
}

function Receive-EDashRefresh {
    $job = $script:Job
    if (-not $job -or -not $job.Handle.IsCompleted) { return }
    $script:Job = $null
    try {
        $snap = Complete-EDashJob $job
        Show-EDashSnapshot $snap
    } catch {
        Show-EDashError $_.Exception.Message
    }
}

function Invoke-EDashManualRefresh {
    $script:RefreshTimer.Stop()
    $script:RefreshTimer.Start()
    Start-EDashRefresh
}

function Update-EDashClock {
    $now = Get-EDashKstNow
    $c = $script:Config
    $isNight = ($now.Hour -ge $c.NightStartHour) -or ($now.Hour -lt $c.NightEndHour)
    $script:LblDayNight.Text = $(if ($isNight) { '야간' } else { '주간' })
    $script:LblDayNight.ForeColor = $(if ($isNight) { New-EDashColor 255 210 120 } else { $Theme.HeaderFore })
    $hm = $c.DailyReset.Split(':')
    $reset = $now.Date.AddHours([int]$hm[0]).AddMinutes([int]$hm[1])
    if ($reset -le $now) { $reset = $reset.AddDays(1) }
    $left = $reset - $now
    $script:LblReset.Text = '{0} 초기화까지 {1}:{2:00}:{3:00}' -f $c.DailyReset, [int][Math]::Floor($left.TotalHours), $left.Minutes, $left.Seconds
    if ($script:Job) { $script:StatusNext.Text = '갱신 중...' }
    else { $script:StatusNext.Text = '다음 갱신 {0}초' -f [int][Math]::Max(0, ($script:NextRefreshAt - (Get-Date)).TotalSeconds) }
}

function Open-EDashItem {
    param([string]$Which)
    $s = Get-EDashSelectedSystem
    if (-not $s) { return }
    if ($script:Config.Demo) {
        [void][System.Windows.Forms.MessageBox]::Show('DEMO 모드에서는 SharePoint 항목을 열 수 없습니다.', 'E-Dashboard')
        return
    }
    $list = $script:Config.DashboardList
    $id   = $s.DashId
    if ($Which -eq 'Alert' -or -not $id) { $list = $script:Config.AlertList; $id = $s.AlertId }
    if (-not $id) { return }
    Start-Process ('{0}/Lists/{1}/DispForm.aspx?ID={2}' -f $script:Config.SiteUrl, $list, $id)
}

function Export-EDashCsv {
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = 'CSV (*.csv)|*.csv'
    $dlg.FileName = 'EDashboard-{0:yyyyMMdd-HHmm}.csv' -f (Get-Date)
    if ($dlg.ShowDialog($script:Form) -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $cols = @($script:Grid.Columns | Where-Object { $_.Visible } | Sort-Object DisplayIndex)
    $rows = foreach ($drv in $script:View) {
        $o = [ordered]@{}
        foreach ($c in $cols) {
            $v = $drv.Row[$c.DataPropertyName]
            if ($v -is [datetime])  { $v = $v.ToString('yyyy-MM-dd HH:mm') }
            elseif ($v -is [DBNull]) { $v = '' }
            $o[$c.HeaderText] = $v
        }
        [pscustomobject]$o
    }
    @($rows) | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8
    $script:StatusMsg.Text = '내보냄: ' + $dlg.FileName
}

# ----- 이벤트 -----
$grid.Add_CellFormatting({
    param($s, $e)
    try {
        if ($e.RowIndex -lt 0) { return }
        $drv = $s.Rows[$e.RowIndex].DataBoundItem
        if ($null -eq $drv) { return }
        $row  = $drv.Row
        $code = [string]$row['StatusCode']
        if ($script:StatusBack.ContainsKey($code)) { $e.CellStyle.BackColor = $script:StatusBack[$code] }
        $name = $s.Columns[$e.ColumnIndex].Name
        if ($script:ErrorColumnNames -contains $name) {
            if ([string]$e.Value) { $e.CellStyle.ForeColor = $Theme.ErrorDot; $e.CellStyle.Font = $Theme.Bold }
        } elseif ($row['Stale'] -eq $true -and $name -in @('StatusText', 'Age', 'LastErr')) {
            $e.CellStyle.ForeColor = $Theme.StaleFore
            $e.CellStyle.Font = $Theme.Italic
        }
    } catch { }
})
$grid.Add_SelectionChanged({ try { Update-EDashDetail } catch { } })
$grid.Add_CellDoubleClick({ param($s, $e) if ($e.RowIndex -ge 0) { Open-EDashItem 'Dash' } })
$grid.Add_CellMouseDown({
    param($s, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Right -and $e.RowIndex -ge 0 -and $e.ColumnIndex -ge 0) {
        $s.CurrentCell = $s.Rows[$e.RowIndex].Cells[$e.ColumnIndex]
    }
})

$menu = New-Object System.Windows.Forms.ContextMenuStrip
$miDash  = $menu.Items.Add('EasDashboard 항목 열기')
$miAlert = $menu.Items.Add('EASDB2 항목 열기')
$miCopy  = $menu.Items.Add('SRN 복사')
$miDash.Add_Click({ Open-EDashItem 'Dash' })
$miAlert.Add_Click({ Open-EDashItem 'Alert' })
$miCopy.Add_Click({ $s = Get-EDashSelectedSystem; if ($s -and $s.SRN) { [System.Windows.Forms.Clipboard]::SetText($s.SRN) } })
$grid.ContextMenuStrip = $menu

$btnRefresh.Add_Click({ Invoke-EDashManualRefresh })
$btnExport.Add_Click({ try { Export-EDashCsv } catch { [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'CSV 내보내기') } })
$btnClear.Add_Click({
    $script:SuppressFilter = $true
    $script:TxtSearch.Text = ''
    $script:CboRegion.SelectedIndex = 0
    $script:CboFse.SelectedIndex = 0
    $script:ChkErrors.Checked = $false
    $script:SuppressFilter = $false
    Set-EDashTileFilter 'All'
})
$script:TxtSearch.Add_TextChanged({ Update-EDashFilter })
$script:CboRegion.Add_SelectedIndexChanged({ Update-EDashFilter })
$script:CboFse.Add_SelectedIndexChanged({ Update-EDashFilter })
$script:ChkErrors.Add_CheckedChanged({ Update-EDashFilter })
$form.Add_KeyDown({ param($s, $e) if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F5) { Invoke-EDashManualRefresh } })

# 트레이 아이콘 (알림 풍선)
$script:Tray = New-Object System.Windows.Forms.NotifyIcon
$script:Tray.Icon = [System.Drawing.SystemIcons]::Information
$script:Tray.Text = 'E-Dashboard'
$script:Tray.Visible = $true
$restore = { if ($script:Form.WindowState -eq 'Minimized') { $script:Form.WindowState = 'Normal' }; $script:Form.Activate() }
$script:Tray.Add_DoubleClick($restore)
$script:Tray.Add_BalloonTipClicked($restore)

# 타이머
$script:PollTimer = New-Object System.Windows.Forms.Timer
$script:PollTimer.Interval = 300
$script:PollTimer.Add_Tick({ Receive-EDashRefresh })
$script:RefreshTimer = New-Object System.Windows.Forms.Timer
$script:RefreshTimer.Interval = $Config.RefreshSeconds * 1000
$script:RefreshTimer.Add_Tick({ Start-EDashRefresh })
$script:ClockTimer = New-Object System.Windows.Forms.Timer
$script:ClockTimer.Interval = 1000
$script:ClockTimer.Add_Tick({ try { Update-EDashClock } catch { } })

$form.Add_Shown({
    $script:Grid.DataSource = $script:View
    foreach ($c in $script:Grid.Columns) {
        if ($script:Headers.Contains($c.Name)) {
            $c.HeaderText = $script:Headers[$c.Name]
            if ($script:Widths.ContainsKey($c.Name)) { $c.Width = $script:Widths[$c.Name] }
        } else {
            $c.Visible = $false
        }
    }
    foreach ($n in $script:ErrorColumnNames) {
        $col = $script:Grid.Columns[$n]
        $col.Width = 72
        $col.DefaultCellStyle.Alignment = 'MiddleCenter'
    }
    foreach ($n in 'ActiveCount', 'ReminderCount', 'AllOkText', 'SentText') { $script:Grid.Columns[$n].DefaultCellStyle.Alignment = 'MiddleCenter' }
    $script:Grid.Columns['LastErr'].DefaultCellStyle.Format = 'MM-dd HH:mm'
    $split.SplitterDistance = [int]($split.Height * 0.66)
    $script:Detail.Text = 'SharePoint에서 데이터를 가져오는 중...'
    Update-EDashClock
    $script:PollTimer.Start()
    $script:ClockTimer.Start()
    $script:RefreshTimer.Start()
    Start-EDashRefresh
})

$form.Add_FormClosed({
    foreach ($t in @($script:PollTimer, $script:RefreshTimer, $script:ClockTimer)) { $t.Stop(); $t.Dispose() }
    $script:Tray.Visible = $false
    $script:Tray.Dispose()
    if ($script:Job) { try { [void]$script:Job.PS.BeginStop($null, $null) } catch { } }
})

try {
    $script:Runspace = New-EDashWorker
    Write-EDashLog ('시작 {0} {1}' -f $Config.SiteUrl, $(if ($Demo) { '[DEMO]' } else { '' }))
    [void]$form.ShowDialog()
} catch {
    Write-EDashLog ('FATAL ' + $_.Exception.Message)
    [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'E-Dashboard 오류')
    exit 1
} finally {
    try { $script:Runspace.Dispose() } catch { }
    try { $script:Mutex.ReleaseMutex() } catch { }
}
