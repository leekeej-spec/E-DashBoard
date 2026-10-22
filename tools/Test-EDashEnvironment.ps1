<#
.SYNOPSIS
    E-Dashboard P0 environment check (read-only).

.DESCRIPTION
    Checks the local PC (PowerShell, .NET, WinForms, execution policy), signs in to the
    SharePoint site with the current Windows account, and inspects the EasDashboard and
    EASDB2 lists: columns and types, value formats, your permission, and SRN quality
    (duplicate rows, quotes/whitespace, matching between the two lists).

    Nothing is written to SharePoint. A report is saved to the Desktop.

    Masking: free-text values appear only as "shapes" (9 = digit, A/a = letter,
    k = Hangul). Actual values are printed only for columns with few distinct values
    (<= MaxDistinctValues) that are not listed in MaskFields and contain no '@'.
    Review the report before sharing it.

    This file is ASCII-only on purpose, so Windows PowerShell 5.1 parses it correctly
    no matter how it was saved or downloaded.

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Test-EDashEnvironment.ps1

.EXAMPLE
    .\Test-EDashEnvironment.ps1 -ListNames EasDashboard -MaxDistinctValues 5
#>
[CmdletBinding()]
param(
    [string]   $SiteUrl           = 'https://share.philips.com/sites/EASalarmhandling',
    [string[]] $ListNames         = @('EasDashboard', 'EASDB2'),
    [string[]] $MaskFields        = @('Title', 'field_2', 'field_3', 'field_4', 'SRN', 'SystemReferenceNumber'),
    [int]      $MaxDistinctValues = 15,
    [string]   $OutFile
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$SiteUrl = $SiteUrl.TrimEnd('/')

# With powershell.exe -File, "-ListNames a,b" arrives as one string; split it here.
$ListNames  = @($ListNames  | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })
$MaskFields = @($MaskFields | ForEach-Object { $_ -split ',' } | ForEach-Object { $_.Trim() } | Where-Object { $_ })

# SRN column per list (used for duplicate and cross-list checks).
$SrnFieldByList = @{ EasDashboard = 'SRN'; EASDB2 = 'SystemReferenceNumber' }

# Internal names the dashboard design relies on.
$ExpectedFields = @{
    EasDashboard = @('SRN', 'field_2', 'field_3', 'field_4', 'field_5', 'field_6', 'field_7', 'field_8', 'field_9',
                     'field_10', 'field_14', 'field_16', 'AllOk', 'StartMonitoring', 'ActiveErrorCount',
                     'ReminderCount', 'LastUpdatedErrorTime')
    EASDB2       = @('SystemReferenceNumber', 'AlertSent')
}

if (-not $OutFile) {
    $folder = [Environment]::GetFolderPath('Desktop')
    if (-not $folder -or -not (Test-Path -LiteralPath $folder)) { $folder = (Get-Location).Path }
    $OutFile = Join-Path $folder ('EDash-test-{0:yyyyMMdd-HHmm}.txt' -f (Get-Date))
}
# .NET file APIs resolve relative paths against the process directory, not the PowerShell location.
$OutFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutFile)

$script:Accept  = 'application/json;odata=nometadata'
$script:Report  = New-Object System.Collections.Generic.List[string]
$script:Results = New-Object System.Collections.Generic.List[object]
$script:Verdict = [ordered]@{ Hosting = 'not checked'; Provider = 'not checked'; OData = 'not checked'; Charts = 'not checked' }

#region Output

$script:StateColor = @{ OK = 'Green'; WARN = 'Yellow'; FAIL = 'Red'; INFO = 'Gray' }

function Out-Line {
    param([string]$Text = '', [string]$Color = 'Gray')
    Write-Host $Text -ForegroundColor $Color
    $script:Report.Add($Text)
}

function Out-Section {
    param([string]$Title)
    Out-Line ''
    Out-Line ('=' * 78) 'Cyan'
    Out-Line $Title 'Cyan'
    Out-Line ('=' * 78) 'Cyan'
}

# INFO lines stay in their section; OK/WARN/FAIL lines are repeated in the summary.
function Out-Check {
    param([string]$Name, [ValidateSet('OK', 'WARN', 'FAIL', 'INFO')][string]$State, [string]$Detail = '')
    $line = '[{0,-4}] {1}' -f $State, $Name
    if ($Detail) { $line += ' : ' + $Detail }
    Out-Line $line $script:StateColor[$State]
    if ($State -ne 'INFO') {
        $script:Results.Add([pscustomobject]@{ Line = $line; Color = $script:StateColor[$State] })
    }
}

#endregion

#region Values

function Limit-Text {
    param([string]$Text, [int]$Max = 40)
    if ($null -eq $Text) { return '' }
    if ($Text.Length -le $Max) { return $Text }
    return $Text.Substring(0, $Max) + ('...(len={0})' -f $Text.Length)
}

function Test-Empty {
    param($Value)
    if ($null -eq $Value) { return $true }
    if ($Value -is [string]) { return [string]::IsNullOrWhiteSpace($Value) }
    if ($Value -is [array]) { return $Value.Count -eq 0 }
    return $false
}

function Get-TypeLabel {
    param($Value)
    if ($Value -is [array]) { return 'Array' }
    return $Value.GetType().Name
}

# Strings are quoted so "true" (text) and true (bool) stay distinguishable.
function Format-Value {
    param($Value)
    if ($null -eq $Value)      { return 'null' }
    if ($Value -is [string])   { return '"' + (Limit-Text $Value 40) + '"' }
    if ($Value -is [bool])     { return ([string]$Value).ToLowerInvariant() + ' (bool)' }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ') + ' (datetime)' }
    if ($Value -is [array])    { return 'Array(' + $Value.Count + ')' }
    if ($Value -is [System.Management.Automation.PSCustomObject]) { return '{object}' }
    return ('{0} ({1})' -f $Value, $Value.GetType().Name)
}

# Hangul jamo + syllables, built from code points to keep this file ASCII-only.
$script:HangulClass = '[{0}-{1}{2}-{3}]' -f [char]0x3131, [char]0x318E, [char]0xAC00, [char]0xD7A3

function Get-Shape {
    param($Value)
    if ($Value -isnot [string]) { return '<' + (Get-TypeLabel $Value) + '>' }
    $s = $Value -replace '\r?\n', '|'
    $s = $s -creplace '[0-9]', '9'
    $s = $s -creplace '[A-Z]', 'A'
    $s = $s -creplace '[a-z]', 'a'
    $s = $s -replace $script:HangulClass, 'k'
    return '"' + (Limit-Text $s 40) + '"'
}

function New-CountMap { New-Object 'System.Collections.Generic.Dictionary[string,int]' }

function Add-Count {
    param($Map, [string]$Key)
    if ($Map.ContainsKey($Key)) { $Map[$Key] = $Map[$Key] + 1 } else { $Map[$Key] = 1 }
}

function Format-Top {
    param($Map, [int]$Count = 5)
    $top   = $Map.GetEnumerator() | Sort-Object -Property Value -Descending | Select-Object -First $Count
    $parts = foreach ($kv in $top) { '{0} x{1}' -f $kv.Key, $kv.Value }
    $text  = $parts -join ' | '
    if ($Map.Count -gt $Count) { $text += (' | ... ({0} distinct)' -f $Map.Count) }
    return $text
}

function Format-Kst {
    param($Value)
    if (Test-Empty $Value) { return '-' }
    try {
        if ($Value -is [datetime]) { $dto = [datetimeoffset]$Value.ToUniversalTime() }
        else { $dto = [datetimeoffset]::Parse([string]$Value, [Globalization.CultureInfo]::InvariantCulture) }
        $tz = [TimeZoneInfo]::FindSystemTimeZoneById('Korea Standard Time')
        return [TimeZoneInfo]::ConvertTime($dto, $tz).ToString('yyyy-MM-dd HH:mm') + ' KST'
    } catch {
        return [string]$Value
    }
}

function Get-SrnKey {
    param([string]$Srn)
    if ($null -eq $Srn) { return '' }
    return ($Srn -replace '["\s]', '').ToUpperInvariant()
}

function Measure-Srn {
    param($Items, [string]$Field)
    $m = [pscustomobject]@{
        Total = 0; Empty = 0; Quote = 0; OuterSpace = 0; InnerSpace = 0; NonAscii = 0
        KeyIds = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.List[string]]'
        KeyRaw = New-Object 'System.Collections.Generic.Dictionary[string,System.Collections.Generic.HashSet[string]]'
    }
    foreach ($it in $Items) {
        $m.Total++
        $raw = $it.$Field
        if (Test-Empty $raw) { $m.Empty++; continue }
        $s = [string]$raw
        if ($s.Contains('"'))          { $m.Quote++ }
        if ($s -ne $s.Trim())          { $m.OuterSpace++ }
        if ($s.Trim() -match '\s')     { $m.InnerSpace++ }
        if ($s -match '[^\x20-\x7E]')  { $m.NonAscii++ }
        $key = Get-SrnKey $s
        if (-not $m.KeyIds.ContainsKey($key)) {
            $m.KeyIds[$key] = New-Object System.Collections.Generic.List[string]
            $m.KeyRaw[$key] = New-Object System.Collections.Generic.HashSet[string]
        }
        $m.KeyIds[$key].Add([string]$it.ID)
        [void]$m.KeyRaw[$key].Add($s)
    }
    return $m
}

#endregion

#region HTTP

function Get-HeaderValue {
    param($Headers, [string]$Name)
    if ($null -eq $Headers) { return $null }
    $value = $null
    try { $value = $Headers[$Name] } catch { }
    if ($null -eq $value) { try { $value = $Headers.GetValues($Name) } catch { } }
    if ($null -eq $value) { return $null }
    return (@($value) -join ', ')
}

# Works for both Windows PowerShell (WebException) and PowerShell 7 (HttpResponseException).
function Get-HttpError {
    param($ErrorRecord)
    $info = [ordered]@{ Status = $null; Auth = $null; Body = $null; Message = $ErrorRecord.Exception.Message }
    $resp = $null
    try { $resp = $ErrorRecord.Exception.Response } catch { }
    if ($resp) {
        try { $info.Status = [int]$resp.StatusCode } catch { }
        $info.Auth = Get-HeaderValue $resp.Headers 'WWW-Authenticate'
    }
    if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
        $raw = $ErrorRecord.ErrorDetails.Message
        $msg = $null
        try {
            $json = $raw | ConvertFrom-Json
            if ($json.'odata.error') { $msg = $json.'odata.error'.message.value }     # odata=nometadata
            elseif ($json.error)     { $msg = $json.error.message.value }             # odata=verbose
        } catch { }
        if (-not $msg) { $msg = ($raw -replace '<[^>]+>', ' ' -replace '\s+', ' ').Trim() }
        $info.Body = Limit-Text $msg 300
    }
    return [pscustomobject]$info
}

function Get-FinalUri {
    param($Response)
    $uri = $null
    try { $uri = $Response.BaseResponse.ResponseUri } catch { }
    if (-not $uri) { try { $uri = $Response.BaseResponse.RequestMessage.RequestUri } catch { } }
    if (-not $uri) { return 'unknown' }
    return ([uri]$uri).GetLeftPart('Path')
}

function Invoke-SpRequest {
    param([string]$Url)
    $request = @{ Uri = $Url; Headers = @{ Accept = $script:Accept }; UseDefaultCredentials = $true; UseBasicParsing = $true; ErrorAction = 'Stop' }
    # PowerShell 7 refuses Windows sign-in over plain http unless told otherwise (intranet sites may be http).
    if ($PSVersionTable.PSVersion.Major -ge 6 -and $Url -like 'http:*') { $request.AllowUnencryptedAuthentication = $true }
    Invoke-WebRequest @request
}

function ConvertFrom-SpResponse {
    param($Response)
    # Decode as UTF-8 ourselves so Korean text survives even without a charset header.
    $text = [System.Text.Encoding]::UTF8.GetString($Response.RawContentStream.ToArray())
    # List items carry both "Id" and "ID"; PowerShell's JSON parser rejects keys that differ only by case.
    $text = $text -creplace '"Id":', '"Id_":'
    $obj = $text | ConvertFrom-Json
    if ($obj.PSObject.Properties['d']) { return $obj.d }
    return $obj
}

function Get-SpEntity {
    param([string]$Url)
    ConvertFrom-SpResponse (Invoke-SpRequest $Url)
}

function Get-SpCollection {
    param([string]$Url)
    $all = New-Object System.Collections.Generic.List[object]
    while ($Url) {
        $page = ConvertFrom-SpResponse (Invoke-SpRequest $Url)
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

function Get-SpVersionHint {
    param([string]$Version)
    if ($Version -notmatch '^(\d+)\.\d+\.\d+\.(\d+)') { return 'unrecognized' }
    $major = [int]$Matches[1]
    $build = [int]$Matches[2]
    if ($major -eq 14) { return 'SharePoint 2010' }
    if ($major -eq 15) { return 'SharePoint 2013' }
    if ($major -ne 16) { return 'unrecognized' }
    if ($build -lt 10000) { return 'SharePoint 2016' }
    if ($build -lt 14000) { return 'SharePoint 2019' }
    if ($build -lt 20000) { return 'SharePoint Subscription Edition' }
    return 'SharePoint Online (by build number)'
}

#endregion

$started = Get-Date
try {
    Out-Line ('E-Dashboard environment check   {0:yyyy-MM-dd HH:mm:ss}' -f $started) 'White'
    Out-Line ('Site: {0}' -f $SiteUrl) 'White'

    #region 1. Local PC
    Out-Section '1. Local PC'

    $psv = $PSVersionTable.PSVersion
    Out-Check 'PowerShell' 'INFO' ('{0} {1}' -f $PSVersionTable.PSEdition, $psv)
    if ($psv.Major -eq 5 -and $psv.Minor -ge 1) {
        Out-Check 'Windows PowerShell 5.1' 'OK'
    } elseif ($psv.Major -ge 6) {
        Out-Check 'Windows PowerShell 5.1' 'WARN' 'this run used PowerShell 7; please run it with powershell.exe (5.1) as well, the dashboard targets 5.1'
    } else {
        Out-Check 'Windows PowerShell 5.1' 'WARN' ('version {0} is older than 5.1' -f $psv)
    }

    $bits = if ([Environment]::Is64BitProcess) { '64-bit process' } else { '32-bit process' }
    Out-Check 'OS' 'INFO' ('{0} ({1})' -f [Environment]::OSVersion.VersionString, $bits)
    Out-Check 'Domain' 'INFO' ([Environment]::UserDomainName)
    Out-Check 'Culture' 'INFO' ('{0} / UI {1}' -f (Get-Culture).Name, (Get-UICulture).Name)
    Out-Check 'Time zone' 'INFO' ([TimeZoneInfo]::Local.Id)
    try {
        [void][TimeZoneInfo]::FindSystemTimeZoneById('Korea Standard Time')
        Out-Check 'Time zone "Korea Standard Time"' 'OK'
    } catch {
        Out-Check 'Time zone "Korea Standard Time"' 'WARN' 'not found'
    }

    try {
        $release = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release -ErrorAction Stop).Release
        if ($release -ge 528040)     { $fx = '4.8 or later' }
        elseif ($release -ge 461808) { $fx = '4.7.2' }
        elseif ($release -ge 394802) { $fx = '4.6.2' }
        else                         { $fx = 'older than 4.6.2' }
        $state = if ($release -ge 461808) { 'OK' } else { 'WARN' }
        Out-Check '.NET Framework' $state ('{0} (release {1})' -f $fx, $release)
    } catch {
        Out-Check '.NET Framework' 'INFO' 'not found (non-Windows)'
    }

    try {
        $policies = @(Get-ExecutionPolicy -List)
        Out-Check 'Execution policy' 'INFO' ((@($policies | ForEach-Object { '{0}={1}' -f $_.Scope, $_.ExecutionPolicy })) -join ', ')
        $gpo = @($policies | Where-Object { "$($_.Scope)" -in 'MachinePolicy', 'UserPolicy' -and "$($_.ExecutionPolicy)" -ne 'Undefined' })
        if ($gpo.Count -eq 0) {
            Out-Check 'Execution policy set by Group Policy' 'OK' 'no'
        } else {
            $set = (@($gpo | ForEach-Object { '{0}={1}' -f $_.Scope, $_.ExecutionPolicy })) -join ', '
            if (@($gpo | Where-Object { "$($_.ExecutionPolicy)" -in 'Restricted', 'AllSigned' }).Count) {
                Out-Check 'Execution policy set by Group Policy' 'WARN' ($set + ' - the dashboard will need code signing or an IT exception')
            } else {
                Out-Check 'Execution policy set by Group Policy' 'OK' ($set + ' - unblock downloaded scripts with Unblock-File')
            }
        }
    } catch {
        Out-Check 'Execution policy' 'INFO' ('could not read: {0}' -f $_.Exception.Message)
    }

    try {
        Add-Type -AssemblyName System.Windows.Forms
        Out-Check 'WinForms' 'OK'
    } catch {
        Out-Check 'WinForms' 'FAIL' 'System.Windows.Forms could not be loaded'
    }
    try {
        Add-Type -AssemblyName System.Windows.Forms.DataVisualization
        Out-Check 'WinForms charting' 'OK'
        $script:Verdict.Charts = 'built-in WinForms charting available'
    } catch {
        Out-Check 'WinForms charting' 'WARN' 'System.Windows.Forms.DataVisualization not available'
        $script:Verdict.Charts = 'built-in charting NOT available'
    }

    # The dashboard reads its data file from the synced OneDrive for Business folder.
    $oneDrive = $env:OneDriveCommercial
    if ($oneDrive -and (Test-Path -LiteralPath $oneDrive)) {
        $dataFile = Join-Path (Join-Path $oneDrive 'EDashboard') 'edash-data.json'
        $fileState = if (Test-Path -LiteralPath $dataFile) { 'data file EDashboard\edash-data.json present' } else { 'data file EDashboard\edash-data.json not created yet' }
        Out-Check 'OneDrive for Business folder' 'OK' ('"{0}" synced; {1}' -f (Split-Path -Leaf $oneDrive), $fileState)
    } else {
        Out-Check 'OneDrive for Business folder' 'WARN' 'not found - sign in to the OneDrive app, or pass -DataFile to the dashboard'
    }

    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
    } catch { }
    try {
        $proxy = [System.Net.WebRequest]::DefaultWebProxy
        if ($proxy) {
            $proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
            $via = $proxy.GetProxy([uri]$SiteUrl)
            if ($proxy.IsBypassed([uri]$SiteUrl) -or -not $via -or -not $via.Host -or $via.Host -eq ([uri]$SiteUrl).Host) {
                Out-Check 'Proxy' 'INFO' 'direct connection'
            } else {
                Out-Check 'Proxy' 'INFO' ('via {0}' -f $via.Authority)
            }
        }
    } catch { }
    #endregion

    #region 2. Server and sign-in
    Out-Section '2. SharePoint server and sign-in'

    $siteUri = [uri]$SiteUrl
    $cnames  = @()
    try {
        if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
            $dns    = @(Resolve-DnsName -Name $siteUri.Host -ErrorAction Stop)
            $cnames = @($dns | Where-Object { "$($_.Type)" -eq 'CNAME' } | ForEach-Object { $_.NameHost })
            $count  = @($dns | Where-Object { "$($_.Type)" -in 'A', 'AAAA' }).Count
        } else {
            $entry = [System.Net.Dns]::GetHostEntry($siteUri.Host)
            if ($entry.HostName -and $entry.HostName -ne $siteUri.Host) { $cnames = @($entry.HostName) }
            $count = $entry.AddressList.Count
        }
        $detail = '{0} address(es)' -f $count
        if ($cnames.Count) { $detail += '; alias ' + ($cnames -join ' -> ') }
        Out-Check ('DNS ' + $siteUri.Host) 'OK' $detail
    } catch {
        Out-Check ('DNS ' + $siteUri.Host) 'FAIL' ('cannot resolve ({0}) - company network / VPN?' -f $_.Exception.Message)
    }

    # Old SharePoint 2013 builds reject odata=nometadata; fall back to odata=verbose.
    $probeUrl   = $SiteUrl + '/_api/web?$select=Title,Url'
    $probe      = $null
    $probeError = $null
    foreach ($accept in @('application/json;odata=nometadata', 'application/json;odata=verbose')) {
        $script:Accept = $accept
        try {
            $probe = Invoke-SpRequest $probeUrl
            $probeError = $null
            break
        } catch {
            $probeError = Get-HttpError $_
            if ($accept -notlike '*nometadata*' -or "$($probeError.Body)" -notmatch 'ACCEPT') { break }
        }
    }

    $authOk    = $false
    $spVersion = $null
    $signIn    = 'REST sign-in with current Windows account'
    if ($probe) {
        $spVersion = Get-HeaderValue $probe.Headers 'MicrosoftSharePointTeamServices'
        $ctype     = Get-HeaderValue $probe.Headers 'Content-Type'
        if ("$ctype" -notmatch 'json') {
            Out-Check $signIn 'FAIL' ('got "{0}" instead of JSON - probably redirected to a sign-in page ({1})' -f $ctype, (Get-FinalUri $probe))
        } else {
            $web    = ConvertFrom-SpResponse $probe
            $authOk = $true
            Out-Check $signIn 'OK' ('site title "{0}"' -f $web.Title)
            Out-Check 'OData format' 'INFO' $script:Accept
            $script:Verdict.OData = $script:Accept
        }
    } elseif ($probeError) {
        $e      = $probeError
        $detail = 'HTTP {0}' -f $e.Status
        if ($e.Auth) { $detail += '; WWW-Authenticate: ' + (Limit-Text $e.Auth 80) }
        $state  = 'FAIL'
        $isSpo  = @($cnames | Where-Object { $_ -match 'sharepoint\.com' }).Count -gt 0
        if ($isSpo -and $e.Status -in 401, 403) {
            # SharePoint Online never accepts Windows-only sign-in for REST; this is the expected answer, not a permission problem.
            $state = 'WARN'
            $hint  = 'expected for SharePoint Online: REST needs a sign-in token, so the dashboard reads a Power Automate export file instead (docs/FLOW-EXPORT.md)'
        } elseif ($e.Status -eq 401 -and "$($e.Auth)" -match 'Bearer') {
            $hint = 'token sign-in required -> looks like SharePoint Online; use the Power Automate export file (docs/FLOW-EXPORT.md)'
        } elseif ($e.Status -eq 401) {
            $hint = 'Windows sign-in rejected; check that the list opens in a browser and the site is in the Local intranet zone'
        } elseif ($e.Status -eq 403) {
            $hint = 'access denied; ask the site owner for at least Read permission'
        } elseif ($e.Status -eq 404) {
            $hint = 'site not found; check -SiteUrl'
        } elseif ($e.Status -eq 407) {
            $hint = 'proxy authentication required'
        } elseif (-not $e.Status) {
            $hint = 'no HTTP response: ' + $e.Message + ' (network, VPN, TLS or proxy)'
        } else {
            $hint = $e.Message
        }
        Out-Check $signIn $state ($detail + ' - ' + $hint)
        if ($e.Body) { Out-Line ('         server said: ' + $e.Body) }
    }

    if ($spVersion) {
        Out-Check 'SharePoint version header' 'INFO' ('{0} -> {1}' -f $spVersion, (Get-SpVersionHint $spVersion))
    }

    if (@($cnames | Where-Object { $_ -match 'sharepoint\.com' }).Count) {
        $script:Verdict.Hosting = 'SharePoint Online (DNS points to sharepoint.com)'
    } elseif ($authOk) {
        $hint = if ($spVersion) { Get-SpVersionHint $spVersion } else { 'version unknown' }
        $script:Verdict.Hosting = 'SharePoint Server on-premises (Windows sign-in works; {0})' -f $hint
    } elseif ($probeError -and $probeError.Status -eq 401 -and "$($probeError.Auth)" -match 'Bearer') {
        $script:Verdict.Hosting = 'SharePoint Online (token sign-in requested)'
    } else {
        $script:Verdict.Hosting = 'unknown - see FAIL lines'
    }
    if ($authOk) {
        $script:Verdict.Provider = 'Rest (Windows integrated sign-in, no extra module)'
    } elseif ($script:Verdict.Hosting -like 'SharePoint Online*') {
        $script:Verdict.Provider = 'File (Power Automate export to OneDrive, see docs/FLOW-EXPORT.md)'
    } else {
        $script:Verdict.Provider = 'undecided - see FAIL lines'
    }
    #endregion

    if ($authOk) {
        #region 3. Lists
        $measured  = [ordered]@{}
        $listIndex = 0
        foreach ($name in $ListNames) {
            $listIndex++
            Out-Section ('3.{0} List {1}' -f $listIndex, $name)

            $listPath = $siteUri.AbsolutePath.TrimEnd('/') + '/Lists/' + $name
            $a1       = "@a1='" + ($listPath -replace ' ', '%20') + "'"
            $base     = $SiteUrl + '/_api/web/GetList(@a1)'

            try {
                $meta = Get-SpEntity ($base + '?' + $a1 + '&$select=Title,ItemCount,LastItemModifiedDate,ListItemEntityTypeFullName')
            } catch {
                $e = Get-HttpError $_
                Out-Check ('{0}: open list' -f $name) 'FAIL' ('{0} -> HTTP {1} {2}' -f $listPath, $e.Status, $e.Body)
                continue
            }
            Out-Check ('{0}: open list' -f $name) 'OK' ('title "{0}", {1} items, last change {2}' -f $meta.Title, $meta.ItemCount, (Format-Kst $meta.LastItemModifiedDate))
            Out-Check ('{0}: item type' -f $name) 'INFO' $meta.ListItemEntityTypeFullName

            try {
                $perm = (Get-SpEntity ($base + '?' + $a1 + '&$select=EffectiveBasePermissions')).EffectiveBasePermissions
                $low  = [uint64]$perm.Low
                if ($low -band 4)     { $access = 'read + edit (write features possible later)' }
                elseif ($low -band 1) { $access = 'read only' }
                else                  { $access = 'no item access' }
                Out-Check ('{0}: your permission' -f $name) 'OK' $access
            } catch {
                Out-Check ('{0}: your permission' -f $name) 'INFO' 'could not be read'
            }

            # Columns
            try {
                $fields = Get-SpCollection ($base + '/fields?' + $a1 + '&$filter=' + [uri]::EscapeDataString('Hidden eq false') +
                                            '&$select=Title,InternalName,TypeAsString,ReadOnlyField,FromBaseType')
            } catch {
                $e = Get-HttpError $_
                Out-Check ('{0}: columns' -f $name) 'FAIL' ('HTTP {0} {1}' -f $e.Status, $e.Body)
                continue
            }
            $custom = @($fields | Where-Object { -not $_.FromBaseType -or $_.InternalName -eq 'Title' })
            Out-Line ''
            Out-Line ('  Columns: {0} custom + Title ({1} visible incl. built-in)' -f ($custom.Count - 1), $fields.Count)
            Out-Line ('  {0,-30} {1,-32} {2,-14} {3}' -f 'InternalName', 'Display name', 'Type', 'ReadOnly')
            foreach ($f in $custom) {
                Out-Line ('  {0,-30} {1,-32} {2,-14} {3}' -f $f.InternalName, $f.Title, $f.TypeAsString, $f.ReadOnlyField)
            }
            Out-Line ''

            $expected = $ExpectedFields[$name]
            if ($expected) {
                $names   = @($fields | ForEach-Object { $_.InternalName })
                $missing = @($expected | Where-Object { $names -notcontains $_ })
                if ($missing.Count) {
                    Out-Check ('{0}: expected columns' -f $name) 'WARN' ('missing: ' + ($missing -join ', '))
                } else {
                    Out-Check ('{0}: expected columns' -f $name) 'OK' ('all {0} present' -f $expected.Count)
                }
            }

            # Items
            $selMap = [ordered]@{}
            foreach ($f in $custom) {
                if ($f.TypeAsString -in @('Computed', 'Attachments')) { continue }
                $sel = $f.InternalName
                if ($sel.StartsWith('_')) { $sel = 'OData_' + $sel }
                if ($f.TypeAsString -in @('User', 'UserMulti', 'Lookup', 'LookupMulti')) { $sel += 'Id' }
                $selMap[$sel] = $f
            }
            $itemsUrl = $base + '/items?' + $a1 + '&$top=500&$select='
            $items = $null
            try {
                $items = Get-SpCollection ($itemsUrl + ((@('ID', 'Created', 'Modified') + @($selMap.Keys)) -join ','))
            } catch {
                $e = Get-HttpError $_
                Out-Check ('{0}: read items' -f $name) 'WARN' ('reading all columns failed (HTTP {0} {1}); retrying with the SRN column only' -f $e.Status, $e.Body)
                $srnOnly = $SrnFieldByList[$name]
                $reduced = [ordered]@{}
                if ($srnOnly -and $selMap.Contains($srnOnly)) { $reduced[$srnOnly] = $selMap[$srnOnly] }
                $selMap = $reduced
                try {
                    $items = Get-SpCollection ($itemsUrl + ((@('ID', 'Created', 'Modified') + @($selMap.Keys)) -join ','))
                } catch {
                    $e = Get-HttpError $_
                    Out-Check ('{0}: read items' -f $name) 'FAIL' ('HTTP {0} {1}' -f $e.Status, $e.Body)
                    continue
                }
            }
            $state = if ($items.Count -eq [int]$meta.ItemCount) { 'OK' } else { 'WARN' }
            Out-Check ('{0}: read items' -f $name) $state ('{0} read, list reports {1}' -f $items.Count, $meta.ItemCount)

            Out-Line ''
            Out-Line '  Value formats  (shape: 9 = digit, A/a = letter, k = Hangul)'
            foreach ($sel in @($selMap.Keys)) {
                $f      = $selMap[$sel]
                $filled = 0
                $hasAt  = $false
                $types  = New-CountMap
                $shapes = New-CountMap
                $values = New-CountMap
                foreach ($it in $items) {
                    $v = $it.$sel
                    if (Test-Empty $v) { continue }
                    $filled++
                    Add-Count $types (Get-TypeLabel $v)
                    Add-Count $shapes (Get-Shape $v)
                    $fv = Format-Value $v
                    if ($fv.Contains('@')) { $hasAt = $true }
                    Add-Count $values $fv
                }
                Out-Line ''
                Out-Line ('  {0}  "{1}"  [{2}]  filled {3}/{4}  types: {5}' -f $f.InternalName, $f.Title, $f.TypeAsString, $filled, $items.Count, (@($types.Keys) -join ', '))
                if ($filled -eq 0) { continue }
                Out-Line ('      shapes : ' + (Format-Top $shapes 5))
                if (($MaskFields -contains $f.InternalName) -or $hasAt) {
                    Out-Line '      values : (hidden: masked column)'
                } elseif ($values.Count -gt $MaxDistinctValues) {
                    Out-Line ('      values : (hidden: {0} distinct values)' -f $values.Count)
                } else {
                    Out-Line ('      values : ' + (Format-Top $values $MaxDistinctValues))
                }
            }

            # SRN quality
            $srnName = $SrnFieldByList[$name]
            if ($srnName -and $selMap.Contains($srnName)) {
                $m = Measure-Srn $items $srnName
                $measured[$name] = $m
                Out-Line ''
                Out-Check ('{0}: SRN rows' -f $name) 'INFO' ('{0} rows, {1} empty, {2} distinct' -f $m.Total, $m.Empty, $m.KeyIds.Count)

                $dups = @($m.KeyIds.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 })
                if ($dups.Count) {
                    $sample = (@($dups | Select-Object -First 10 | ForEach-Object { 'IDs ' + ($_.Value -join ',') })) -join ' | '
                    Out-Check ('{0}: duplicate SRN rows' -f $name) 'WARN' ('{0} SRN(s) on more than one row: {1}' -f $dups.Count, $sample)
                } else {
                    Out-Check ('{0}: duplicate SRN rows' -f $name) 'OK' 'none'
                }

                $issues = @()
                if ($m.Empty)      { $issues += ('{0} empty' -f $m.Empty) }
                if ($m.Quote)      { $issues += ('{0} with double quotes' -f $m.Quote) }
                if ($m.OuterSpace) { $issues += ('{0} with leading/trailing spaces' -f $m.OuterSpace) }
                if ($m.InnerSpace) { $issues += ('{0} with inner spaces/line breaks' -f $m.InnerSpace) }
                if ($m.NonAscii)   { $issues += ('{0} with non-ASCII/control characters' -f $m.NonAscii) }
                if ($issues.Count) {
                    Out-Check ('{0}: SRN formatting' -f $name) 'WARN' ($issues -join ', ')
                } else {
                    Out-Check ('{0}: SRN formatting' -f $name) 'OK' 'clean'
                }
            } elseif ($srnName) {
                Out-Check ('{0}: SRN column' -f $name) 'WARN' ('column "{0}" not found or not readable' -f $srnName)
            }
        }
        #endregion

        #region 4. Cross-list
        if ($measured.Count -ge 2) {
            $keys  = @($measured.Keys)
            $nameA = $keys[0]
            $nameB = $keys[1]
            $A     = $measured[$nameA]
            $B     = $measured[$nameB]
            Out-Section ('4. SRN matching {0} <-> {1}' -f $nameA, $nameB)

            $both = 0; $normOnly = 0; $onlyA = 0; $onlyB = 0
            foreach ($k in $A.KeyRaw.Keys) {
                if ($B.KeyRaw.ContainsKey($k)) {
                    $both++
                    $exact = $false
                    foreach ($raw in $A.KeyRaw[$k]) { if ($B.KeyRaw[$k].Contains($raw)) { $exact = $true; break } }
                    if (-not $exact) { $normOnly++ }
                } else {
                    $onlyA++
                }
            }
            foreach ($k in $B.KeyRaw.Keys) { if (-not $A.KeyRaw.ContainsKey($k)) { $onlyB++ } }

            Out-Check 'SRNs in both lists' 'INFO' ([string]$both)
            if ($normOnly) {
                Out-Check 'SRN spelling between lists' 'WARN' ('{0} SRN(s) match only after removing quotes/spaces/case' -f $normOnly)
            } else {
                Out-Check 'SRN spelling between lists' 'OK' 'matching SRNs are written identically'
            }
            $stateA = if ($onlyA) { 'WARN' } else { 'OK' }
            $stateB = if ($onlyB) { 'WARN' } else { 'OK' }
            Out-Check ('SRNs only in {0}' -f $nameA) $stateA ([string]$onlyA)
            Out-Check ('SRNs only in {0}' -f $nameB) $stateB ([string]$onlyB)
        }
        #endregion
    }
}
catch {
    Out-Check 'Unexpected error' 'FAIL' ('{0} (line {1})' -f $_.Exception.Message, $_.InvocationInfo.ScriptLineNumber)
}
finally {
    Out-Section 'Summary'
    foreach ($k in $script:Verdict.Keys) { Out-Line ('  {0,-9}: {1}' -f $k, $script:Verdict[$k]) 'White' }
    Out-Line ''
    foreach ($r in $script:Results) { Out-Line $r.Line $r.Color }
    Out-Line ''
    Out-Line ('Elapsed {0:n1}s' -f ((Get-Date) - $started).TotalSeconds)
    try {
        [System.IO.File]::WriteAllLines($OutFile, $script:Report.ToArray(), (New-Object System.Text.UTF8Encoding($true)))
        Write-Host ''
        Write-Host ('Report saved: {0}' -f $OutFile) -ForegroundColor Green
        Write-Host 'Please review it before sharing.' -ForegroundColor Green
    } catch {
        Write-Host ('Could not save the report: {0}' -f $_.Exception.Message) -ForegroundColor Red
    }
}
