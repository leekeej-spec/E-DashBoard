<#
.SYNOPSIS
    Runs the Power Fx formulas from docs/POWERAPPS.md against mock SharePoint data.

.DESCRIPTION
    Every code block in the guide preceded by <!-- fx:Name --> is extracted and evaluated with the
    open-source Power Fx interpreter (downloaded from nuget.org into .lib on first run), so the
    formulas users paste into Power Apps are checked for syntax and for the expected statuses.
    Canvas-only functions are adapted: Refresh is skipped, Notify prints its message, and
    ClearCollect/Set store results as variables.

    Needs PowerShell 7 (the interpreter targets .NET Standard 2.0).

.EXAMPLE
    pwsh tests/powerfx/Test-PowerAppsFormulas.ps1
    pwsh tests/powerfx/Test-PowerAppsFormulas.ps1 -TextDates      # section 12 variant: text date columns
    pwsh tests/powerfx/Test-PowerAppsFormulas.ps1 -ChoiceRegion   # section 12 variant: REGION is a Choice column
#>
param(
    [string] $Doc = (Join-Path $PSScriptRoot '../../docs/POWERAPPS.md'),
    [switch] $TextDates,
    [switch] $ChoiceRegion
)
$ErrorActionPreference = 'Stop'

$lib = Join-Path $PSScriptRoot '.lib'
if (-not (Test-Path (Join-Path $lib 'Microsoft.PowerFx.Interpreter.dll'))) {
    New-Item -ItemType Directory -Force -Path (Join-Path $lib 'en-US') | Out-Null
    foreach ($pkg in 'microsoft.powerfx.core', 'microsoft.powerfx.interpreter', 'microsoft.powerfx.transport.attributes') {
        $zip = Join-Path $lib "$pkg.zip"
        Invoke-WebRequest "https://api.nuget.org/v3-flatcontainer/$pkg/1.4.0/$pkg.1.4.0.nupkg" -OutFile $zip
        $dir = Join-Path $lib $pkg
        Expand-Archive $zip -DestinationPath $dir -Force
        Copy-Item (Join-Path $dir 'lib/netstandard2.0/*.dll') $lib
        $res = Join-Path $dir 'lib/netstandard2.0/en-US'
        if (Test-Path $res) { Copy-Item (Join-Path $res '*.dll') (Join-Path $lib 'en-US') }
    }
}
Get-ChildItem $lib -Filter *.dll -File | ForEach-Object { Add-Type -Path $_.FullName }
$config = [Microsoft.PowerFx.PowerFxConfig]::new([Microsoft.PowerFx.Features]::PowerFxV1)
[Microsoft.PowerFx.PowerFxConfigExtensions]::EnableSetFunction($config)
$engine = [Microsoft.PowerFx.RecalcEngine]::new($config)
$opts = [Microsoft.PowerFx.ParserOptions]::new(); $opts.AllowsSideEffects = $true

function Eval([string]$expr, [string]$label) {
    try { return $engine.Eval($expr, $null, $opts) }
    catch {
        $e = $_.Exception; while ($e.InnerException) { $e = $e.InnerException }
        throw ("[{0}] {1}" -f $label, $e.Message)
    }
}
function Show($v) {
    if ($null -eq $v) { return '<null>' }
    if ($v -is [Microsoft.PowerFx.Types.BlankValue]) { return '<blank>' }
    if ($v -is [Microsoft.PowerFx.Types.ErrorValue]) { return 'ERROR: ' + (($v.Errors | % Message) -join '; ') }
    if ($v.PSObject.Properties['Value']) { return [string]$v.Value }
    return $v.ToString()
}

# --- formulas from the guide
$md = Get-Content -Raw $Doc
$fx = [ordered]@{}
foreach ($m in [regex]::Matches($md, '<!-- fx:(\S+) -->\s*```powerfx\r?\n(.*?)```', 'Singleline')) { $fx[$m.Groups[1].Value] = $m.Groups[2].Value }
"formulas found: " + ($fx.Keys -join ', ')
# apply the guide's section-12 variants exactly as written in its table
$variants = @()
if ($TextDates) {
    $variants += , @('AllOkT, AllOkTime,', 'AllOkT, DateTimeValue(AllOkTime),')
    $variants += , @('ErrT, Coalesce(LastUpdatedErrorTime, LastUpdated),', 'ErrT, DateTimeValue(Coalesce(LastUpdatedErrorTime, LastUpdated)),')
    $variants += , @('StartT, StartMonitoring', 'StartT, DateTimeValue(StartMonitoring)')
}
if ($ChoiceRegion) { $variants += , @('RegionTxt, REGION,', 'RegionTxt, REGION.Value,') }
foreach ($v in $variants) {
    $row = '| `' + $v[0] + '` | `' + $v[1] + '` |'
    if (-not $md.Contains($row)) { throw "variant not in guide table: $($v[0])" }
    if (-not $fx['btnLoad.OnSelect'].Contains($v[0])) { throw "variant source line not in formula: $($v[0])" }
    $fx['btnLoad.OnSelect'] = $fx['btnLoad.OnSelect'].Replace($v[0], $v[1])
    "variant applied: " + $v[1]
}

# split a behavior formula into top-level statements (outside strings, parens, comments)
function Split-Statements([string]$text) {
    $out = New-Object System.Collections.Generic.List[string]; $sb = New-Object System.Text.StringBuilder
    $depth = 0; $inStr = $false; $i = 0
    while ($i -lt $text.Length) {
        $c = $text[$i]
        if ($inStr) { [void]$sb.Append($c); if ($c -eq '"') { if ($i + 1 -lt $text.Length -and $text[$i+1] -eq '"') { [void]$sb.Append('"'); $i++ } else { $inStr = $false } }; $i++; continue }
        if ($c -eq '/' -and $i + 1 -lt $text.Length -and $text[$i+1] -eq '/') { while ($i -lt $text.Length -and $text[$i] -ne "`n") { $i++ }; continue }
        if ($c -eq '"') { $inStr = $true }
        elseif ('([{'.Contains($c)) { $depth++ } elseif (')]}'.Contains($c)) { $depth-- }
        if ($c -eq ';' -and $depth -eq 0) { $out.Add($sb.ToString().Trim()); [void]$sb.Clear() } else { [void]$sb.Append($c) }
        $i++
    }
    if ($sb.ToString().Trim()) { $out.Add($sb.ToString().Trim()) }
    return $out
}
function Split-FirstArg([string]$inner) {   # "name, expr" -> name, expr
    $idx = $inner.IndexOf(','); return @($inner.Substring(0, $idx).Trim(), $inner.Substring($idx + 1).Trim())
}
function Run-Behavior([string]$name) {
    foreach ($st in (Split-Statements $fx[$name])) {
        if ($st -match '^Refresh\(') { continue }
        if ($st -match '^(ClearCollect|Set)\((?s)(.*)\)$') {
            $parts = Split-FirstArg $Matches[2]
            if ($parts[1] -eq 'Blank()') { continue }   # interpreter can't retype a Blank variable; Power Apps infers it
            if ($parts[0] -eq 'colPrevCrit' -and -not $script:HaveSys) {
                $engine.UpdateVariable('colPrevCrit', (Eval 'Filter(Table({SrnKey: ""}), false)' 'init'))
                continue
            }
            $v = Eval $parts[1] ("{0} -> {1}" -f $name, $parts[0])
            if ($v -is [Microsoft.PowerFx.Types.ErrorValue]) { throw ("{0}: {1}" -f $parts[0], (Show $v)) }
            $engine.UpdateVariable($parts[0], $v)
            if ($parts[0] -eq 'colSys') { $script:HaveSys = $true }
            continue
        }
        # other statements: Notify(...) becomes the message text so we can print it
        $expr = [regex]::Replace($st, 'Notify\((.+?), NotificationType\.\w+(?:, \d+)?\)', '$1')
        $v = Eval $expr ("{0} stmt" -f $name)
        if (-not ($v -is [Microsoft.PowerFx.Types.BlankValue])) { "   NOTIFY: " + (Show $v) }
    }
}

# --- mock SharePoint data (display names as Power Apps shows them); gNow = 2026-10-22 20:00
$D = { param($h, $m) if ($TextDates) { '"2026-10-22 {0:00}:{1:00}"' -f $h, $m } else { 'DateTime(2026,10,22,{0},{1},0)' -f $h, $m } }
$noDate = if ($TextDates) { 'If(false, "")' } else { 'If(false, Now())' }
$T = 'If(false, "")'
$regionLit = if ($ChoiceRegion) { '{Value: "서울"}' } else { '"서울"' }
function Row($id, $srn, $site, $err, $allOk, $allOkT, $errT, $startT, $act, $rem) {
    $e = @{ E7 = $T; E9 = $T; E11 = $T; E13 = $T; E20 = $T }; foreach ($k in $err.Keys) { $e[$k] = $err[$k] }
    "{ID: $id, SRN: $srn, SiteName: ""$site"", SystemName: ""Ingenia 1.5T"", FSE: ""FSE-A"", REGION: $regionLit, " +
    "Err7_SupplyWaterHigh: $($e.E7), Err9_ReturnWaterHigh: $($e.E9), Err11_DeltaTempHigh: $($e.E11), Err13_ColdheadGasTempHigh: $($e.E13), Err20_Reminder: $($e.E20), " +
    "AllOk: $allOk, AllOkTime: $allOkT, LastUpdated: $errT, LastUpdatedErrorTime: $errT, StartMonitoring: $startT, ActiveErrorCount: $act, ReminderCount: $rem}"
}
$rows = @(
    (Row 1 '"710001"' '가람병원' @{ E7 = '"2026-10-22 19:35"' } 'false' $noDate (& $D 19 35) (& $D 19 35) 1 0)
    (Row 2 '"""710002"""' '나래병원' @{ E20 = '"2026-10-22 17:00"' } 'false' $noDate (& $D 17 0) (& $D 15 0) 0 2)
    (Row 3 '"710003"' '다온의원' @{} 'true' (& $D 18 0) $noDate $noDate 0 0)
    (Row 4 '"710004"' '라온병원' @{ E13 = '"Yes"' } 'false' $noDate (& $D 12 0) (& $D 12 0) 1 0)
    (Row 5 '"710001"' '가람병원' @{ E9 = '"2026-10-22 19:20"' } 'false' $noDate (& $D 19 20) (& $D 19 20) 1 0)
    (Row 6 '"710005"' '마루병원' @{ E7 = '"2026-10-22 10:00"' } 'true' (& $D 12 0) (& $D 10 0) (& $D 10 0) 1 0)
    (Row 7 '"710006"' '바다병원' @{ E11 = '"2026-10-22 19:50"' } 'true' (& $D 9 0) (& $D 19 50) (& $D 19 50) 1 0)
    (Row 8 '"710007"' '새봄병원' @{ E7 = '"No"' } 'false' $noDate $noDate $noDate 0 0)
)
$engine.UpdateVariable('EasDashboard', (Eval ('Table(' + ($rows -join ",`n") + ')') 'mock EasDashboard'))
$engine.UpdateVariable('EASDB2', (Eval 'Table({ID:1, SystemReferenceNumber:"710001", AlertSent:"true"},{ID:2, SystemReferenceNumber:"710002", AlertSent:"false"},{ID:3, SystemReferenceNumber:"710003", AlertSent:"false"},{ID:4, SystemReferenceNumber:"710004", AlertSent:"true"},{ID:5, SystemReferenceNumber:"719999", AlertSent:"false"},{ID:6, SystemReferenceNumber:" 710006 ", AlertSent:"TRUE"})' 'mock EASDB2'))
$engine.UpdateVariable('txtSearch', (Eval '{Text: ""}' 'ctl'))
$engine.UpdateVariable('chkErrors', (Eval '{Value: false}' 'ctl'))
foreach ($cb in 'cbRegion', 'cbFse') { $engine.UpdateVariable($cb, (Eval '{SelectedItems: Filter(Table({Value: ""}), false), Selected: {Value: If(false, "")}}' 'ctl')) }

# --- App.OnStart, then pin the clock for deterministic results
Run-Behavior 'App.OnStart'
$engine.UpdateVariable('gSelId', 1)
$fx['btnLoad.OnSelect'] = $fx['btnLoad.OnSelect'].Replace('Set(gNow, Now())', 'Set(gNow, DateTime(2026,10,22,20,0,0))')

"--- btnLoad (1st run)"
Run-Behavior 'btnLoad.OnSelect'
"--- colSys"
$sys = Eval 'colSys' 'read'
foreach ($r in $sys.Rows) {
    $f = $r.Value
    '  ID={0} SRN={1,-7} Status={2,-8} Stale={3,-5} Dup={4,-5} Sent={5,-5} HasAlert={6,-5} Groups={7,-14} ErrT={8} Since={9} Sort={10}' -f `
        (Show $f.GetField('ID')), (Show $f.GetField('SrnText')), (Show $f.GetField('Status')), (Show $f.GetField('Stale')), (Show $f.GetField('Dup')),
        (Show $f.GetField('Sent')), (Show $f.GetField('HasAlert')), (Show $f.GetField('Groups')), (Show $f.GetField('ErrT')), (Show $f.GetField('Since')), (Show $f.GetField('SortKey'))
}
"  unmatched: " + ((Eval 'Concat(colUnmatched, SystemReferenceNumber, "|")' 'u').Value)

"--- galKpi.Items"
"  " + ((Eval ('Concat(' + $fx['galKpi.Items'] + ', L & "=" & C, ", ")') 'kpi').Value)

"--- galSys.Items under filters"
$cases = @(
    @{ f = 'All' }, @{ f = 'Critical' }, @{ f = 'Reminder' }, @{ f = 'Stale' }, @{ f = 'Ok' }, @{ f = 'Sent' },
    @{ f = 'All'; q = '가람' }, @{ f = 'All'; q = '710' ; err = $true }, @{ f = 'All'; region = '서울' }
)
foreach ($c in $cases) {
    $engine.UpdateVariable('gFilter', [string]$c.f)
    $engine.UpdateVariable('txtSearch', (Eval ('{Text: "' + [string]$c.q + '"}') 'ctl'))
    $engine.UpdateVariable('chkErrors', (Eval ('{Value: ' + $(if ($c.err) { 'true' } else { 'false' }) + '}') 'ctl'))
    if ($c.region) { $engine.UpdateVariable('cbRegion', (Eval ('{SelectedItems: Table({Value: "' + $c.region + '"}), Selected: {Value: "' + $c.region + '"}}') 'ctl')) }
    else { $engine.UpdateVariable('cbRegion', (Eval '{SelectedItems: Filter(Table({Value: ""}), false), Selected: {Value: If(false, "")}}' 'ctl')) }
    $v = Eval ('Concat(' + $fx['galSys.Items'] + ', SrnText & "(" & Status & ")", ", ")') 'galSys'
    '  filter={0,-8} q={1,-4} errOnly={2,-5} region={3,-3} -> {4}' -f $c.f, $c.q, [bool]$c.err, $c.region, $v.Value
}
$engine.UpdateVariable('gFilter', 'All'); $engine.UpdateVariable('txtSearch', (Eval '{Text: ""}' 'ctl')); $engine.UpdateVariable('chkErrors', (Eval '{Value: false}' 'ctl'))

"--- cbRegion.Items / cbFse.Items"
"  " + (Eval 'Concat(Distinct(Filter(colSys, !IsBlank(RegionTxt)), RegionTxt), Value, ",")' 'cb').Value

"--- gallery templates per row"
$tpl = 'galSys.TemplateFill', 'rectBar.Fill', 'lblStatus.Text', 'lblStatus.Color', 'lblE7.Fill', 'lblE7.Color', 'lblLast.Text', 'lblAge.Text'
foreach ($r in $sys.Rows) {
    $id = Show $r.Value.GetField('ID')
    $parts = foreach ($t in $tpl) {
        $expr = 'With({ti: LookUp(colSys, ID = ' + $id + ')}, ' + ($fx[$t] -replace 'ThisItem\.', 'ti.') + ')'
        $v = Eval $expr $t
        if ($t -like '*.Text') { '{0}="{1}"' -f $t.Split('.')[0], $v.Value } else { $null = $v }
    }
    '  ID {0}: {1}' -f $id, ($parts -join '  ')
}
foreach ($r in (Eval 'colUnmatched' 'u').Rows) {
    $expr = 'With({ti: First(colUnmatched)}, ' + ($fx['lblUnmatched.Text'] -replace 'ThisItem\.', 'ti.') + ')'
    '  unmatched label: ' + (Eval $expr 'lblUnmatched').Value
}

"--- htmlDetail"
foreach ($id in 4, 2) {
    $engine.UpdateVariable('gSelId', $id)
    $h = (Eval $fx['htmlDetail.HtmlText'] 'html').Value
    '  ID {0}: {1}' -f $id, (($h -replace '<br>', ' | ' -replace '<[^>]+>', '').Trim())
}
$engine.UpdateVariable('gSelId', 999)
'  ID 999: ' + ((Eval $fx['htmlDetail.HtmlText'] 'html').Value -replace '<[^>]+>', '')

"--- lblHeader.Text"
foreach ($t in 'DateTime(2026,10,22,20,0,0)', 'DateTime(2026,10,22,7,30,0)', 'DateTime(2026,10,22,17,59,0)') {
    $engine.UpdateVariable('gNow', (Eval $t 't')); $engine.UpdateVariable('gLoadedAt', (Eval $t 't'))
    '  {0} -> {1}' -f $t, (Eval $fx['lblHeader.Text'] 'header').Value
}

"--- btnLoad (2nd run: 710003 turns critical, 710001 rows recover)"
$upd = 'ForAll(EasDashboard As r, If(r.ID = 3, Patch(r, {Err7_SupplyWaterHigh: "2026-10-22 19:59", AllOk: false, LastUpdatedErrorTime: ' + (& $D 19 59) + '}), r.ID = 1 || r.ID = 5, Patch(r, {Err7_SupplyWaterHigh: "", Err9_ReturnWaterHigh: "", AllOk: true, AllOkTime: ' + (& $D 19 58) + '}), r))'
$engine.UpdateVariable('EasDashboard', (Eval $upd 'mutate'))
Run-Behavior 'btnLoad.OnSelect'
"  statuses: " + (Eval 'Concat(colSys, SrnText & "=" & Status, ", ")' 's').Value
