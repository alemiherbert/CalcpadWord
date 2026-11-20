<#
.SYNOPSIS
  End-to-end test of the installed add-in, in a separate hidden Word instance.
.DESCRIPTION
  Starts its own Word (it never touches a Word you have open), builds a test document and
  drives the add-in: Evaluate, {= } pick-up, switches, errors and error navigation, Edit Source
  and evaluate-on-leaving, headers/footers/text boxes, settings, copy and paste, save and
  reopen, undo, and a clean Word shutdown. Writes e2e.log, e2e.docx and e2e-*.png (pictures of
  the rendered page) to -OutDir. If Word hangs, only the Word instance this script started is
  killed after -TimeoutSeconds. Run tools\Register.ps1 first.
  Not covered (needs a visible, focused Word): the ribbon UI, right-click menus, keyboard
  shortcuts, double-click. The script does check the add-in's answers behind those.
#>
param(
    [string]$OutDir = (Join-Path (Split-Path $PSScriptRoot -Parent) 'out\e2e'),
    [int]$TimeoutSeconds = 180,
    [switch]$Worker
)

if (-not $Worker) {
    New-Item -ItemType Directory -Force $OutDir | Out-Null
    $pidFile = Join-Path $OutDir 'word.pid'
    $inUse = Join-Path $OutDir 'word.in-use'
    $log = Join-Path $OutDir 'e2e.log'
    Remove-Item $pidFile, $inUse, $log -ErrorAction SilentlyContinue
    $child = Start-Process powershell.exe -PassThru -WindowStyle Hidden -ArgumentList @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"", '-OutDir', "`"$OutDir`"", '-Worker')
    $finished = $child.WaitForExit($TimeoutSeconds * 1000)
    if (-not $finished) { Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue }
    # Never leave the test's hidden Word behind (it is identified by PID, so a Word you opened is
    # safe), unless your own window ended up in it.
    if (Test-Path $inUse) {
        Write-Output "The test's Word (PID $(Get-Content $pidFile)) has your documents open: $(Get-Content $inUse). It was left running."
    }
    elseif (Test-Path $pidFile) {
        $leftover = Get-Process -Id ([int](Get-Content $pidFile)) -ErrorAction SilentlyContinue
        if ($leftover -and $leftover.MainWindowHandle -ne [IntPtr]::Zero) {
            Write-Output "The test's Word (PID $($leftover.Id)) has a visible window, perhaps yours: it was left running."
        }
        elseif ($leftover) { $leftover | Stop-Process -Force; Write-Output "Cleaned up the test's Word instance (PID $($leftover.Id))." }
    }
    if (Test-Path $log) { Get-Content $log -Encoding UTF8 }
    if (-not $finished) { Write-Output "TIMEOUT after $TimeoutSeconds s." ; exit 1 }
    exit $child.ExitCode
}

# ------------------------------------------------------------------------------------------
# Worker
# ------------------------------------------------------------------------------------------
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'e2e.log'
$script:failures = 0
function Say([string]$m) { Add-Content -Path $log -Value $m -Encoding UTF8 }
function Check([bool]$ok, [string]$what) {
    if ($ok) { Say "  PASS  $what" } else { Say "  FAIL  $what"; $script:failures++ }
}
function Show([string]$s) { $s -replace "`r", '\r' -replace "`a", '\a' -replace "`v", '\v' }

# Late-bound calls through IDispatch: the add-in has no type library (as with any .NET Core COM
# server), and Word methods with by-ref arguments (Quit, Close, SaveAs2, Open, Undo) need it too.
# Arguments must be plain values: COM rejects PowerShell's PSObject wrappers.
function Call($obj, [string]$name, [object[]]$arguments = @()) {
    [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::InvokeMethod, $null, $obj, $arguments)
}
function Prop($obj, [string]$name) {
    [System.__ComObject].InvokeMember($name, [Reflection.BindingFlags]::GetProperty, $null, $obj, $null)
}
# Word runs the add-in's deferred work (evaluate after leaving a control) while it is idle.
function WaitFor([scriptblock]$condition, [int]$milliseconds = 4000) {
    $deadline = [DateTime]::Now.AddMilliseconds($milliseconds)
    while ([DateTime]::Now -lt $deadline) {
        # While the add-in redraws, a control can be briefly unreadable: that is "not yet", not a failure.
        try { if (& $condition) { return $true } } catch { }
        Start-Sleep -Milliseconds 100
    }
    return [bool](& $condition)
}
# A command that recalculates may finish in the background: wait for that, then return how it ended.
function Command([string]$name, [object[]]$arguments = @()) {
    $result = Call $api $name $arguments
    if ($result -match 'calculating in the background') {
        if (-not (WaitFor { -not (Prop $api 'IsWorking') } 120000)) { throw "$name did not finish in the background" }
        $result = Prop $api 'LastStatus'
    }
    return $result
}

# Word throws for an idMso it doesn't know; an unknown context menu id could break the ribbon.
function Test-MsoId([string]$id) {
    try { [void]$word.CommandBars.GetLabelMso($id); $true } catch { $false }
}
# A hidden Word can't draw icons (E_UNEXPECTED, even for "Bold"), but an unknown icon name fails
# differently, with "Value does not fall within the expected range" (E_INVALIDARG).
function Test-ImageMso([string]$id) {
    try { [void]$word.CommandBars.GetImageMso($id, 32, 32); $true }
    catch { -not ($_.Exception.Message -match 'expected range') }
}
function BlockControl($doc, [string]$state = '*') {
    for ($i = 1; $i -le $doc.ContentControls.Count; $i++) {
        $cc = $doc.ContentControls.Item($i)
        if ($cc.Tag -like "calcpad-block:$state`:*") { return $cc }
    }
    return $null
}
function CursorIn($range) { $r = $range.Duplicate; $r.Collapse(0); $r.Select() }   # end of a control's content is inside it

# Where to save a test document: its usual name, or a fresh one if you have that file open in Word.
function SavePath([string]$name) {
    $path = Join-Path $OutDir "$name.docx"
    try { if (Test-Path $path) { ([IO.File]::Open($path, 'Open', 'ReadWrite', 'None')).Dispose() }; return [string]$path }
    catch { return [string](Join-Path $OutDir ("{0}-{1:HHmmss}.docx" -f $name, (Get-Date))) }
}

function SourceOf($cc) { $cc.Tag.Split(':', 4)[3] }
function StateOf($cc) { $cc.Tag.Split(':', 4)[1] }
function Control($range, [string]$sourceStart) {
    $controls = $range.ContentControls
    for ($i = 1; $i -le $controls.Count; $i++) {
        $cc = $controls.Item($i)
        if ($cc.Tag -like 'calcpad:*' -and (SourceOf $cc).StartsWith($sourceStart)) { return $cc }
    }
    throw "No Calcpad control with source starting '$sourceStart'"
}
function Dump($doc) {
    $controls = $doc.Content.ContentControls
    for ($i = 1; $i -le $controls.Count; $i++) {
        $cc = $controls.Item($i)
        Say ("    [{0,2}] {1} maths={2} locked={3,-5} src=[{4}] text=[{5}] title=[{6}]" -f
            $i, (StateOf $cc), $cc.Range.OMaths.Count, $cc.LockContents, (SourceOf $cc), (Show $cc.Range.Text), $cc.Title)
    }
}
function Picture($range, [string]$name) {
    try {
        Add-Type -AssemblyName System.Drawing
        $emf = Join-Path $OutDir "$name.emf"
        [IO.File]::WriteAllBytes($emf, [byte[]]$range.EnhMetaFileBits)
        $mf = New-Object System.Drawing.Imaging.Metafile($emf)
        $scale = 1.5
        $bmp = New-Object System.Drawing.Bitmap([int]($mf.Width * $scale), [int]($mf.Height * $scale))
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White)
        $g.DrawImage($mf, 0, 0, $bmp.Width, $bmp.Height)
        $png = Join-Path $OutDir "$name.png"
        $bmp.Save($png, [System.Drawing.Imaging.ImageFormat]::Png)
        $g.Dispose(); $bmp.Dispose(); $mf.Dispose(); Remove-Item $emf
        Say "  picture: $png"
    }
    catch { Say "  (could not save picture $name : $($_.Exception.Message))" }
}

$sigma = [char]0x03C3
$lines = @(
    'Simply supported beam'
    'L = 6m'
    'q = 12kN/m'
    'M = q*L^2/8'
    'The span is {=#val L} and the midspan moment is {=#val M|kNm}.'
    'Reaction: {= R = q*L/2}'
    '{=#hide k = 2}'
    "Stress: {=#nosub $sigma = M/(0.2m*(0.3m)^2/6)|MPa}"
    'Root: {= x_r = $Root{x^2 - 2 @ x = 0 : 2}}'
    'Broken: {= y = nothing*2}'
    'Hidden k is used here: {=#val k*3}'
)

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Say 'Could not start a separate Word instance; nothing was changed.'; exit 2 }
Set-Content (Join-Path $OutDir 'word.pid') $new[0]

try {
    $word.Visible = $false
    $word.DisplayAlerts = 0

    Say 'Loading the add-in'
    $addin = $word.COMAddIns.Item('CalcpadWord.Connect')
    if (-not $addin.Connect) { $addin.Connect = $true }
    $api = $addin.Object
    Check ($null -ne $api) 'add-in is connected and exposes its API'
    Say "  version $(Prop $api 'Version')"
    $ribbonXml = [xml](Call $api 'GetCustomUI' @('Microsoft.Word.Document'))
    Check ($ribbonXml.customUI.ribbon.tabs.tab.label -eq 'Calcpad') 'ribbon XML is well-formed and has the Calcpad tab'
    $icons = @($ribbonXml.SelectNodes('//*[@imageMso]') | ForEach-Object { $_.imageMso } | Select-Object -Unique)
    $badIcons = @($icons | Where-Object { -not (Test-ImageMso $_) })
    Check ($badIcons.Count -eq 0) "every ribbon icon exists in this Word $(if ($badIcons) { '(unknown: ' + ($badIcons -join ', ') + ')' })"
    $menus = @($ribbonXml.customUI.contextMenus.contextMenu)
    $equationMenu = $menus | Where-Object { $_.idMso -eq 'ContextMenuEquation' }
    $ids = @($menus | ForEach-Object { $_.button } | ForEach-Object { $_.id })
    Check ($null -ne $equationMenu -and @($equationMenu.button).Count -eq 3 -and ($ids | Select-Object -Unique).Count -eq $ids.Count) 'right-click menus: equation menu has Edit Source, Calculate and Block to Text, all ids unique'
    $badMenus = @($menus | Where-Object { -not (Test-MsoId $_.idMso) } | ForEach-Object { $_.idMso })
    Check ($badMenus.Count -eq 0) "every context menu id exists in this Word $(if ($badMenus) { '(unknown: ' + ($badMenus -join ', ') + ')' })"
    $noControl = [System.Runtime.InteropServices.DispatchWrapper]::new([object]$null)   # Office passes an IRibbonControl here
    Check ((Call $api 'GetDecimalsIndex' @($noControl)) -eq 2) 'ribbon callbacks are callable through IDispatch'
    Check (WaitFor { -not (Prop $api 'IsWorking') } 30000) 'Calcpad warms up in the background at start-up'

    $doc = $word.Documents.Add()
    $doc.Content.Text = $lines -join "`r"
    $paragraphs = $doc.Paragraphs.Count

    Say 'Step 1: select the three formula lines, press Evaluate'
    $doc.Range($doc.Paragraphs.Item(2).Range.Start, $doc.Paragraphs.Item(4).Range.End).Select()
    Say ('  -> ' + (Command 'Evaluate'))
    Dump $doc
    Check ($doc.ContentControls.Count -eq 11) '3 selected lines + 8 {= } expressions became controls'
    Check ($doc.Paragraphs.Count -eq $paragraphs) 'no paragraphs were added or removed'
    $m = Control $doc.Content 'M = '
    Check ($m.Range.OMaths.Count -eq 1 -and $m.LockContents) 'M is a locked Word equation'
    Check ((Control $doc.Content '#val M').Range.Text -match '54') '#val M|kNm shows 54 kNm'
    Check ((Control $doc.Content '#val k').Range.Text -match '6') 'hidden k = 2 is usable later (k*3 = 6)'
    Check ((Control $doc.Content '#hide').Range.Font.Hidden -eq -1) '#hide expression is hidden text'
    $broken = Control $doc.Content 'y = '
    Check ((StateOf $broken) -eq 'x' -and $broken.Title -like 'Error:*nothing*') 'error stays editable, message as its title'
    Check ($broken.Range.NoProofing -eq -1) 'Calcpad source is excluded from spell checking'
    Picture $doc.Content 'e2e-1-calcpad-style'

    Say 'Step 1b: what the right-click menu offers'
    CursorIn $m.Range
    Check ((Prop $api 'CalcpadAtSelection') -eq 'result') 'inside an equation: Edit Calcpad Source'
    CursorIn $broken.Range
    Check ((Prop $api 'CalcpadAtSelection') -eq 'source') 'inside failed source: Calculate Calcpad'
    $doc.Paragraphs.Item(1).Range.Select()
    Check ((Prop $api 'CalcpadAtSelection') -eq 'none') 'in ordinary text: no Calcpad items'

    Say 'Step 2: errors list and Next Error'
    Check ((Prop $api 'ErrorCount') -eq 1) 'one error is listed'
    $doc.Paragraphs.Item(1).Range.Select()
    Say ('  -> ' + (Call $api 'NextError'))
    $at = $word.Selection.Range.ParentContentControl
    Check ($null -ne $at -and (SourceOf $at) -eq 'y = nothing*2') 'Next Error puts the cursor in the failed expression'

    Say 'Step 3: Edit Source on L, change it to 8m, then click elsewhere (no Evaluate)'
    $l = Control $doc.Content 'L = '
    $r = $l.Range; $r.Collapse(1); $r.Select()
    Say ('  -> ' + (Command 'EditSource'))
    Check ($l.Range.OMaths.Count -eq 0 -and $l.Range.Text -eq 'L = 6m') 'Edit Source shows plain source text'
    $l.Range.Text = 'L = 8m'
    $doc.Paragraphs.Item(1).Range.Select()   # leave the control
    Check (WaitFor { $l.Range.OMaths.Count -eq 1 }) 'leaving the source evaluates it'
    Check ((Control $doc.Content '#val M').Range.Text -match '96') 'and everything after it follows (M = 96 kNm)'

    Say 'Step 3b: automatic calculation off (Excel''s manual mode): leaving edited source draws it alone'
    $settings = Call $api 'GetSettings'
    [void](Call $api 'SetSettings' @("$settings;auto=0"))
    function Edit-L([string]$text) {
        $r = $l.Range; $r.Collapse(1); $r.Select()
        [void](Command 'EditSource')
        if ($text) { $l.Range.Text = $text }
        $doc.Paragraphs.Item(1).Range.Select()   # leave the control
        return (WaitFor { $l.Range.OMaths.Count -eq 1 })
    }
    $m = { (Control $doc.Content '#val M').Range.Text }
    Check ((Edit-L $null) -and (& $m) -match '96') 'opened and left unchanged: the source closes as an equation'
    Check ((Edit-L 'L = 6m') -and $l.Range.Text -match '6') 'changed to 6m: L closes, showing 6 m'
    Check ((& $m) -match '96' -and (Prop $api 'LastStatus') -match 'out of date') 'but M waits for Calculate (still 96 kNm), and the results are out of date'
    Say ('  -> ' + (Command 'Recalculate'))
    Check ((& $m) -match '54') 'Calculate updates M (54 kNm)'
    Check ((Edit-L 'L = 8m') -and (Command 'Recalculate') -and (& $m) -match '96') 'back to 8m, and Calculate: M = 96 kNm'
    [void](Call $api 'SetSettings' @($settings))

    Say 'Step 4: fix the error in place, then click elsewhere'
    $broken = Control $doc.Content 'y = '
    $r = $broken.Range; $r.Collapse(0); $r.Select()
    $broken.Range.Text = 'y = L*2'
    $doc.Paragraphs.Item(1).Range.Select()
    Check (WaitFor { (Control $doc.Content 'y = ').Range.OMaths.Count -eq 1 }) 'the fixed expression renders after leaving it'
    Check ((Prop $api 'ErrorCount') -eq 0) 'no errors left'

    Say 'Step 5: expressions in a header, a footer and a text box'
    $header = $doc.Sections.Item(1).Headers.Item(1)
    $header.Range.Text = 'Header: {=#val M|kNm}'
    $footer = $doc.Sections.Item(1).Footers.Item(1)
    $footer.Range.Text = 'Footer: {=#val R}'
    $box = $doc.Shapes.AddTextbox(1, 300, 500, 220, 40)
    $box.TextFrame.TextRange.Text = 'Box: {=#val R*2}'
    $summary = Command 'Recalculate'
    Say "  -> $summary"
    Check ($summary -match '14 items, 0 errors') 'all 14 expressions evaluate'
    Check ($header.Range.ContentControls.Count -eq 1 -and $header.Range.Text -match '96') 'header shows M (96 kNm)'
    Check ($footer.Range.ContentControls.Count -eq 1 -and $footer.Range.Text -match '48') 'footer shows R (48 kN)'
    Check ($box.TextFrame.TextRange.ContentControls.Count -eq 1 -and $box.TextFrame.TextRange.Text -match '96') 'text box shows 2R (96 kN)'

    Say 'Step 6: settings -> plain style, 3 decimals'
    $r = Control $doc.Content 'R = '
    Check ($r.Range.WordOpenXML -match '0066DD') 'Calcpad style colours variables'
    Say ('  -> ' + (Command 'SetSettings' @('decimals=3;substitute=1;angles=deg;style=plain')))
    Check (-not ($r.Range.WordOpenXML -match '0066DD')) 'plain style removes the colour'
    Check ((Call $api 'GetSettings') -like 'decimals=3;substitute=1;angles=deg;style=plain;font=*') 'settings are stored in the document'

    Say 'Step 7: copy the reaction expression to the end of the document'
    $r = Control $doc.Content 'R = '
    $copy = $doc.Range($r.Range.Start - 1, $r.Range.End + 1)
    $end = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $end.InsertAfter(' Copy: ')
    $end = $doc.Range($doc.Content.End - 1, $doc.Content.End - 1)
    $end.FormattedText = $copy.FormattedText
    $summary = Command 'Recalculate'
    Say "  -> $summary"
    Check ($summary -match '15 items, 0 errors') 'pasted copy keeps its source and evaluates'

    Say 'Step 8: save, close, reopen, recalculate'
    $path = SavePath 'e2e'
    [void](Call $doc 'SaveAs2' @($path, 16))
    [void](Call $doc 'Close' @(0))
    $doc = Call $word.Documents 'Open' @($path)
    $summary = Command 'Recalculate'
    Say "  -> $summary"
    Check ($summary -match '15 items, 0 errors') 'sources survive save and reopen'
    Check ((Call $api 'GetSettings') -like '*style=plain*') 'settings survive save and reopen'
    Dump $doc
    Picture $doc.Content 'e2e-2-plain-style'
    [void](Call $doc 'Close' @(0))

    Say 'Step 9: a fenced Calcpad block between inline expressions'
    $doc = $word.Documents.Add()
    $blockLines = @(
        'L = 6m'
        '{='
        '"Beam check'
        "'The span is 'L"
        'M = 3*L'
        '#if M > 10m'
        "'Long beam"
        '#else'
        "'Short beam"
        '#end if'
        '#for i = 1 : 3'
        "'Step 'i"
        '#loop'
        '$Plot{x^2 @ x = 0 : 2}'
        '}'
        'Moment from the block: {=#val M}'
    )
    $doc.Content.Text = $blockLines -join "`r"
    $doc.Paragraphs.Item(1).Range.Select()
    $summary = Command 'Evaluate'
    Say "  -> $summary"
    $block = BlockControl $doc
    Check ($null -ne $block -and $summary -match '3 items, 0 errors') 'the fence became a block; 1 block + 2 inline items, no errors'
    $fences = @(1..$doc.Paragraphs.Count | Where-Object { $doc.Paragraphs.Item($_).Range.Text.Trim() -in '{=', '}' })
    Check ($fences.Count -eq 0) 'the {= and } fence lines are gone'
    Check ($block.Range.Text -match 'Beam check' -and $block.Range.Text -match 'Long beam' -and $block.Range.Text -notmatch 'Short beam') 'block output: heading, text, and the #if branch that applies'
    Check ($block.Range.OMaths.Count -ge 4) "block output has its equations (M and 3 loop steps): $($block.Range.OMaths.Count)"
    Check ($block.Range.InlineShapes.Count -ge 1) 'the plot is a picture in the block'
    Check ((Control $doc.Content '#val M').Range.Text -match '18') 'an inline expression after the block uses its variable (M = 18 m)'
    Check ($block.LockContents) 'block output is read-only'
    CursorIn $block.Range
    Check ((Prop $api 'CalcpadAtSelection') -eq 'result') 'right-click in a block offers Edit Calcpad Source'
    Picture $doc.Content 'e2e-3-block'

    Say 'Step 10: edit the block source, then click outside'
    CursorIn $block.Range
    Say ('  -> ' + (Command 'EditSource'))
    Check ($block.Range.OMaths.Count -eq 0 -and $block.Range.Text -match '#if M > 10m') 'Edit Source shows the script, one line per paragraph'
    $script = ($blockLines[2..13] -replace '^M = 3\*L$', 'M = L') -join "`r"
    $block.Range.Text = $script
    $doc.Paragraphs.Item(1).Range.Select()
    Check (WaitFor { $block.Tag -like 'calcpad-block:r:*' }) 'leaving the block evaluates it'
    Check ($block.Range.Text -match 'Short beam' -and $block.Range.Text -notmatch 'Long beam') 'the other #if branch now applies'
    Check ((Control $doc.Content '#val M').Range.Text -match '6') 'and the inline expression after it follows (M = 6 m)'

    Say 'Step 11: Calcpad Block on an empty line, with an error inside'
    $doc.Content.InsertParagraphAfter()
    $doc.Paragraphs.Last.Range.Select()
    Say ('  -> ' + (Command 'MakeBlock'))
    $empty = BlockControl $doc 'e'
    Check ($null -ne $empty) 'an empty block is ready to type into'
    $empty.Range.Text = "u = nothing*2`rw = 3"
    $doc.Paragraphs.Item(1).Range.Select()
    Check (WaitFor { $empty.Tag -like 'calcpad-block:r:*' }) 'leaving it evaluates it'
    Check ($empty.Title -like 'Error:*nothing*') 'its error is reported in its title'
    Check ((Prop $api 'ErrorCount') -eq 1) 'and counted in the error list'

    Say 'Step 12: Evaluate on lines that read like a script makes a block; an unclosed #if is reported'
    $doc.Content.InsertAfter("'Check`r#if L > 1m")
    $n = $doc.Paragraphs.Count
    $doc.Range($doc.Paragraphs.Item($n - 1).Range.Start, $doc.Paragraphs.Item($n).Range.End).Select()
    Say ('  -> ' + (Command 'Evaluate'))
    $bad = BlockControl $doc 'x'
    Check ($null -ne $bad -and $bad.Range.Text -match '#if L > 1m' -and $bad.Title -like 'Error:*#if*') 'it became a block that shows its source and the problem'
    Check ((Prop $api 'ErrorCount') -eq 2) 'two problems listed'
    Check ((Control $doc.Content '#val M').Range.Text -match '6') 'the rest of the document still calculates'

    Say 'Step 13: blocks survive save and reopen'
    $path = SavePath 'e2e-blocks'
    [void](Call $doc 'SaveAs2' @($path, 16))
    [void](Call $doc 'Close' @(0))
    $doc = Call $word.Documents 'Open' @($path)
    $summary = Command 'Recalculate'
    Say "  -> $summary"
    Check ($summary -match '5 items, 2 errors') 'same items and errors after reopening'
    Picture $doc.Content 'e2e-4-blocks-final'
    [void](Call $doc 'Close' @(0))

    Say 'Step 14: Shift+Enter adds expressions while editing; Enter finishes'
    $doc = $word.Documents.Add()
    $doc.Content.Text = "a = 2`rb = a*3`rText after."
    $doc.Range($doc.Paragraphs.Item(1).Range.Start, $doc.Paragraphs.Item(2).Range.End).Select()
    [void](Command 'Evaluate')
    $a = Control $doc.Content 'a = '
    CursorIn $a.Range
    [void](Command 'EditSource')
    Check ((Prop $api 'EnterLeavesControl') -eq $true) 'while editing a source, Enter belongs to Calcpad'
    $word.Selection.TypeText("`v")           # Shift+Enter: Word's line break
    $word.Selection.TypeText('c = a + 1')
    $word.Selection.TypeText("`v")
    $word.Selection.TypeText('d = c*2')
    Check ((StateOf $a) -eq 'e' -and $a.Range.Text -match 'c = a \+ 1' -and $doc.ContentControls.Count -eq 2) 'Shift+Enter adds lines inside the source'
    Check ((Call $api 'PressEnter') -eq 'left') 'Enter leaves the source'
    Check (WaitFor { $doc.ContentControls.Count -eq 4 }) 'and each line becomes its own expression'
    $order = @(1..$doc.ContentControls.Count | ForEach-Object { SourceOf $doc.ContentControls.Item($_) })
    Check (($order -join ' | ') -eq 'a = 2 | c = a + 1 | d = c*2 | b = a*3') "in the order typed: $($order -join ' | ')"
    Check (@(1..4 | Where-Object { $doc.ContentControls.Item($_).Range.OMaths.Count -eq 1 }).Count -eq 4) 'all four are evaluated equations'
    Check ((Control $doc.Content 'd = ').Range.Text -match '6') 'd = c*2 = 6'
    $here = $word.Selection.Range
    Check ($null -eq $here.ParentContentControl -and $here.Paragraphs.Item(1).Range.Text -eq "`r" -and $doc.Paragraphs.Count -eq 4) 'the cursor is in a new, ordinary paragraph below'
    Check ((Prop $api 'EnterLeavesControl') -eq $false) 'where Enter is ordinary again'

    $b = Control $doc.Content 'b = '
    CursorIn $b.Range
    Check ((Prop $api 'EnterLeavesControl') -eq $true -and (Call $api 'PressEnter') -eq 'left') 'Enter on a (locked) result moves on as well'
    $here = $word.Selection.Range
    Check ($null -eq $here.ParentContentControl -and $doc.Paragraphs.Count -eq 5) 'to a new paragraph after it'

    $last = $doc.Paragraphs.Last.Range
    $doc.Range($last.End - 1, $last.End - 1).Select()
    Check ((Prop $api 'EnterLeavesControl') -eq $false -and (Call $api 'PressEnter') -eq 'typed' -and $doc.Paragraphs.Count -eq 6) 'in ordinary text Enter is an ordinary Enter'
    [void](Call $doc 'Close' @(0))

    Say 'Step 15: a real worksheet (Calcpad soil stress example) pasted as text, then Evaluate'
    $fixture = Join-Path (Split-Path $PSScriptRoot -Parent) 'tests\CalcpadWord.Engine.Tests\Fixtures\SoilStress.cpd'
    $soil = @([IO.File]::ReadAllText($fixture, [Text.Encoding]::UTF8).TrimEnd() -split "`r?`n")
    $doc = $word.Documents.Add()
    $doc.Content.Text = ($soil -join "`r") + "`rThe map reaches {=#val k*a} either side of the centre."
    $doc.Range(0, $doc.Paragraphs.Item($soil.Count).Range.End).Select()
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $summary = Call $api 'Evaluate'
    Say ("  -> {0}  ({1:F1} s)" -f $summary, $timer.Elapsed.TotalSeconds)
    Check ($summary -match 'background' -and $timer.Elapsed.TotalSeconds -lt 5) ("the long calculation runs in the background; Evaluate returned after {0:F1} s" -f $timer.Elapsed.TotalSeconds)
    $typing = [Diagnostics.Stopwatch]::StartNew()
    $doc.Range($doc.Content.End - 1, $doc.Content.End - 1).InsertAfter(' Typed while Calcpad calculates.')
    Check ((Prop $api 'IsWorking') -and $typing.Elapsed.TotalSeconds -lt 2) ("Word responds meanwhile: typing took {0:F2} s" -f $typing.Elapsed.TotalSeconds)
    Check (WaitFor { -not (Prop $api 'IsWorking') } 180000) ("the results arrive ({0:F1} s after Evaluate)" -f $timer.Elapsed.TotalSeconds)
    $summary = Prop $api 'LastStatus'
    Say "  -> $summary"
    Check ($doc.Paragraphs.Last.Range.Text -match 'Typed while Calcpad calculates') 'and the text typed meanwhile is kept'
    $blocks = @(1..$doc.ContentControls.Count | ForEach-Object { $doc.ContentControls.Item($_) } | Where-Object { $_.Tag -like 'calcpad-block:*' })
    Check ($summary -match ' 0 errors') 'the pasted script calculates without errors'
    Check ($doc.Paragraphs.Item(1).OutlineLevel -eq 1 -and $doc.Paragraphs.Item(1).Range.Text -match '^Soil stress') 'its title is a Word Heading 1'
    $results = @(1..$doc.Paragraphs.Count | Where-Object { $doc.Paragraphs.Item($_).Range.Text -match '^Results\r' })
    Check ($results.Count -eq 1 -and $doc.Paragraphs.Item($results[0]).OutlineLevel -eq 2) '<h4>Results</h4> is a Word Heading 2'
    Check ($doc.Paragraphs.Item(2).Borders.Item(-3).LineStyle -ne 0) '<hr/> is a rule line'
    $load = @(1..$doc.Paragraphs.Count | Where-Object { $doc.Paragraphs.Item($_).Range.Text -match '^Surface load - ' })
    Check ($load.Count -eq 1 -and $doc.Paragraphs.Item($load[0]).Range.ContentControls.Count -eq 1 -and $null -eq $doc.Paragraphs.Item($load[0]).Range.Characters.Item(1).ParentContentControl) 'text is a Word paragraph with the value inline'
    Check ($blocks.Count -eq 2) "only plots and the table stay Calcpad blocks ($($blocks.Count))"
    Check ($doc.InlineShapes.Count -eq 6 -and $doc.Tables.Count -ge 1) "4 plots and 2 maps are pictures, the results table a Word table ($($doc.InlineShapes.Count) pictures)"
    Check ((Control $doc.Content '#val k*a').Range.Text -match '3' -and $doc.Paragraphs.Last.Range.Text -match '^The map reaches') 'the sentence after it is untouched and uses k, set inside an HTML comment (k*a = 3 m)'
    Picture $doc.Content 'e2e-5-soil-stress'
    $tableBlock = @($blocks | Where-Object { $_.Range.Tables.Count -ge 1 })[0]
    CursorIn $tableBlock.Range
    [void](Command 'EditSource')
    Check ($tableBlock.Tag -like 'calcpad-block:e:*' -and $tableBlock.Range.Tables.Count -eq 0 -and $tableBlock.Range.Text -match '\$Map') 'Edit Source works on the block holding the table'
    $doc.Paragraphs.Item(1).Range.Select()
    Check (WaitFor { $tableBlock.Tag -like 'calcpad-block:r:*' } 90000) 'and leaving it draws the block again'
    Check ($tableBlock.Range.Tables.Count -ge 1 -and $tableBlock.Range.InlineShapes.Count -eq 5) "with its table and pictures ($($tableBlock.Range.InlineShapes.Count))"
    $timer.Restart()
    $summary = Command 'Recalculate'
    $seconds = $timer.Elapsed.TotalSeconds
    Check ($seconds -lt 5 -and $summary -match ' 0 errors') ("recalculating with the blocks unchanged is quick ({0:F1} s)" -f $seconds)
    $path = SavePath 'e2e-soil-stress'
    [void](Call $doc 'SaveAs2' @($path, 16))
    [void](Call $doc 'Close' @(0))
    $doc = Call $word.Documents 'Open' @($path)
    $timer.Restart()
    $summary = Command 'Recalculate'
    $seconds = $timer.Elapsed.TotalSeconds
    Check ($seconds -lt 5 -and $summary -match ' 0 errors') ("and stays quick after saving and reopening ({0:F1} s)" -f $seconds)
    [void](Call $doc 'Close' @(0))

    Say 'Step 15b: Cancel, and results greyed while they may be out of date'
    $doc = $word.Documents.Add()
    $doc.Content.Text = ($soil -join "`r")
    $doc.Content.Select()
    $started = Call $api 'Evaluate'
    Check ($started -match 'background' -and (Prop $api 'CanCancel')) 'while the soil example calculates, Cancel is enabled'
    $timer = [Diagnostics.Stopwatch]::StartNew()
    $summary = Call $api 'Cancel'
    Say "  -> $summary"
    $stopped = WaitFor { -not (Prop $api 'IsWorking') } 15000
    Check ($summary -match 'cancelled the calculation' -and $stopped -and -not (Prop $api 'CanCancel')) ("Cancel stops it ({0:F1} s) and is disabled again" -f $timer.Elapsed.TotalSeconds)
    Check ($doc.InlineShapes.Count -eq 0 -and (Prop $api 'LastStatus') -match 'cancelled') 'no results were drawn after cancelling'
    Check ((Call $api 'Cancel') -match 'nothing to cancel') 'with nothing running, Cancel does nothing'
    $summary = Command 'Recalculate'
    Check ($summary -match ' 0 errors' -and $doc.InlineShapes.Count -eq 6) "Recalculate then calculates it ($summary)"

    function Stale($doc) { @(1..$doc.ContentControls.Count | ForEach-Object { $doc.ContentControls.Item($_) } | Where-Object { $_.Title -like 'Out of date: *' }) }
    # Word's own bookkeeping, renewed when anything is touched: edit IDs, placeholders, and the
    # effect margin of a picture whose colour mode was switched and back (0.75 pt, not visible).
    function Body($doc) { [regex]::Replace(([regex]::Match($doc.Content.WordOpenXML, '<w:body>.*</w:body>', 'Singleline').Value), ' w:rsid\w*="[^"]*"|<w:docPart w:val="[^"]*"/>| wp14:(anchorId|editId)="[^"]*"|<wp:effectExtent [^>]*/>', '') }
    $original = Body $doc
    $settings = Call $api 'GetSettings'
    $results = @(1..$doc.ContentControls.Count | Where-Object { $doc.ContentControls.Item($_).Tag -like 'calcpad*:r:*' }).Count
    $started = Call $api 'SetSettings' @($settings -replace 'decimals=\d+', 'decimals=4')
    Check ($started -match 'background' -and (Stale $doc).Count -eq 0) 'not greyed at once: a calculation that ends soon after going to the background is drawn without greying'
    [void](WaitFor { (Stale $doc).Count -eq $results } 5000)
    $grey = Stale $doc
    $plot = @($grey | Where-Object { $_.Range.InlineShapes.Count -gt 0 })[0]
    $value = @($grey | Where-Object { $_.Tag -like 'calcpad:*' })[0]
    Check ((Prop $api 'IsWorking') -and $grey.Count -eq $results) "still calculating a moment later, all $results results are greyed ($($grey.Count))"
    Check ($null -ne $plot -and $plot.Range.InlineShapes.Item(1).PictureFormat.ColorType -eq 2 -and $null -ne $value -and $value.Range.Font.Shading.BackgroundPatternColor -eq 15132390) 'greyed: plots in grey, light grey behind values'
    Check ((Prop $api 'ErrorCount') -eq 0) 'greying does not touch the error list'
    [void](Call $api 'Cancel')
    [void](WaitFor { -not (Prop $api 'IsWorking') } 15000)
    Check ((Stale $doc).Count -eq $results -and (Prop $api 'LastStatus') -match 'stay greyed') 'after Cancel they stay greyed, and the status says so'
    $summary = Command 'SetSettings' @($settings)
    Check ((Stale $doc).Count -eq 0 -and $summary -match ' 0 errors') "recalculating removes the grey ($summary)"
    Check ((Body $doc) -eq $original) 'results that did not change are exactly as before (not redrawn, only ungreyed)'
    [void](Call $api 'SetSettings' @($settings -replace 'decimals=\d+', 'decimals=4'))
    [void](Call $api 'Cancel')
    [void](WaitFor { -not (Prop $api 'IsWorking') } 15000)
    Check ((Stale $doc).Count -eq $results) 'cancelled before the grey came, the results are greyed all the same'
    [void](Call $doc 'Close' @(0))

    Say 'Step 16: Block to Text on an existing block'
    $doc = $word.Documents.Add()
    $doc.Content.Text = @('{=', '"Section', "'Width 'w = 3", '$Plot{x @ x = 0 : w}', "'<table><tr><td>", "'cell 'w", "'</td></tr></table>", '}', 'End.') -join "`r"
    $doc.Paragraphs.Item(1).Range.Select()
    [void](Command 'Evaluate')
    $block = BlockControl $doc
    CursorIn $block.Range
    Check ($null -ne $block -and (Prop $api 'CalcpadAtSelection') -eq 'result') 'a fenced block to start from'
    $summary = Command 'ConvertBlockToText'
    Say "  -> $summary"
    $blocks = @(1..$doc.ContentControls.Count | ForEach-Object { $doc.ContentControls.Item($_) } | Where-Object { $_.Tag -like 'calcpad-block:*' })
    Check ($doc.Paragraphs.Item(1).OutlineLevel -eq 1 -and $doc.Paragraphs.Item(1).Range.Text -match '^Section') 'the heading is a Word heading'
    Check ($doc.Paragraphs.Item(2).Range.Text -match '^Width ' -and (Control $doc.Content 'w = ').Range.OMaths.Count -eq 1) 'the text line is a Word paragraph with w inline'
    Check ($blocks.Count -eq 1 -and $blocks[0].Range.InlineShapes.Count -eq 1 -and $blocks[0].Range.Tables.Count -eq 1 -and $summary -match ' 0 errors') 'only the plot and the table are still a block (it had a table: Word refuses to replace those directly)'
    Check ($doc.Paragraphs.Last.Range.Text -match '^End\.') 'the text after the block is untouched'
    [void](Call $doc 'Close' @(0))

    Say 'Step 17: math fonts (free ones are downloaded and installed for the current user)'
    $doc = $word.Documents.Add()
    $doc.Content.Text = "F = 120kN`rb = 0.3m`rh = 0.5m`r$([char]0x03C3) = F/(b*h)|MPa"
    $doc.Content.Select()
    [void](Command 'Evaluate')
    $fonts = @((Prop $api 'MathFontList') -split '\|')
    $names = @($fonts | ForEach-Object { $_.Substring(1) })
    $starters = 'Asana Math', 'Cambria Math', 'Euler Math', 'Fira Math', 'Latin Modern Math', 'Libertinus Math', 'Noto Sans Math', 'STIX Two Math', 'TeX Gyre Pagella Math', 'XITS Math'
    $missing = @($starters | Where-Object { $names -notcontains $_ })
    Check ($missing.Count -eq 0 -and $names -notcontains 'Minion Math') "the starters are on offer, Fira Math instead of Minion Math: $($fonts -join ', ')"
    Check (($names -join '|') -eq ((@($names | Sort-Object) -join '|'))) 'the list is in alphabetical order'   # Sort-Object: culture-aware, ignoring case, like the ribbon
    $found = @($names | Where-Object { $starters -notcontains $_ })
    Say "  math fonts found on this computer besides the starters: $(if ($found) { $found -join ', ' } else { 'none' })"
    try { [void](Call $api 'SetMathFont' @('No Such Math')); $refused = $false } catch { $refused = $_.Exception.InnerException.Message -match 'not one of the math fonts' }
    Check $refused 'a font that is not on offer is refused'
    foreach ($family in 'Latin Modern Math', 'STIX Two Math', 'Fira Math', 'Libertinus Math', 'XITS Math', 'TeX Gyre Pagella Math', 'Asana Math', 'Noto Sans Math', 'Euler Math', 'Cambria Math') {
        $timer = [Diagnostics.Stopwatch]::StartNew()
        [void](Call $api 'SetMathFont' @($family))
        $done = WaitFor { -not (Prop $api 'IsWorking') } 120000
        $math = [regex]::Match((Control $doc.Content 'F = ').Range.WordOpenXML, '<m:oMath>.*</m:oMath>').Value
        $used = @([regex]::Matches($math, 'w:ascii="([^"]+)"') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
        Check ($done -and (Prop $api 'MathFont') -eq $family -and $doc.OMathFontName -eq $family -and ($used -join ',') -eq $family) ("{0}: installed and used by every part of the equations ({1:F1} s; fonts in F = 120kN: {2})" -f $family, $timer.Elapsed.TotalSeconds, ($used -join ', '))
        Picture $doc.Content ('e2e-font-' + ($family -replace ' ', '-'))
    }
    $fonts = @((Prop $api 'MathFontList') -split '\|')
    Check (@($fonts | Where-Object { $_ -like '-*' }).Count -eq 0) "all of them are installed now: $($fonts -join ', ')"
    [void](Call $doc 'Close' @(0))

    Say 'Step 18: one Undo reverts an Evaluate'
    $doc = $word.Documents.Add()
    $doc.Content.Text = 'a = 1 + 2'
    $doc.Content.Select()
    [void](Command 'Evaluate')
    $ok = $doc.ContentControls.Count -eq 1
    [void](Call $doc 'Undo' @(1))
    Check ($ok -and $doc.ContentControls.Count -eq 0 -and $doc.Content.Text -eq "a = 1 + 2`r") 'single undo step'
    [void](Call $doc 'Close' @(0))
}
catch {
    Say ("ERROR: " + $_.Exception.Message + " (line " + $_.InvocationInfo.ScriptLineNumber + ")")
    $script:failures++
}
finally {
    # Opening Word while the test runs can put your window in the test's hidden Word (Windows
    # reuses a running Word). Then never quit it: that would close your documents unsaved.
    # The test's Word is hidden; a window you open in it makes it visible.
    $yours = @()
    try {
        if ($word.Visible) {
            for ($i = 1; $i -le $word.Documents.Count; $i++) { $yours += $word.Documents.Item($i).Name }
            if ($yours.Count -eq 0) { $yours += '(a visible window)' }
        }
    } catch { }
    if ($yours.Count -gt 0) {
        Set-Content (Join-Path $OutDir 'word.in-use') ($yours -join ', ')
        Say "  FAIL  someone opened Word during the test and it joined the test's Word ($($yours -join ', ')): results may be off. Word is left running; close it yourself."
        $script:failures++
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
        exit $script:failures
    }
    Say 'Quitting Word'
    try { [void](Call $word 'Quit' @(0)) } catch { Say "  Quit failed: $($_.Exception.Message)"; $script:failures++ }
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    $word = $null; $api = $null; $addin = $null; $doc = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    $gone = $null -eq (Get-Process -Id $new[0] -ErrorAction SilentlyContinue)
    if (-not $gone) { $gone = (Get-Process -Id $new[0]).WaitForExit(20000) }
    Check $gone 'Word exits after Quit (the add-in does not keep it alive)'
}

$addinLog = Join-Path $env:LOCALAPPDATA 'CalcpadWord\addin.log'
if (Test-Path $addinLog) {
    Say '--- add-in log (this run)'
    $since = (Get-Process -Id $PID).StartTime.ToString('yyyy-MM-dd HH:mm:ss')
    $thisRun = @(Get-Content $addinLog -Encoding UTF8 | Where-Object { $_.Length -ge 19 -and $_.Substring(0, 19) -ge $since })
    $thisRun | ForEach-Object { Say "  $_" }
    Check (-not ($thisRun -match ' ERROR ')) 'the add-in logged no errors'
}
Say $(if ($script:failures -eq 0) { 'ALL PASSED' } else { "$($script:failures) FAILED" })
exit $script:failures
