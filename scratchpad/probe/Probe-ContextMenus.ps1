param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value $m -Encoding UTF8 }

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $cbs = $word.CommandBars

    # 1. Legacy pop-up menus (context menus) whose names hint at equations, content controls or text.
    $popups = for ($i = 1; $i -le $cbs.Count; $i++) {
        $cb = $cbs.Item($i)
        if ($cb.Type -eq 2) { $cb.Name }
    }
    Step ("Popup count: {0}" -f $popups.Count)
    Step ("Matching popups: " + (($popups | Where-Object { $_ -match 'equat|math|content|^text$|field|form' }) -join ' | '))

    # 2. Calibrate GetLabelMso: a known context menu, a known button, and a bogus id.
    foreach ($id in 'ContextMenuText', 'Bold', 'ContextMenuDoesNotExist') {
        try { Step ("GetLabelMso {0} -> '{1}'" -f $id, $cbs.GetLabelMso($id)) }
        catch { Step ("GetLabelMso {0} -> ERROR {1}" -f $id, $_.Exception.Message) }
    }

    # 3. Candidate ribbon idMso values for the menu shown inside an equation / content control.
    $candidates = 'ContextMenuEquation', 'ContextMenuEquationArea', 'ContextMenuEquations', 'ContextMenuMath',
        'ContextMenuMathStructure', 'ContextMenuMathEquation', 'ContextMenuOMath', 'ContextMenuEquationText',
        'ContextMenuContentControl', 'ContextMenuContentControls', 'ContextMenuTextContentControl',
        'ContextMenuTextLinked', 'ContextMenuHeading', 'ContextMenuList', 'ContextMenuFields', 'ContextMenuTable',
        'ContextMenuHeaderArea', 'ContextMenuFooterArea', 'ContextMenuTextBox', 'ContextMenuSpell', 'ContextMenuGrammar'
    foreach ($id in $candidates) {
        try { [void]$cbs.GetLabelMso($id); Step "VALID   $id" }
        catch { Step "invalid $id" }
    }

    # 4. Every popup whose derived id ("ContextMenu" + name without spaces) is valid.
    $derived = foreach ($name in $popups) {
        $id = 'ContextMenu' + ($name -replace '[^A-Za-z0-9]', '')
        try { [void]$cbs.GetLabelMso($id); $id } catch { }
    }
    Step ("Derived valid ids: " + ($derived -join ', '))
    Step ("All popup names: " + ($popups -join ' | '))
    $doc.Close([ref]0)
}
catch { Step ("FAILED: " + $_.Exception.Message) }
finally { try { $word.Quit([ref]0) } catch { Step "quit failed: $($_.Exception.Message)" } }
