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
    $ids = 'Bold', 'Copy', 'Paste', 'FunctionWizard', 'CalculateNow', 'ContentControlRichText', 'EquationLinearFormat',
           'TraceError', 'ReviewNextComment', 'EquationInsertGallery', 'EquationProfessionalFormat', 'ReviewNextChange',
           'ReviewPreviousComment', 'ContentControlText', 'ContentControlsGroup', 'GroupContentControls', 'VisualBasic',
           'MacroPlay', 'Refresh', 'RefreshAll', 'SpellingAndGrammar', 'Spelling', 'ErrorChecking', 'InsertEquation',
           'EquationEdit', 'ControlsGallery', 'TableInsert', 'CodeEdit', 'DesignMode', 'ContentControlBuildingBlockGallery',
           'ContentControlPicture', 'ContentControlComboBox', 'SymbolInsert', 'FormulaMoreFunctionsMenu', 'AutoSum',
           'ReviewDeleteComment', 'DataValidation', 'ProofingErrorChecking', 'WordCountList', 'ResearchPane'
    foreach ($id in $ids) {
        $label = try { "'" + $cbs.GetLabelMso($id) + "'" } catch { 'NO-ID' }
        $image = try { $p = $cbs.GetImageMso($id, 32, 32); if ($p) { 'image' } else { 'null' } } catch { 'NO-IMAGE: ' + $_.Exception.InnerException.Message + $_.Exception.Message.Substring(0, [Math]::Min(80, $_.Exception.Message.Length)) }
        $enabled = try { $cbs.GetEnabledMso($id) } catch { 'n/a' }
        Step ("{0,-36} label={1,-30} image={2} enabled={3}" -f $id, $label, $image, $enabled)
    }
    $doc.Close([ref]0)
}
catch { Step ("FAILED: " + $_.Exception.Message) }
finally { try { $word.Quit([ref]0) } catch { } }
