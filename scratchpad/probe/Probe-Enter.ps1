# Probe: what Word does with Enter (Selection.TypeParagraph) inside an inline rich-text content control.
param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value $m -Encoding UTF8 }
function Show([string]$s) { $s -replace "`r", '\r' -replace "`a", '\a' -replace "`v", '\v' }
function Report($doc, [string]$label) {
    Step ("{0}: paragraphs={1} ccs={2} doc=[{3}]" -f $label, $doc.Paragraphs.Count, $doc.ContentControls.Count, (Show $doc.Content.Text))
    for ($i = 1; $i -le $doc.ContentControls.Count; $i++) {
        $cc = $doc.ContentControls.Item($i)
        $cs = $cc.Range.Start; $ce = $cc.Range.End
        Step ("    cc{0} tag={1} range={2}..{3} paragraphsInside={4} text=[{5}]" -f $i, $cc.Tag, $cs, $ce, $cc.Range.Paragraphs.Count, (Show $cc.Range.Text))
    }
    $sel = $word.Selection.Range
    $parent = $sel.ParentContentControl
    Step ("    cursor at {0}, inside cc: {1}" -f $sel.Start, $(if ($parent) { $parent.Tag } else { 'none' }))
}

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false

    # Case 1: an expression alone in its paragraph, cursor at its end (the normal Edit Source state).
    $doc = $word.Documents.Add()
    $doc.Content.Text = "L = 6m`rAfter"
    $cc = $doc.ContentControls.Add(0, $doc.Range(0, 6)); $cc.Tag = 'probe-own-paragraph'
    $r = $cc.Range; $r.Collapse(0); $r.Select()
    Report $doc 'case 1 before'
    $word.Selection.TypeParagraph()
    Report $doc 'case 1 after Enter'
    $word.Selection.TypeText('q = 3kN/m')
    $word.Selection.TypeParagraph()
    $word.Selection.TypeText('M = q*L^2/8')
    Report $doc 'case 1 after typing two more lines'

    # Unwrap the control but keep its text; does a range around it follow the text?
    $cc = $doc.ContentControls.Item(1)
    $outer = $doc.Range($cc.Range.Start - 1, $cc.Range.End + 1)
    Step ("    outer before unwrap: {0}..{1} [{2}]" -f $outer.Start, $outer.End, (Show $outer.Text))
    $cc.Delete($false)
    Step ("    outer after unwrap:  {0}..{1} [{2}] ccs={3}" -f $outer.Start, $outer.End, (Show $outer.Text), $doc.ContentControls.Count)
    $doc.Close([ref]0)

    # Case 2: an expression in the middle of a sentence, cursor at its end.
    $doc = $word.Documents.Add()
    $doc.Content.Text = "The span is X metres."
    $cc = $doc.ContentControls.Add(0, $doc.Range(12, 13)); $cc.Tag = 'probe-in-sentence'
    $r = $cc.Range; $r.Collapse(0); $r.Select()
    $word.Selection.TypeParagraph()
    $word.Selection.TypeText('y = 2')
    Report $doc 'case 2 after Enter + typing'
    $doc.Close([ref]0)

    # Case 3: cursor in the middle of the expression text.
    $doc = $word.Documents.Add()
    $doc.Content.Text = "a = 1b = 2"
    $cc = $doc.ContentControls.Add(0, $doc.Range(0, 10)); $cc.Tag = 'probe-split-middle'
    $r = $doc.Range(5, 5); $r.Select()
    $word.Selection.TypeParagraph()
    Report $doc 'case 3 after Enter in the middle'
    $doc.Close([ref]0)
    Step 'Done.'
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { $word.Quit([ref]0) } catch { } }
