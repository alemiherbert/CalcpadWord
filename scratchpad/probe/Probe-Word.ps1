# Probe: content-control Tag/Title limits and OMML insertion via InsertXML.
# Runs in its OWN hidden Word instance; refuses to touch an already-running Word.
param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value ((Get-Date -Format 'HH:mm:ss.fff') + '  ' + $m) -Encoding UTF8 }
function Flat([string]$body) {
@"
<?xml version="1.0" standalone="yes"?>
<?mso-application progid="Word.Document"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
 <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>
  <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
 </pkg:xmlData></pkg:part>
 <pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>
  <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body>$body</w:body></w:document>
 </pkg:xmlData></pkg:part>
</pkg:package>
"@
}
function Show([string]$s) { $s -replace "`r", '\r' -replace "`a", '\a' }

Step 'start'
$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
Step 'creating Word.Application'
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
Step "created; new Word PIDs: $new"
if ($new.Count -ne 1) { Step 'Not a separate Word instance - aborting without changes.'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]

try {
    Step 'setting Visible/DisplayAlerts'
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $n = $word.COMAddIns.Count
    Step "COM add-ins: $n"
    for ($i = 1; $i -le $n; $i++) { $a = $word.COMAddIns.Item($i); Step ("  {0} connected={1}" -f $a.ProgId, $a.Connect) }

    Step 'adding document'
    $doc = $word.Documents.Add()
    $doc.Content.Text = "Before X middle Y after"
    Step ("Body font: {0} {1}pt" -f $doc.Content.Font.Name, $doc.Content.Font.Size)

    # --- Tag / Title capacity
    $cc = $doc.ContentControls.Add(0, $doc.Range(7, 8))   # rich text around "X"
    foreach ($len in 64, 65, 255, 256, 1000, 5000) {
        try { $cc.Tag = ('t' * $len); Step ("Tag   set {0,5} -> read {1}" -f $len, $cc.Tag.Length) }
        catch { Step ("Tag   set {0,5} -> ERROR {1}" -f $len, $_.Exception.Message) }
    }
    foreach ($len in 64, 65, 255, 256) {
        try { $cc.Title = ('t' * $len); Step ("Title set {0,5} -> read {1}" -f $len, $cc.Title.Length) }
        catch { Step ("Title set {0,5} -> ERROR {1}" -f $len, $_.Exception.Message) }
    }
    $uni = 'c = sqrt(a^2 + b^2) σ_x <>&"''{}|;'
    $cc.Tag = $uni
    Step ("Tag unicode/specials round-trip: " + ($cc.Tag -eq $uni))

    # --- OMML into the control
    $omml = '<m:r><w:rPr><w:color w:val="0066DD"/></w:rPr><m:t>c</m:t></m:r><m:r><m:t>=</m:t></m:r><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f><m:r><m:t>=</m:t></m:r><m:r><m:t>5</m:t></m:r><m:r><m:t> </m:t></m:r><m:r><m:rPr><m:nor/></m:rPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>m</m:t></m:r>'
    Step 'InsertXML #1'
    $cc.Range.InsertXML((Flat "<w:p><m:oMath>$omml</m:oMath></w:p>"))
    Step ("After InsertXML #1: paragraphs={0} ccs={1} omaths={2}" -f $doc.Paragraphs.Count, $doc.ContentControls.Count, $doc.OMaths.Count)
    Step ("  doc text: [{0}]" -f (Show $doc.Content.Text))
    Step ("  math font: {0} {1}pt" -f $doc.OMaths.Item(1).Range.Font.Name, $doc.OMaths.Item(1).Range.Font.Size)

    # --- Re-render: lock/unlock, replace with text, then with math again
    $cc.LockContents = $true
    $cc.LockContents = $false
    $cc.Range.Text = 'error text'
    Step ("After Range.Text: ccs={0} cc text=[{1}]" -f $doc.ContentControls.Count, $cc.Range.Text)
    $cc.Range.InsertXML((Flat "<w:p><m:oMath>$omml</m:oMath></w:p>"))
    Step ("After InsertXML #2: paragraphs={0} ccs={1} omaths={2} doc text=[{3}]" -f $doc.Paragraphs.Count, $doc.ContentControls.Count, $doc.OMaths.Count, (Show $doc.Content.Text))

    # --- Second control in the same paragraph (around "Y")
    $r2 = $doc.Content; $f = $r2.Find; $f.Text = 'Y'; [void]$f.Execute()
    $cc2 = $doc.ContentControls.Add(0, $r2)
    $cc2.Range.InsertXML((Flat "<w:p><m:oMath>$omml</m:oMath></w:p>"))
    Step ("Two controls: paragraphs={0} ccs={1} omaths={2} doc text=[{3}]" -f $doc.Paragraphs.Count, $doc.ContentControls.Count, $doc.OMaths.Count, (Show $doc.Content.Text))
    $starts = for ($i = 1; $i -le $doc.ContentControls.Count; $i++) { $doc.ContentControls.Item($i).Range.Start }
    Step ("  cc starts in collection order: " + ($starts -join ', '))

    Step 'exporting'
    $doc.ExportAsFixedFormat((Join-Path $OutDir 'probe.pdf'), 17)
    $doc.SaveAs2((Join-Path $OutDir 'probe.docx'), 16)
    $doc.Close(0)
    Step 'Done.'
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ line " + $_.InvocationInfo.ScriptLineNumber) }
finally {
    Step 'quitting'
    $word.Quit(0)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
    Step 'quit'
}
