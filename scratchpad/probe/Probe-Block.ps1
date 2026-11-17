# Probe: long multi-line tags, and a block-level content control filled with several paragraphs via InsertXML.
param([Parameter(Mandatory)][string]$OutDir)
$ErrorActionPreference = 'Stop'
$log = Join-Path $OutDir 'steps.log'
function Step([string]$m) { Add-Content -Path $log -Value $m -Encoding UTF8 }
function Show([string]$s) { $s -replace "`r", '\r' -replace "`n", '\n' -replace "`a", '\a' }
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

$before = @(Get-Process WINWORD -ErrorAction SilentlyContinue | ForEach-Object Id)
$word = New-Object -ComObject Word.Application
$new = @(Get-Process WINWORD | ForEach-Object Id | Where-Object { $before -notcontains $_ })
if ($new.Count -ne 1) { Step 'not a separate instance'; return }
Set-Content (Join-Path $OutDir 'pid.txt') $new[0]
try {
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $doc.Content.Text = "Before`rline one`rline two`rline three`rAfter"
    Step ("paragraphs: {0}" -f $doc.Paragraphs.Count)

    # Block-level control around paragraphs 2..4 (whole paragraphs, including the last paragraph mark).
    $r = $doc.Range($doc.Paragraphs.Item(2).Range.Start, $doc.Paragraphs.Item(4).Range.End)
    $cc = $doc.ContentControls.Add(0, $r)
    Step ("block cc: paragraphs inside={0}, doc paragraphs={1}, cc text=[{2}]" -f $cc.Range.Paragraphs.Count, $doc.Paragraphs.Count, (Show $cc.Range.Text))

    # Tags: multi-line and long.
    $line = 'x = 1' + "`n"
    foreach ($n in 2000, 10000, 40000, 70000, 140000) {
        $value = ($line * [math]::Ceiling($n / $line.Length)).Substring(0, $n)
        try { $cc.Tag = $value; Step ("Tag {0,6} chars with \n -> read {1}, equal={2}" -f $n, $cc.Tag.Length, ($cc.Tag -eq $value)) }
        catch { Step ("Tag {0,6} -> ERROR {1}" -f $n, $_.Exception.Message) }
    }
    $cc.Tag = 'calcpad-block'

    # Replace the block's content with 3 paragraphs (text, display math, text).
    $body = '<w:p><w:r><w:t>Output heading</w:t></w:r></w:p>' +
            '<w:p><m:oMathPara><m:oMath><m:r><m:t>a=1</m:t></m:r></m:oMath></m:oMathPara></w:p>' +
            '<w:p><w:r><w:t>Output last</w:t></w:r></w:p>'
    $cc.Range.InsertXML((Flat $body))
    Step ("after InsertXML 3 paras: cc paragraphs={0}, doc paragraphs={1}, ccs={2}" -f $cc.Range.Paragraphs.Count, $doc.Paragraphs.Count, $doc.ContentControls.Count)
    Step ("  doc text=[{0}]" -f (Show $doc.Content.Text))

    # And back to 1 paragraph.
    $cc.Range.InsertXML((Flat '<w:p><w:r><w:t>single</w:t></w:r></w:p>'))
    Step ("after InsertXML 1 para: cc paragraphs={0}, doc paragraphs={1}" -f $cc.Range.Paragraphs.Count, $doc.Paragraphs.Count)
    Step ("  doc text=[{0}]" -f (Show $doc.Content.Text))

    # Range.Text of a block with several paragraphs, as read back for editing.
    $cc.Range.InsertXML((Flat '<w:p><w:r><w:t>L = 6m</w:t></w:r></w:p><w:p><w:r><w:t>#if L &gt; 5m</w:t></w:r></w:p><w:p><w:r><w:t>#end if</w:t></w:r></w:p>'))
    Step ("source read back: [{0}]" -f (Show $cc.Range.Text))
    $doc.Close([ref]0)
    Step 'Done.'
}
catch { Step ("FAILED: " + $_.Exception.Message + " @ " + $_.InvocationInfo.ScriptLineNumber) }
finally { try { $word.Quit([ref]0) } catch { Step "quit failed: $($_.Exception.Message)" } }
