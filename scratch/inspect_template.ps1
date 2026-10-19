Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.Xml.Linq
$path = Join-Path $env:APPDATA 'WordRegulator\Templates\3. rapor.docx'
$zip = [System.IO.Compression.ZipFile]::OpenRead($path)
$entry = $zip.GetEntry('word/document.xml')
$stream = $entry.Open()
$reader = [System.IO.StreamReader]::new($stream)
$xml = $reader.ReadToEnd()
$reader.Dispose()
$zip.Dispose()

$xDoc = [System.Xml.Linq.XDocument]::Parse($xml)
$w = [System.Xml.Linq.XNamespace]"http://schemas.openxmlformats.org/wordprocessingml/2006/main"

$tblIdx = 0
foreach ($tbl in $xDoc.Descendants($w + "tbl")) {
    $tblText = ($tbl.Descendants($w + "t") | ForEach-Object { $_.Value }) -join ""
    if ($tblText.Contains("SERIM_1") -or $tblText.Contains("SS_1_") -or $tblText.Contains("POISSON") -or $tblText.Contains("SHEAR") -or $tblText.Contains("BULK") -or $tblText.Contains("ELASTICITY") -or $tblText.Contains("SEİSMİC")) {
        Write-Output "=== TABLE $tblIdx ==="
        $rIdx = 0
        foreach ($tr in $tbl.Elements($w + "tr")) {
            $cIdx = 0
            $cellsInfo = @()
            foreach ($tc in $tr.Elements($w + "tc")) {
                $tcPr = $tc.Element($w + "tcPr")
                $gridSpan = ""
                if ($tcPr -ne $null) {
                    $gsEl = $tcPr.Element($w + "gridSpan")
                    if ($gsEl -ne $null) {
                        $gridSpan = $gsEl.Attribute($w + "val").Value
                    }
                }
                $cText = ($tc.Descendants($w + "t") | ForEach-Object { $_.Value }) -join ""
                $cellsInfo += "C$cIdx(span=$gridSpan):'$cText'"
                $cIdx++
            }
            Write-Output ("R" + $rIdx + ": " + ($cellsInfo -join " | "))
            $rIdx++
        }
    }
    $tblIdx++
}
