param($inPath, $outPath)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($inPath)
$doc.SaveAs([ref] $outPath, [ref] 17)
$doc.Close()
$word.Quit()
