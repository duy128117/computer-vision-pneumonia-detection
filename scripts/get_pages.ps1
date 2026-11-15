param([string]$DocPath = 'd:\CODE\tgmt\baocao_final_viemphoi.docx')
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($DocPath)
$pages = $doc.ComputeStatistics(2)
Write-Host "$DocPath pages: $pages"
$doc.Close([ref]$false)
$word.Quit()
