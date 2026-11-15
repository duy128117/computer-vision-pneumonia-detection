param([string]$DocPath = 'd:\CODE\tgmt\baocao_final_viemphoi.docx')
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($DocPath)

Write-Host "=== PAGE ANALYSIS FOR $DocPath ==="
$paras = $doc.Paragraphs
for ($i = 1; $i -le $paras.Count; $i++) {
    $p = $paras.Item($i)
    $text = $p.Range.Text.Trim()
    if ($text -like "CHƯƠNG*" -or $text -like "KẾT LUẬN VÀ HƯỚNG*" -or $text -like "TÀI LIỆU THAM KHẢO*" -or $text -like "PHỤ LỤC*") {
        $page = $p.Range.Information(3) # wdActiveEndPageNumber
        Write-Host "Page $page : $text"
    }
}
$totalPages = $doc.ComputeStatistics(2)
Write-Host "Total Document Pages: $totalPages"
$doc.Close([ref]$false)
$word.Quit()

