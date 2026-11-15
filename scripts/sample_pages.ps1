$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open('d:\CODE\tgmt\Báo cáo cuối kì Thị Giác máy tính.docx')

Write-Host "=== PAGE HEADINGS FOR Báo cáo cuối kì Thị Giác máy tính.docx ==="
$paras = $doc.Paragraphs
for ($i = 1; $i -le $paras.Count; $i++) {
    $p = $paras.Item($i)
    $text = $p.Range.Text.Trim()
    if ($text -like "CHƯƠNG*" -or $text -like "TÀI LIỆU*" -or $text -like "MỤC LỤC*" -or $text -like "BÁO CÁO*") {
        $page = $p.Range.Information(3)
        Write-Host "Page $page : $text"
    }
}
$doc.Close([ref]$false)
$word.Quit()

