$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open('d:\CODE\tgmt\baocao_final_viemphoi.docx')
$paras = $doc.Paragraphs
for ($i = 1; $i -le $paras.Count; $i++) {
    $p = $paras.Item($i)
    $text = $p.Range.Text.Trim()
    if ($text -like "1.*" -or $text -like "2.*" -or $text -like "3.*" -or $text -like "CHƯƠNG*" -or $text -like "KẾT LUẬN*" -or $text -like "TÀI LIỆU*" -or $text -like "PHỤ LỤC*") {
        if (-not ($text -like "*....*")) {
            $page = $p.Range.Information(3)
            Write-Host "Page $page : $text"
        }
    }
}
$doc.Close([ref]$false)
$word.Quit()

