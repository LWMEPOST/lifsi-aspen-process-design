$ImagePath = "D:\文档\xwechat_files\wxid_v24r1je3ambj22_958d\temp\RWTemp\2026-05\ebb974efccd942d24d9ca7c27c91f1e8\2e70b065f81c773fb99318116bfec138_final.png"
$OutputPath = "D:\文档\xwechat_files\wxid_v24r1je3ambj22_958d\temp\RWTemp\2026-05\ebb974efccd942d24d9ca7c27c91f1e8\WPS_化学式示例.docx"

if (-not (Test-Path -LiteralPath $ImagePath)) {
    throw "Image not found: $ImagePath"
}

$wps = $null
$doc = $null

try {
    $wps = New-Object -ComObject kwps.Application
    $wps.Visible = $false
    $doc = $wps.Documents.Add()

    $sel = $wps.Selection
    $sel.ParagraphFormat.Alignment = 1

    $shape = $sel.InlineShapes.AddPicture($ImagePath)
    $shape.Width = 340
    $shape.Height = 160

    $sel.TypeParagraph()
    $sel.TypeText("图 1  反应示意式")

    $doc.SaveAs($OutputPath)
    Write-Output "Created: $OutputPath"
}
finally {
    if ($doc -ne $null) {
        try { $doc.Close() } catch {}
    }
    if ($wps -ne $null) {
        try { $wps.Quit() } catch {}
    }
}
