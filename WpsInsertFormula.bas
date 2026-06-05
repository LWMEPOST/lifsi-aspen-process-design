Sub InsertFormulaImage()
    Dim imgPath As String
    imgPath = "D:\文档\xwechat_files\wxid_v24r1je3ambj22_958d\temp\RWTemp\2026-05\ebb974efccd942d24d9ca7c27c91f1e8\2e70b065f81c773fb99318116bfec138_final.png"

    Selection.ParagraphFormat.Alignment = 1

    With Selection.InlineShapes.AddPicture(FileName:=imgPath, LinkToFile:=False, SaveWithDocument:=True)
        .Width = 340
        .Height = 160
    End With

    Selection.TypeParagraph
    Selection.TypeText "图 1  反应示意式"
End Sub
