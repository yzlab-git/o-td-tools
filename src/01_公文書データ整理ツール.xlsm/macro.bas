Attribute VB_Name = "macro"
'STEP1
Sub 公文書データ読取()

    Dim gfc As GetFolderContents
    Set gfc = New GetFolderContents
    
    rs = MsgBox("公文書データ読取を実行しますか？", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "公文書データ読取を中断しました。"
        Exit Sub
    Else
        Call gfc.ListFoldersAndFiles
        MsgBox "公文書データ読取が完了しました。"
    End If

End Sub

'STEP2
Sub 公文書データ整理()

    Dim ox As organizeXML
    Set ox = New organizeXML
    
    rs = MsgBox("公文書データ整理を実行しますか？", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "公文書データ整理を中断しました。"
        Exit Sub
    Else
        Call ox.organizeXML
        MsgBox "公文書データ整理が完了しました。"
    End If

End Sub
