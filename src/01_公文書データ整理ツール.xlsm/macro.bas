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
    End If

End Sub

'STEP2
Sub 出力()

    Dim ox As organizeXML
    Set ox = New organizeXML
    
    rs = MsgBox("データ整理を実行しますか？", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "データ整理を中断しました。"
        Exit Sub
    Else
        Call ox.organizeXML
    End If

End Sub

Sub ファイル出力() 'STEP3

    Dim dc As dataConvert
    Set dc = New dataConvert
    
    rs = MsgBox("CSV取込を実行しますか？", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "CSV取込を中断しました。"
        Exit Sub
    Else
        Call dc.fileOutput
    End If

End Sub

