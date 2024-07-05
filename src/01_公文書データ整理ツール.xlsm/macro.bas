Attribute VB_Name = "macro"
Sub 公文書データ整理()

    Dim sc As startCategorize
    Set sc = New startCategorize
    
    rs = MsgBox("公文書データ整理を実行しますか？", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "公文書データ整理を中断しました。"
        Exit Sub
    Else
        Call sc.startCategorize
        MsgBox "公文書データ整理が完了しました。"
    End If

End Sub

Sub 初期化()

    Dim sc As startCategorize
    Set sc = New startCategorize
    
    rs = MsgBox("初期化を実行しますか？" & vbLf & vbLf & "※以下の情報を削除します。" & vbLf _
    & "・読取フォルダ" & vbLf & "・出力フォルダ" & vbLf & "・実行日時" & vbLf & "・成功件数" & vbLf & "・未整理件数", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "初期化を中断しました。"
        Exit Sub
    Else
        Call sc.deleteValue
        MsgBox "初期化が完了しました。"
    End If

End Sub
