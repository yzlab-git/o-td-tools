Attribute VB_Name = "macro"
Dim im As invoiceManager

'ファイルパス設定
Sub 請求書ファイル選択()

    Set im = New invoiceManager
    rs = MsgBox("請求書ファイルパスを設定しますか？", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "請求書ファイルパスの設定を中断しました。"
        Exit Sub
    Else
        Call im.getFilePath
    End If

End Sub

'ファイルパス設定
Sub メール作成()

    Set im = New invoiceManager
    rs = MsgBox("メールを作成しますか？", vbYesNo)
    
    If rs = vbNo Then
        MsgBox "メール作成を中断しました。"
        Exit Sub
    Else
        Call im.CreateEmailWithPDFInvoices
    End If

End Sub


' テスト用　送付先リストシートのシート名を請求書ファイルのシート名にする
Sub CopyStringsToSheetNames()
    Dim wbA As Workbook
    Dim wbB As Workbook
    Dim wsA As Worksheet
    Dim i As Integer
    Dim sheetName As String
    
    ' 既に開いているAファイルとBファイルを取得
    Set wbA = Workbooks("メール送信_仮マクロ.xlsm") ' Aファイルが開いている場合
    Set wsA = wbA.Sheets("送付先リスト") ' Aファイルの「送付先リスト」シートを指定

    Set wbB = Workbooks("請求書.xlsx") ' Bファイルが開いている場合
    
    ' D2からD101までのループ処理
    For i = 2 To 101
        ' AファイルのD列から文字列を取得
        sheetName = wsA.Cells(i, 4).Value ' D列は4列目に該当
        
            
            ' 新しいシートをBファイルに追加
         wbB.Sheets.Add(After:=wbB.Sheets(wbB.Sheets.Count)).Name = sheetName
        

NextSheet:
    Next i
    
    MsgBox "D2からD101までのシート名がBファイルに追加されました。"
End Sub
