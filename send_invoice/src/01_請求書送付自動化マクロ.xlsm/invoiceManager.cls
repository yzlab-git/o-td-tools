VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "invoiceManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub getFilePath()
    
    Dim ws As Worksheet
    Set ws = Worksheets("マクロ")
    Dim TargetFileName As String '選択したファイルパス格納用

    'ファイルダイアログのタイトルを設定
    Application.FileDialog(msoFileDialogFilePicker).Title = "請求書ファイルを選択してください"
    
    'ファイルダイアログを表示し、OKが押された場合
    If Application.FileDialog(msoFileDialogFilePicker).Show = True Then
       '選択したファイルのパスを格納
       TargetFileName = Application.FileDialog(msoFileDialogFilePicker).SelectedItems(1)
       
       'ファイルパスを入力
       ws.Cells(3, 8).Value = TargetFileName
       
       MsgBox "請求書ファイルパスの設定が完了しました。"
    Else
        MsgBox "請求書ファイルパスの設定を中断しました。"
    End If

End Sub

Sub CreateEmailWithPDFInvoices()
    Dim OutlookApp As Object
    Dim MailItem As Object
    Dim ws As Worksheet
    Dim pdfFilePath As String
    Dim i As Integer
    Dim invoiceFilePath As String
    Dim ws2 As Worksheet
    Dim maxRow As Long
    Dim wb As Workbook
    Dim arrSheetNames() As String
    Dim subject As String
    Dim addressRow As Long
    Dim sheetRow As Long
    
    ' シートを指定
    Set ws = Worksheets("マクロ")
    Set ws2 = Worksheets("送付先リスト")
    
    '会社名の最下行を取得
    maxRow = ws2.Cells(2, 1).End(xlDown).Row
    'シート名の最下行を取得
    sheetRow = ws2.Cells(2, 4).End(xlDown).Row
    'アドレスの最下行を取得
     addressRow = ws2.Cells(2, 3).End(xlDown).Row
    '会社名、アドレス、シート名の最下行を比較し、空白があれば処理を終了
    If maxRow <> addressRow Or maxRow <> sheetRow Or addressRow <> sheetRow Then
        MsgBox "送信先リストの会社名、アドレス、シート名のいずれかに不備があります。"
        End
    End If
    
    '請求書ファイルパスを取得
    invoiceFilePath = ws.Cells(3, 8).Value
    If invoiceFilePath = "" Then
        MsgBox "請求書ファイルパスがありません"
        End
    End If
    
    '件名の取得
    subject = ws.Cells(4, 8).Value
    If subject = "" Then
        MsgBox "件名がありません"
        End
    End If
    
    '本文を取得
    sentence = ws.Cells(5, 8).Value
    If sentence = "" Then
        MsgBox "本文がありません"
        End
    End If
    
    'ファイルを開く
    Set wb = Workbooks.Open(invoiceFilePath)
    
    'ワークブック内のシート数を取得
    totalSheets = wb.Sheets.Count
    If sheetRow = totalSheets Then
        MsgBox "送信先リストのシート名の数と請求書ファイルのシート数が一致しません"
        End
    End If
    
    'シート名を格納する配列をシート数に基づいて再定義
    ReDim arrSheetNames(1 To totalSheets)
    
    'シート名を配列に格納
    t = 1
    For Each targetSheet In wb.Sheets
        arrSheetNames(t) = targetSheet.Name
        t = t + 1
    Next targetSheet
    
    
    For u = 1 To sheetRow - 1
    sendListSheetName = ws2.Cells(u + 1, 4).Value
        If arrSheetNames(u) <> sendListSheetName Then
            MsgBox "送信先リストのシート名と請求書ファイルのシート名が一致しません"
            End
        End If
    Next u
    
    ' Outlookアプリケーションのインスタンスを作成または取得
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    
    ' Outlookが既に起動していない場合は新しく起動
    If OutlookApp Is Nothing Then
        Set OutlookApp = CreateObject("Outlook.Application")
    End If
    
    Application.ScreenUpdating = False
    
    '各シートをPDF化してメールに添付
    For i = 2 To maxRow
    
        '新しいメールを作成
        Set MailItem = OutlookApp.CreateItem(0) ' 0 = olMailItem
        
        '会社名
        companyName = ws2.Cells(i, 1).Value
        '担当者名
        toName = ws2.Cells(i, 2).Value
        'アドレス
        toAddress = ws2.Cells(i, 3).Value
        'シート名
        sheetName = ws2.Cells(i, 4).Value
        '本文
        sentence = ws.Cells(5, 8).Value
        
        sentence = Replace(sentence, "<会社名>", companyName)
        If toName <> "" Then
            sentence = Replace(sentence, "<担当者>", toName & "様")
        Else
            sentence = Replace(sentence, "<担当者>", "御中")
        End If
        
        pdfFilePath = convertSheetToPDF(invoiceFilePath, sheetName)
        If pdfFilePath <> "" Then
            MailItem.Attachments.Add pdfFilePath
        End If
        
        ' メールの内容を設定
        With MailItem
            .To = toAddress  ' 宛先
            .subject = subject  ' 件名
            .body = sentence  ' 本文
            .Display ' メールを表示（送信せずに表示する場合）
            ' .Send ' 直接送信する場合はこちらを使用
        End With
        
        'メモリの解放
        Set MailItem = Nothing
    Next i
    
    Application.ScreenUpdating = True
    
    Set OutlookApp = Nothing
End Sub

Function getSheetNames(invoiceFilePath As String) As Variant
    Dim wb As Workbook
    Dim sheetList As New Collection
    Dim ws As Worksheet
    Dim i As Integer
    Dim sheetNames() As String
    
    ' ワークブックを開く
    Set wb = Workbooks.Open(invoiceFilePath)
    
    ' シート名をリストに追加
    For Each ws In wb.Sheets
        sheetList.Add ws.Name
    Next ws
    
    ' シート名の配列を作成
    ReDim sheetNames(1 To sheetList.Count)
    For i = 1 To sheetList.Count
        sheetNames(i) = sheetList(i)
    Next i
    
    ' ワークブックを閉じる
    wb.Close False
    
    ' シート名の配列を返す
    getSheetNames = sheetNames
End Function

Function convertSheetToPDF(invoiceFilePath As String, sheetName As Variant) As String
    Dim wb As Workbook
    Dim pdfFilePath As String
    Dim folderPath As String
    
    ' ワークブックを開く
    Set wb = Workbooks.Open(invoiceFilePath)
    
    ' シートを指定
    Set ws = wb.Sheets(sheetName)
    
    folderPath = GetParentFolderPath(invoiceFilePath)
    
    ' PDFファイルのパスを設定
    pdfFilePath = folderPath & "\" & sheetName & ".pdf"
    
    ' シートをPDFとして保存
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfFilePath, Quality:=xlQualityStandard
    
    ' ワークブックを閉じる
    wb.Close False
    
    ' PDFファイルのパスを返す
    convertSheetToPDF = pdfFilePath
End Function

Function GetParentFolderPath(filePath As String) As String
    ' Dir関数を使用して、ファイル名を取り除き、親フォルダまでのパスを取得
    GetParentFolderPath = Left(filePath, Len(filePath) - Len(Dir(filePath)) - 1)
End Function

