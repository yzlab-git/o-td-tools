VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StartCategorize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub startCategorize()
    
    '元々ボタンを2つに分けていたが1つに集約することにしたので
    'このメソッドで各クラスモジュールから呼び出し
    Dim gfc As GetFolderContents
    Set gfc = New GetFolderContents
    
    Dim ox As organizeXML
    Set ox = New organizeXML
    
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Set ws = Worksheets("マクロ")
    Set ws2 = Worksheets("様式IDマスタ")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '実行件数確認用
    Dim flg As Boolean
    Dim OKcnt As Long
    Dim NGcnt As Long
    flg = False
    
    'ログ出力重複確認用
    Dim coll As Collection
    Set coll = New Collection
    
    '読取フォルダパス
    Dim readPath As String
    readPath = ws.Cells(2, 5).Value
    
    '出力フォルダパス
    Dim outputPath As String
    outputPath = ws.Cells(3, 5).Value
    
    'エラーチェック
    Call folderCheck(ws, readPath, outputPath)
    
    Application.ScreenUpdating = True 'テスト時はTrue、納品時はFalseにする
    
    'パスに問題なければ色塗りつぶしを消す
    ws.Cells(2, 5).Interior.ColorIndex = xlNone
    ws.Cells(3, 5).Interior.ColorIndex = xlNone
    
    '未整理フォルダがあれば削除して、その後フォルダ作成
    noInfoPath = outputPath & "\未整理"
    If Dir(noInfoPath, vbDirectory) <> "" Then
        fso.DeleteFolder noInfoPath
    End If
    fso.createFolder noInfoPath
    
    'ログフォルダ作成
    logFolderPath = outputPath & "\log"
    If Dir(logFolderPath, vbDirectory) = "" Then
        fso.createFolder logFolderPath
    End If
    
    '実行日時
    ws.Cells(6, 3).Value = Now
    
    'ログファイル作成
    Dim logFile As Object
    timestamp = format(Now, "yyyymmdd_HHmmss")
    logFilePath = logFolderPath & "\log_" & timestamp & ".txt"
    Set logFile = fso.CreateTextFile(logFilePath, True)
    Set logFile = Nothing
    Set logFile = fso.OpenTextFile(logFilePath, 8)
    logFile.writeline "[INFO]   ：公文書のデータ整理を実行しました。実行日時 " & format(Now, "yyyy/mm/dd HH:mm:ss")
    'ここはXMLではなく、何件フォルダコピーに成功したか、未整理フォルダにコピーしたかを最後にログ出力
    logFile.writeline "[INFO]   ：---------------------------------------------------"
    
    '読取フォルダ配下の親フォルダの数（読取フォルダ配下にはXMLファイルは存在しない想定）
    Dim folColl As Collection
    Set folColl = New Collection
    Call getFolColl(folColl, readPath)
    
    For i = 1 To folColl.Count
    
        '項目名が無ければセット
        Call setItemName(ws)
        
        '一覧表削除
        Call deleteList
        
        'フォルダ読取とXMLタグより情報取得
        Call gfc.getFoldersAndFiles(folColl(i), 4, ws, ws2)
        
        '取得した情報からフォルダ整理
        flg = ox.getXMLInfo(outputPath, ws, logFile, coll)
        
        '成功件数と未整理件数
        If flg Then
            OKcnt = OKcnt + 1
        Else
            NGcnt = NGcnt + 1
        End If
    
    Next i
    
    '最後に一覧表削除
    Call deleteList
    
    
    logFile.writeline "[INFO]   ：成功件数   " & OKcnt & "件"
    logFile.writeline "[INFO]   ：未整理件数 " & NGcnt & "件"
    logFile.Close
    
    '件数
    ws.Cells(7, 3).Value = OKcnt
    ws.Cells(8, 3).Value = NGcnt

    Application.ScreenUpdating = True
End Sub

'読取フォルダから取得したフォルダとファイルのリストを削除
Sub deleteList()
    
    Dim maxRow As Long
    Dim ws As Worksheet
    Set ws = Worksheets("マクロ")
    
    maxRow = ws.Cells(Rows.Count, 5).End(xlUp).row
    If maxRow > 4 Then
    
        ws.Range(Cells(5, 5), Cells(maxRow, 10)).Clear
        
        'もし仮に出力フォルダのパスが削除されてしまったら処理を中断
        If ws.Cells(3, 5).Value = "" Then
            MsgBox "出力フォルダパスを再度入力してください。"
            End
        End If
    End If
End Sub

'操作ミスで出力フォルダパスが削除されないようにVBAで値を設定することにする
Sub setItemName(ws As Worksheet)
    If ws.Cells(4, 5).Value = "" Or ws.Cells(4, 6).Value = "" Or ws.Cells(4, 7).Value = "" Or ws.Cells(4, 8).Value = "" Or ws.Cells(4, 9).Value = "" Or ws.Cells(4, 10).Value = "" Then
        ws.Cells(4, 5).Value = "フォルダ・ファイル名"
        ws.Cells(4, 6).Value = "情報有無"
        ws.Cells(4, 7).Value = "申請日付"
        ws.Cells(4, 8).Value = "会社名"
        ws.Cells(4, 9).Value = "対象者名"
        ws.Cells(4, 10).Value = "申請種別"
    End If
End Sub

Sub getFolColl(folColl As Collection, readPath As String)

    Dim folder As Object
    Dim fol As Object
    Dim fso As Object
    Dim isIn As Boolean
    Dim isParent As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(readPath)
    isParent = False
    
    If folder.SubFolders.Count > 0 Then
        For Each fol In folder.SubFolders
            '再帰的呼び出し
            Call getFolColl(folColl, fol.Path)
        Next fol
        isParent = True
    End If
    
    isIn = False
    For i = 1 To folColl.Count
        If folder.Path = folColl(i) Then
            isIn = True
            Exit For
        End If
    Next i

    If Not isIn And Not isParent Then
        folColl.Add folder.Path
    End If

End Sub

Sub deleteValue()
    
    Dim ws As Worksheet
    Set ws = Worksheets("マクロ")
    
    With ws
        Cells(2, 5).Value = Null
        Cells(3, 5).Value = Null
        Cells(6, 3).Value = Null
        Cells(7, 3).Value = Null
        Cells(8, 3).Value = Null
    End With

End Sub

Sub getFolderPath(i As Long)
    
    Dim ws As Worksheet
    Set ws = Worksheets("マクロ")
    Dim TargetFldName As String '選択したフォルダパス格納用

    'ダイアログフォルダ表示
    If i = 1 Then
        titleStr = "読取"
        row = 2
    Else
        titleStr = "出力"
        row = 3
    End If
    
    Application.FileDialog(msoFileDialogFolderPicker).Title = titleStr & "するフォルダを選択してください"
    If Application.FileDialog(msoFileDialogFolderPicker).Show = True Then 'フォルダを選んでダイアログOK押下
       TargetFldName = Application.FileDialog(msoFileDialogFolderPicker).SelectedItems(1) '選択したフォルダのパスを格納
       ws.Cells(row, 5).Value = TargetFldName 'アクティブセルにフォルダパスを入力
       MsgBox titleStr & "フォルダ設定が完了しました。"
    Else
        MsgBox titleStr & "フォルダ設定を中断しました。"
    End If

End Sub

Sub folderCheck(ws As Worksheet, readPath As String, outputPath As String)

    '選択したパスが同じ場合はエラー
    If readPath = outputPath Then
        MsgBox "読取フォルダと出力フォルダのパスが同じです。" & vbLf & "処理を中断します。"
        End
    End If
    
    'パスが存在しない場合はエラー
    If Dir(readPath, vbDirectory) = "" Then
        MsgBox "読取フォルダが存在しませんでした。" & vbLf & "処理を中断します。"
        End
    ElseIf Dir(outputPath, vbDirectory) = "" Then
        MsgBox "読取フォルダが存在しませんでした。" & vbLf & "処理を中断します。"
        End
    End If
    
    '未入力はエラー
    If readPath = "" Then
        MsgBox "読取したいフォルダのパスを入力してください。" & vbLf & "処理を中断します。"
        ws.Cells(2, 5).Interior.ColorIndex = 6 'フォルダパスのセルを黄色ぬりつぶし
        End
    ElseIf outputPath = "" Then
        MsgBox "出力したいフォルダのパスを入力してください。" & vbLf & "処理を中断します。"
        ws.Cells(3, 5).Interior.ColorIndex = 6 'フォルダパスのセルを黄色ぬりつぶし
        End
    End If

End Sub
