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
    Dim gfc As GetFolderContents
    Set gfc = New GetFolderContents
    
    Dim ox As organizeXML
    Set ox = New organizeXML
    
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Set ws = Worksheets("マクロ")
    Set ws2 = Worksheets("社会保険_様式IDマスタ")
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '実行件数確認用
    Dim flg As Boolean
    Dim OKcnt As Long
    Dim NGcnt As Long
    Dim NOcnt As Long
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
    
    'フォルダ作成
    Call initializeFolder(outputPath)

    Application.ScreenUpdating = True 'テスト時はTrue、納品時はFalseにする
    
    'ボタン押下により処理開始なので問題ないとは思うが、念のためマクロシートを表示する（マクロシート以外が表示されているとエラーとなってしまう）
    ws.Activate
    
    'パスに問題なければ色塗りつぶしを消す
    ws.Cells(2, 5).Interior.ColorIndex = xlNone
    ws.Cells(3, 5).Interior.ColorIndex = xlNone
    
    '実行日時
    ws.Cells(6, 3).Value = Now
    
    'ログファイル作成
    Dim logFolderPath As String
    Dim logFile As Object
    
    logFolderPath = outputPath & "\log"
    timestamp = format(Now, "yyyymmdd_HHmmss")
    logFilePath = logFolderPath & "\log_" & timestamp & ".txt"
    Set logFile = fso.CreateTextFile(logFilePath, True)
    Set logFile = Nothing
    Set logFile = fso.OpenTextFile(logFilePath, 8)
    logFile.writeline "[INFO]   ：公文書のデータ整理を実行しました。実行日時 " & format(Now, "yyyy/mm/dd HH:mm:ss")
    'ここはXMLではなく、何件フォルダコピーに成功したか、未整理フォルダにコピーしたかを最後にログ出力
    logFile.writeline "[INFO]   ：---------------------------------------------------"
    
    '読取フォルダ配下のフォルダを1申請とみなし、処理を行う
    Set folder = fso.GetFolder(readPath)
    
    'サブフォルダのコレクションを取得
    Dim subFolders As Object
    Set subFolders = folder.subFolders
    
    'サブフォルダの数を取得
    Dim subFolderArray() As Object
    ReDim subFolderArray(1 To folder.subFolders.Count)
    
    'サブフォルダを配列に変換
    w = 1
    For Each subFolder In subFolders
        Set subFolderArray(w) = subFolder
        w = w + 1
    Next subFolder

    For i = 1 To folder.subFolders.Count
    
        '項目名が無ければセット
        Call setItemName(ws)
        
        '一覧表削除
        Call deleteList
        
        '公文書フォルダが3つ以上存在する場合は処理スキップ
        If subFolderArray(i).subFolders.Count > 2 Then
            logFile.writeline "[ERROR]  ：処理フォルダ：" & subFolderArray(i).path
            logFile.writeline "[ERROR]  ：このフォルダには公文書フォルダが3つ以上存在するため、このフォルダの処理をスキップします。"
            logFile.writeline "[INFO]   ：---------------------------------------------------"
            NOcnt = NOcnt + 1
        Else
            'フォルダ読取とXMLタグより情報取得
            Call gfc.getFoldersAndFiles(subFolderArray(i), 4, ws, ws2)
            
            '取得した情報からフォルダ整理
            flg = ox.getXMLInfo(outputPath, ws, logFile, coll)
            
            '成功件数と未整理件数
            If flg Then
                OKcnt = OKcnt + 1
            Else
                NGcnt = NGcnt + 1
            End If
        End If
    Next i
    
    '最後に一覧表削除
    Call deleteList
    
    logFile.writeline "[INFO]   ：成功件数   " & OKcnt & "件"
    logFile.writeline "[INFO]   ：未整理件数 " & NGcnt & "件"
    logFile.writeline "[INFO]   ：未処理件数 " & NOcnt & "件"
    logFile.Close
    
    '件数
    ws.Cells(7, 3).Value = OKcnt
    ws.Cells(8, 3).Value = NGcnt
    ws.Cells(9, 3).Value = NOcnt

    Application.ScreenUpdating = True
End Sub

'読取フォルダから取得したフォルダとファイルのリストを削除
Sub deleteList()
    
    Dim maxRow As Long
    Dim ws As Worksheet
    Set ws = Worksheets("マクロ")
    
    maxRow = ws.Cells(Rows.Count, 5).End(xlUp).row
    If maxRow > 4 Then
    
        ws.Range(Cells(5, 5), Cells(maxRow, 15)).Clear
        
        'もし仮に出力フォルダのパスが削除されてしまったら処理を中断
        If ws.Cells(3, 5).Value = "" Then
            MsgBox "出力フォルダパスを再度入力してください。"
            End
        End If
    End If
End Sub

'操作ミスで出力フォルダパスが削除されないようにVBAで値を設定することにする
Sub setItemName(ws As Worksheet)
    If ws.Cells(4, 5).Value = "" Or ws.Cells(4, 6).Value = "" Or ws.Cells(4, 7).Value = "" Or ws.Cells(4, 8).Value = "" Or ws.Cells(4, 9).Value = "" Or ws.Cells(4, 10).Value = "" _
    Or ws.Cells(4, 11).Value = "" Or ws.Cells(4, 12).Value = "" Or ws.Cells(4, 13).Value = "" Or ws.Cells(4, 14).Value = "" Or ws.Cells(4, 15).Value = "" Then
        
        ws.Cells(4, 5).Value = "フォルダ・ファイル名"
        ws.Cells(4, 6).Value = "情報有無"
        ws.Cells(4, 7).Value = "申請日付"
        ws.Cells(4, 8).Value = "種別"
        ws.Cells(4, 9).Value = "手続名"
        ws.Cells(4, 10).Value = "事業者名"
        ws.Cells(4, 11).Value = "対象者名"
        ws.Cells(4, 12).Value = "状況"
        ws.Cells(4, 13).Value = "審査完了日時"
        ws.Cells(4, 14).Value = "手続完了日時"
        ws.Cells(4, 15).Value = "da_comment.txt"
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
        Cells(9, 3).Value = Null
    End With
End Sub

'ダイアログについて
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

'パスの入力チェック
Sub folderCheck(ws As Worksheet, readPath As String, outputPath As String)
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
    
    'パスが存在しない場合はエラー
    If Dir(readPath, vbDirectory) = "" Then
        MsgBox "読取フォルダが存在しませんでした。" & vbLf & "処理を中断します。"
        End
    ElseIf Dir(outputPath, vbDirectory) = "" Then
        MsgBox "読取フォルダが存在しませんでした。" & vbLf & "処理を中断します。"
        End
    End If
    
    '選択したパスが同じ場合はエラー
    If readPath = outputPath Then
        MsgBox "読取フォルダと出力フォルダのパスが同じです。" & vbLf & "処理を中断します。"
        End
    End If

End Sub

'テキスト読み込み処理開始前のフォルダ初期化と作成メソッド
Sub initializeFolder(outputPath As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(outputPath)
    
    'エラー処理の設定
    On Error GoTo ErrorHandler
    
    '出力フォルダ内にサブフォルダが存在する場合に削除するかどうか確認
    If folder.subFolders.Count > 0 Then
        rs = MsgBox("出力フォルダ内のフォルダを全て消去しますか？", vbYesNo)
            If rs = vbYes Then
                fso.DeleteFolder outputPath & "\*"
            End If
    End If
    
    '手続完了分
    Call createFolderForInitialize(fso, outputPath, "手続完了分")
    
    '審査完了分
    Call createFolderForInitialize(fso, outputPath, "審査完了分")
    
    '返戻・エラー
    Call createFolderForInitialize(fso, outputPath, "返戻・エラー")
    
    '審査中
    Call createFolderForInitialize(fso, outputPath, "審査中")

    'その他
    Call createFolderForInitialize(fso, outputPath, "その他")

    '算定
    Call createFolderForInitialize(fso, outputPath, "算定")
    
    '未整理
    Call createFolderForInitialize(fso, outputPath, "未整理")
    
    'ログ
    Call createFolderForInitialize(fso, outputPath, "log")
    
    Set folder = Nothing
    Set fso = Nothing
    
    Exit Sub
ErrorHandler:
    'エラーが発生した場合の処理
    MsgBox "出力フォルダ内のファイルが開かれています。処理を中断します。", vbYes
    End

End Sub

'フォルダ作成
Sub createFolderForInitialize(fso As Object, path As String, folderName As String)
    
    Dim fullPath As String
    
    'フルパスの存在確認、もし存在すれば削除後にフォルダ作成
    fullPath = path & "\" & folderName
    If Dir(fullPath, vbDirectory) <> "" Then
        fso.DeleteFolder fullPath
    End If
    fso.createFolder fullPath

End Sub
