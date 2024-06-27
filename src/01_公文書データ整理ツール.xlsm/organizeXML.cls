VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "organizeXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub organizeXML()
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim fso As Object
    Set ws = Worksheets("マクロ")
    Set ws2 = Worksheets("フォルダツリー")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim i As Long
    Dim maxRow As Long
    Dim noInfoPath As String
    Dim logFolderPath As String
    Dim logFilePath As String
    Dim logFile As Object
    Dim timestamp As String
    Dim coll As Collection
    Dim cnt As Long
    '出力先
    Dim folderPath As String
    '申請日付
    Dim applicationDate As String
    '会社名
    Dim companyName As String
    '対象者名
    Dim applicntName As String
    '申請種別
    Dim applicationType As String
    'XMLパス
    Dim XMLPath As String
    
    Application.ScreenUpdating = True 'テスト時はTrue、納品時はFalseにする
    
    '取得日時をセット
    ws.Cells(9, 13).Value = Now
    '最下行を取得
    maxRow = ws.Cells(Rows.Count, 5).End(xlUp).row
    '出力先を取得
    folderPath = ws.Cells(2, 16).Value
    'ステータスバー用
    cnt = 1
    
    '未整理フォルダがあれば削除して、その後フォルダ作成
    noInfoPath = folderPath & "\未整理"
    If Dir(noInfoPath, vbDirectory) <> "" Then
        fso.DeleteFolder noInfoPath
    End If
    fso.createfolder noInfoPath
    
    'ログフォルダ作成
    logFolderPath = folderPath & "\log"
    If Dir(logFolderPath, vbDirectory) = "" Then
        fso.createfolder logFolderPath
    End If
    
    'ログファイル作成
    timestamp = format(Now, "yyyymmdd_HHmmss")
    logFilePath = logFolderPath & "\log_" & timestamp & ".txt"
    Set logFile = fso.CreateTextFile(logFilePath, True)
    Set logFile = Nothing
    Set logFile = fso.OpenTextFile(logFilePath, 8)
    logFile.writeline "[INFO]   ：公文書のデータ整理を実行しました。実行日時 " & format(Now, "yyyy/mm/dd HH:mm:ss")
    logFile.writeline "[INFO]   ：XML読込件数 " & ws.Cells(10, 3).Value & "件"
    logFile.writeline "[INFO]   ：---------------------------------------------------"
    
    'ログ出力重複確認用
    Set coll = New Collection
    
    For i = 4 To maxRow
        Call getXMLInfo(ws, i, logFile, coll, cnt)
    Next i
    
    logFile.Close
    Application.StatusBar = False
    Application.ScreenUpdating = True

End Sub

'XMLを格納するフォルダ作成と条件に該当するXMLコピー
Sub getXMLInfo(ws As Worksheet, i As Long, logFile As Object, coll As Collection, cnt As Long)
    Dim maxRow As Long
    Dim regexA As Object
    Dim XMLPath As String
    Dim docPath As String
    Dim srcPath As Variant '検索用パス
    Dim docPaths As Range
    Dim extractedData As Collection
    Dim matches As Object
    Dim match As Object
    Dim data As Variant

    ' XMLパス（E列）の最終行を取得
    maxRow = ws.Cells(Rows.Count, 5).End(xlUp).row

    ' 正規表現オブジェクトの作成
    Set regexA = CreateObject("VBScript.RegExp")
    regexA.Pattern = ".*公文書[12].*.xml"
    regexA.IgnoreCase = True
    
    Set regexB = CreateObject("VBScript.RegExp")
    regexB.Pattern = "(.*\\公文書[12])"
    regexB.IgnoreCase = True
    
    ' パス一覧を取得
    Set docPaths = ws.Range(Cells(4, 5), Cells(maxRow, 5))
    
    '各フォルダの公文書1もしくは公文書2配下のXMLから各種情報を取得する
        XMLPath = ws.Cells(i, 5).Value
        ' 正規表現Aチェック
        If regexA.Test(XMLPath) Then
            Application.StatusBar = "XMLファイルを" & cnt & "件整理中..."
            docPath = ""
            
            '公文書1もしくは公文書2までの文字列を含んだパスを取得
            Set matches = regexB.Execute(XMLPath)
            For Each match In matches
                docPath = match.SubMatches(0)
                Exit For
            Next match
            
            'パス一覧からdocPathが含まれるパスを取得し、コレクションに格納
            t = 4
            Set extractedData = New Collection
            For Each srcPath In docPaths
                If InStr(srcPath, docPath) > 0 And LCase(Right(srcPath, 4)) = ".xml" Then
                    data = Array(ws.Cells(t, 7).Value, ws.Cells(t, 8).Value, ws.Cells(t, 9).Value, ws.Cells(t, 10).Value)
                    extractedData.Add data
                End If
            t = t + 1
            Next srcPath
            
            '取得完了したらフォルダを作成して、XMLを整理
            Call copyXMLFiles(ws, extractedData, docPath, logFile, coll)
            cnt = cnt + 1
        End If

End Sub

'XMLファイルコピー処理
Sub copyXMLFiles(ws As Worksheet, extractedData As Collection, srcPath As String, logFile As Object, coll As Collection)
    Dim outputPath As String
    Dim searchPath As String
    Dim folderName As String
    Dim data As Variant
    Dim fso As Object
    Dim srcFolder As Object
    Dim destFolder As Object
    Dim file As Object
    Dim month As String
    Dim copyFlg As Boolean
    Dim collKey As String
    Dim keyExists As Boolean

    '申請日付
    Dim applicationDate As String
    '会社名
    Dim companyName As String
    '対象者名
    Dim applicntName As String
    '申請種別
    Dim applicationType As String

    ' FileSystemObjectを作成
    Set fso = CreateObject("Scripting.FileSystemObject")
    'マクロシートの出力フォルダパス
    outputPath = ws.Cells(2, 16).Value
    '未整理フォルダパス
    noInfoPath = outputPath & "\未整理"
    'すでにファイル整理済みかどうか
    copyFlg = False
    
    For i = 1 To extractedData.Count
        data = extractedData(i)
        
        'data(i)に値が存在していて、かつ変数に値がない場合はdataの値を設定
        If data(0) <> "" And applicationDate = "" Then
            applicationDate = data(0)
        End If
        
        If data(1) <> "" And companyName = "" Then
            companyName = data(1)
        End If
        
        If data(2) <> "" And applicntName = "" Then
            applicntName = data(2)
        End If
        
        If data(3) <> "" And applicationType = "" Then
            applicationType = data(3)
        End If
    Next i
    
    '4つの変数のうち1つでも値がない場合は、フォルダが作成できないためエラーとする
    If applicationDate = "" Or companyName = "" Or applicntName = "" Or applicationType = "" Then
    
        'XMLコピー
        Set srcFolder = fso.GetFolder(srcPath)
        For Each file In srcFolder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
                If Dir(noInfoPath & "\" & file.Name) = "" Then
                    copyFlg = True
                    logFile.writeline "[INFO]   ：XMLファイル " & file.Path
                    fso.copyfile file.Path, noInfoPath & "\" & file.Name
                    logFile.writeline "[INFO]   ：未整理フォルダに移動 " & noInfoPath & "\" & file.Name
                End If
            End If
        Next file
        
        If copyFlg Then
            'エラー理由
            If applicationDate = "" Then
                logFile.writeline "[ERROR]  ：申請日付の値が取得できませんでした。"
            End If
            
            If companyName = "" Then
                logFile.writeline "[ERROR]  ：会社名の値が取得できませんでした。"
            End If
            
            If applicntName = "" Then
                logFile.writeline "[ERROR]  ：対象者名の値が取得できませんでした。"
            End If
            
            If applicationType = "" Then
                logFile.writeline "[ERROR]  ：申請種別の値が取得できませんでした。"
            End If
            logFile.writeline "[INFO]   ：---------------------------------------------------"
        End If
    Else
        month = Mid(applicationDate, 5, 2)
        If Left(month, 1) = "0" Then
            month = Right(month, 1)
        End If
        
        'パスの存在確認
        searchPath = outputPath & "\" & Left(applicationDate, 4) & "年\" & month & "月\" & applicationDate
        '無ければフォルダ作成する
        If Not fso.FolderExists(searchPath) Then
            ' フォルダパスを分割
            folderParts = Split(searchPath, "\")
            
            ' ルートディレクトリを設定
            currentPath = folderParts(0) & "\"
            
            ' フォルダパスを再構築しながら各フォルダを作成
            For i = 1 To UBound(folderParts)
                currentPath = currentPath & folderParts(i) & "\"
                If Not fso.FolderExists(currentPath) Then
                    fso.createfolder currentPath
                End If
            Next i
        End If
        
        '各種情報を連結し、フォルダ作成
        folderName = companyName & " " & applicntName & " " & applicationType
        destPath = searchPath & "\" & folderName
        If Dir(destPath, vbDirectory) = "" Then
            fso.createfolder destPath
        End If
        
        'XMLコピー
        Set srcFolder = fso.GetFolder(srcPath)
        For Each file In srcFolder.Files
            collKey = destPath & "\" & file.Name
            If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
                'ログ重複確認
                If ExistsItem(coll, collKey) Then
                    copyFlg = True
                    coll.Add collKey
                    logFile.writeline "[INFO]   ：XMLファイル " & file.Path
                    fso.copyfile file.Path, destPath & "\" & file.Name
                    logFile.writeline "[INFO]   ：右記フォルダに移動 " & destPath & "\" & file.Name
                End If
            End If
        Next file
        If copyFlg Then
            logFile.writeline "[INFO]   ：---------------------------------------------------"
        End If
    End If

End Sub

'ログ重複確認処理
Function ExistsItem(objCol As Collection, varItem As Variant) As Boolean
     
    Dim v As Variant
     
    '戻り値の初期値：False
    ExistsItem = True
     
    'Collectionのメンバー数が「0」の場合は戻り値「True」
    If objCol.Count = 0 Then Exit Function
     
    'Collectionの各メンバーと突合
    For Each v In objCol
         
        '突合結果が一致した場合：戻り値「False」にループ抜け（ログ出力しない）
       If v = varItem Then ExistsItem = False: Exit For
         
    Next
     
End Function
