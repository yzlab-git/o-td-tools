VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetFolderContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub ListFoldersAndFiles()
    Dim folderPath As String
    Dim startRow As Long
    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim wsRow As Long
    Dim rng As Range
    Dim XMLcnt As Long

    Set ws = Worksheets("マクロ")
    Set ws2 = Worksheets("フォルダツリー")
    
    ' フォルダパスAを"マクロ"シートのB2セルから取得
    folderPath = ws.Cells(2, 5).Value
    
    If folderPath = "" Then
        MsgBox "読取したいフォルダのパスを入力してください。" & vbLf & "処理を中断します。"
        ws.Cells(2, 5).Interior.ColorIndex = 6 'フォルダパスのセルを黄色ぬりつぶし
        End
    End If
    
    startRow = 1 ' フォルダツリー出力を開始する行番号
    wsRow = 3  '一覧出力を開始する行番号
    XMLcnt = 0 'XML読込件数
    
    Application.ScreenUpdating = True 'テスト時はTrue、納品時はFalseにする
    ws.Cells(2, 5).Interior.ColorIndex = xlNone
    ws.Range(Cells(4, 5), Cells(Rows.Count, 10)).Clear
    ws2.Cells.Clear
    ws2.Cells(startRow, 1).Value = folderPath
    XMLcnt = ListSubFolders(folderPath, startRow, wsRow, 1, ws, ws2, 0)
    
    
    ' 取得日時をセット
    ws.Cells(9, 3).Value = Now
    'XML読込件数をセット
    ws.Cells(10, 3).Value = XMLcnt
    '罫線と列幅調整
    Set rng = ws.Range(Cells(3, 5), Cells(ws.Cells(Rows.Count, 5).End(xlUp).row, 10))
    Call setBorder(rng)
    rng.Columns.AutoFit
    
    Application.ScreenUpdating = True
End Sub

Function ListSubFolders(folderPath As String, row As Long, wsRow As Long, level As Integer, ws As Worksheet, ws2 As Worksheet, cnt As Long)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim indent As String
    Dim startRow As Long
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    
    startRow = row
    
    ' フォルダのリスト
    For Each subFolder In folder.SubFolders
        row = row + 1
        wsRow = wsRow + 1
        ws2.Cells(row, level).Value = subFolder.Name
        ws.Cells(wsRow, 5).Value = subFolder.Path ' "マクロ"シートのH列にフォルダパスをセット
        Call ListSubFolders(subFolder.Path, row, wsRow, level + 1, ws, ws2, cnt)
    Next subFolder
    
    ' ファイルのリスト
    For Each file In folder.Files
        row = row + 1
        wsRow = wsRow + 1
        ws2.Cells(row, level).Value = file.Name
        ws.Cells(wsRow, 5).Value = file.Path  ' "マクロ"シートのH列にファイルパスをセット
        
        If LCase(Right(file.Name, 4)) = ".xml" Then
            ' XMLファイルの内容を読み取る
            Call ReadXMLFile(file.Path, ws, wsRow)
            cnt = cnt + 1
        End If
    Next file
    
    'XML読込件数
    ListSubFolders = cnt
    
End Function

Sub ReadXMLFile(filePath As String, ws As Worksheet, rowIndex As Long)
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load filePath
    
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "XMLファイルの読み込みに失敗しました: " & filePath
        Exit Sub
    End If
    
    Dim applicationDate As String
    Dim applicationDate_year As String
    Dim applicationDate_month As String
    Dim applicationDate_day As String
    Dim companyName As String
    Dim applicantName As String
    Dim applicationType As String
    
    applicationDate = ""
    companyName = ""
    applicantName = ""
    applicationType = ""
    
    On Error Resume Next
    applicationDate = xmlDoc.SelectSingleNode("//DOCNO").Text
    applicationDate_year = xmlDoc.SelectSingleNode("//P1_年").Text
    applicationDate_month = xmlDoc.SelectSingleNode("//P1_月").Text
    applicationDate_day = xmlDoc.SelectSingleNode("//P1_日").Text
    companyName = xmlDoc.SelectSingleNode("//P1_事業所名称").Text
    applicantName = xmlDoc.SelectSingleNode("//P1_被保険者氏名x漢字氏名").Text
    On Error GoTo 0
    
    ' 申請種別（届書コード）取得のため正規表現オブジェクトの作成
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "N\d{7}"
    regex.Global = True
    
    Dim nodes As Object
    Dim node As Object
    Dim matches As Object
    Dim match As Object
    
    ' XML内のすべてのノードを取得
    Set nodes = xmlDoc.getElementsByTagName("*")
    
    ' 特定のパターンに一致する値を検索
    For Each node In nodes
        If regex.Test(node.tagName) Then
            Set matches = regex.Execute(node.tagName)
            If matches.Count > 0 Then
                applicationType = matches(0).Value
            End If
        End If
    Next node
    
    Dim hasInfo As Boolean
    hasInfo = False
    
    If applicationDate <> "" Then
        regex.Pattern = "(\d+)"
        '数値以外が含まれている場合
        If Not regex.Test(applicationDate) Then
            '一旦文字列を半角に変換
            applicationDate = StrConv(applicationDate, vbNarrow)
            ' 数字のみを抽出
            For i = 1 To Len(applicationDate)
                If Mid(applicationDate, i, 1) Like "[0-9]" Then
                    result = result & Mid(applicationDate, i, 1)
                End If
            Next i
            applicationDate = result
        End If
        
        'DOCNOの左から8桁は日付
        ws.Cells(rowIndex, 7).Value = "'" & Left(applicationDate, 8)
        hasInfo = True
    End If
    If companyName <> "" Then
        ws.Cells(rowIndex, 8).Value = companyName
        hasInfo = True
    End If
    If applicantName <> "" Then
        ws.Cells(rowIndex, 9).Value = applicantName
        hasInfo = True
    End If
    If applicationType <> "" Then
        ws.Cells(rowIndex, 10).Value = applicationType
        hasInfo = True
    End If
    
    If hasInfo = True Then
        With ws.Cells(rowIndex, 6)
            .Value = "○"
            .HorizontalAlignment = xlCenter
        End With
    End If
End Sub

'罫線を設定する
Sub setBorder(rng As Range)
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 0 '黒色
    End With
End Sub

