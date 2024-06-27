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

    Application.StatusBar = False
    Application.ScreenUpdating = True
End Sub

'フォルダとファイルを識別して一覧化
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
            Application.StatusBar = "XMLファイルを" & cnt & "件読取中..."
        End If
    Next file
    
    'XML読込件数
    ListSubFolders = cnt
    
End Function

'XMLファイルのタグから情報を取得
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
    Dim companyName As String
    Dim applicantName As String
    Dim applicationType As String
    
    applicationDate = ""
    companyName = ""
    applicantName = ""
    applicationType = ""
    
    On Error Resume Next
    applicationDate = xmlDoc.SelectSingleNode("//DOCNO").Text
    companyName = xmlDoc.SelectSingleNode("//P1_事業所名称").Text

    '会社名は<～事業所名称>となっているタグから取得する
    Dim regexA As Object
    Set regexA = CreateObject("VBScript.RegExp")
    regexA.Pattern = ".*事業所名称$"
    regexA.Global = True
    
    '対象者名は<～漢字氏名>となっているタグから取得する
    Dim regexB As Object
    Set regexB = CreateObject("VBScript.RegExp")
    regexB.Pattern = ".*漢字氏名$"
    regexB.Global = True
    
    Dim nodes As Object
    Dim node As Object
    
    ' XML内のすべてのノードを取得
    Set nodes = xmlDoc.getElementsByTagName("*")
    
    ' 特定のパターンに一致する値を検索
    For Each node In nodes
        '会社名
        If regexA.Test(node.tagName) Then
            companyName = node.Text
        End If
        
        '対象者名
        If regexB.Test(node.tagName) Then
            applicantName = node.Text
        End If
    Next node
    
    Set regexA = Nothing
    Set regexB = Nothing
    
    applicationType = xmlDoc.SelectSingleNode("//様式ID").Text
    On Error GoTo 0

    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    
    Dim hasInfo As Boolean
    hasInfo = False
    
    ' IDリストを取得
    Dim idList As Variant
    idList = getIdList()
    
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
        result = isInIdList(applicationType, idList)
        If result <> "" Then
            ws.Cells(rowIndex, 10).Value = result
            hasInfo = True
        End If
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

' 二次元配列としてIDリストを取得 (ID, 略称)
Function getIdList() As Variant
    '様式IDの追加削除がある場合はこちらを修正する
    Dim idList As Variant
    idList = Array( _
        Array("495000020328029845", ""), _
        Array("495000012454011861", ""), _
        Array("495000020338029859", ""), _
        Array("495000020332029849", ""), _
        Array("495000020332029850", ""), _
        Array("495000020333029851", ""), _
        Array("495000020334029852", ""), _
        Array("495000020329029846", ""), _
        Array("495000020335029853", ""), _
        Array("495000020335029854", ""), _
        Array("495000020336029855", ""), _
        Array("495000020337029856", ""), _
        Array("495000020337029857", ""), _
        Array("495000020325029841", "雇資格取得"), _
        Array("495000020326029842", "雇資格喪失（離職票なし）"), _
        Array("495000020327029843", "雇資格喪失（離職票あり）"), _
        Array("495000020327029844", ""), _
        Array("495000020339029860", ""), _
        Array("495000020331029848", ""), _
        Array("495000020330029847", ""))
        
    getIdList = idList
End Function

'取得した様式IDの値がidListに存在しているかチェック、存在する場合は略称を返却
Function isInIdList(checkValue As String, arr As Variant) As String
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If arr(i)(0) = checkValue Then
            isInIdList = arr(i)(1)
            Exit Function
        End If
    Next i
    
    isInIdList = "" ' IDが見つからない場合は空文字を返す
End Function
