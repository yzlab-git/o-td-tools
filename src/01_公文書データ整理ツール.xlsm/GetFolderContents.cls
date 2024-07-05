VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GetFolderContents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'XMLファイルを含むフォルダの中身を識別して、XMLファイル読取処理を行う
Sub getFoldersAndFiles(folderPath As String, wsRow As Long, ws As Worksheet, ws2 As Worksheet)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim hasXML As Boolean
    Dim var As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    hasXML = False
    
    'XML存在チェック
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
            hasXML = True
            Exit For
        End If
    Next file
    
    If hasXML = False Then
        'フォルダのリスト
        For Each subFolder In folder.SubFolders
            wsRow = wsRow + 1
            ws.Cells(wsRow, 5).Value = subFolder.Path 'マクロシートのE列にフォルダパスをセット
            Call getFoldersAndFiles(subFolder.Path, wsRow, ws, ws2)
            Exit Sub
        Next subFolder
    Else
        'ファイルのリスト
        For Each file In folder.Files
            wsRow = wsRow + 1
            ws.Cells(wsRow, 5).Value = file.Path  'マクロシートのE列にファイルパスをセット
            If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
                ' XMLファイルの内容を読み取る
                Call ReadXMLFile(file.Path, ws, wsRow)
            End If
        Next file
    End If
End Sub

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

' 様式IDマスタシートから二次元配列として様式IDリストを取得 (ID, 略称)
Function getIdList() As Variant
    '様式IDの追加削除がある場合は「様式IDマスタ」シートを修正する
    Dim maxRow As Long
    Dim DataRange As Range
    Dim ws As Worksheet
    Set ws = Worksheets("様式IDマスタ")
    
    '最終行の取得
    maxRow = ws.Cells(Rows.Count, 2).End(xlUp).row
    
    'データ範囲の設定
    Set DataRange = ws.Range(ws.Cells(2, 2), ws.Cells(maxRow, 3))
    
    'データ範囲を配列に格納
    dataArray = DataRange.Value
    
    '二次元配列を返却
    getIdList = dataArray

End Function

'取得した様式IDの値がidListに存在しているかチェック、存在する場合は略称を返却
Function isInIdList(checkValue As String, arr As Variant) As String
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = checkValue Then
            isInIdList = arr(i, 2)
            Exit Function
        End If
    Next i
    
    isInIdList = "" ' IDが見つからない場合は空文字を返す
End Function
