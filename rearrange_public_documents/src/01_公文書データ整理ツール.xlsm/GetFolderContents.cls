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
Sub getFoldersAndFiles(folderPath As Object, wsRow As Long, ws As Worksheet, ws2 As Worksheet)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim hasXML As Boolean
    Dim var As Variant
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)       'ここのパスは案件フォルダ名  ex.C:\work\o-td-tools\資料\01_テスト\0604384629
    hasXML = False
    
    'サブフォルダのコレクションを取得　ここは公文書フォルダ
    Dim subFolders As Object
    Set subFolders = folder.subFolders
    
    'サブフォルダの数を取得 公文書1と公文書2までの想定
    Dim subFolderArray() As Object
    ReDim subFolderArray(1 To folder.subFolders.Count)
    
    'サブフォルダを配列に変換
    w = 1
    For Each subFolder In subFolders
        Set subFolderArray(w) = subFolder
        w = w + 1
    Next subFolder
    
    For i = 1 To folder.subFolders.Count
    
        'XML存在チェック
        For Each file In subFolderArray(i).Files
            If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
                hasXML = True
                Exit For
            End If
        Next file
        
        If hasXML = True Then
            'ファイルのリスト
            For Each file In subFolderArray(i).Files
                wsRow = wsRow + 1
                ws.Cells(wsRow, 5).Value = file.path 'マクロシートのE列にファイルパスをセット
                If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
                    ' XMLファイルの内容を読み取る
                    Call ReadXMLFile(file.path, ws, wsRow)
                End If
            Next file
        End If
    Next i
    
    'テキストファイルの情報は最終行に値反映
    'da_info.txtの存在チェック：存在しない場合の挙動は現時点で指示なし
    Dim daInfoPath As String
    daInfoPath = folderPath & "\da_info.txt"
    
    If isFileExists(daInfoPath) Then
        'da_info.txtの値取得
        Call ReadDaInfo(daInfoPath, ws, wsRow)
    End If
    
    'da_comment.txtの存在チェック
    Dim daCommentPath As String
    daCommentPath = folderPath & "\da_comment.txt"
    If isFileExists(daCommentPath) Then
        With ws.Cells(wsRow, 15)
            .Value = "○"
            .HorizontalAlignment = xlCenter
        End With
    End If
    
End Sub

'XMLファイルのタグから情報を取得
Sub ReadXMLFile(filePath As String, ws As Worksheet, rowIndex As Long)
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load filePath
    
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "XMLファイルの読み込みに失敗しました: " & filePath & vbLf & "このファイルの読み込みはスキップします。"
        Exit Sub
    End If
    
    '社会保険用
    Dim applicationDate As String                '申請日
    applicationDate = ""
    
    '雇用保険用
    Dim empInsFlg As Boolean                     '雇用保険判定用
    Dim docNo As String                          'ドキュメントNo
    Dim affiliateName As String                  '事業所名
    Dim employeeName As String                   '対象者名
    Dim titleName As String                      '申請種別
    
    empInsFlg = False
    docNo = ""
    affiliateName = ""
    employeeName = ""
    titleName = ""

    On Error Resume Next
    
    '雇用保険のタグ取得
    docNo = xmlDoc.SelectSingleNode("//DOCNO").Text
    affiliateName = xmlDoc.SelectSingleNode("//AFF").Text
    titleName = xmlDoc.SelectSingleNode("//TITLE").Text
    employeeName = xmlDoc.SelectSingleNode("//NAME").Text
    empInsAppDate = xmlDoc.SelectSingleNode("//DATE").Text
    
    '情報有無フラグ
    Dim hasInfo As Boolean
    hasInfo = False
    
    Dim nodes As Object
    Dim node As Object
    
    '雇用保険判定 申請日付のみ取得
    If InStr(docNo, "到達番号：") > 0 And affiliateName <> "" And titleName <> "" And empInsAppDate <> "" Then

        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = True
        
        Dim dateType As String                   '日付種類（和暦）
        Dim dateYear As Long                     '和暦→西暦変換用
        
        '日付を和暦から西暦へ変換
        Set nodes = xmlDoc.SelectSingleNode("//DATE")
        dateType = nodes.Attributes.getNamedItem("TYPE").Text
        
        dateYear = Left(empInsAppDate, 2)
        
        'dateTypeの仕様が定かではないが、以下の通り和暦から西暦へ変換する　昭和から令和まで
        If dateType = 3 Then                     '昭和
            empInsAppDate = dateYear + 1925 & Right(empInsAppDate, 4)
        ElseIf dateType = 4 Then                 '平成
            empInsAppDate = dateYear + 1988 & Right(empInsAppDate, 4)
        ElseIf dateType = 5 Then                 '令和
            empInsAppDate = dateYear + 2018 & Right(empInsAppDate, 4)
        End If
        
        ws.Cells(rowIndex, 7).Value = empInsAppDate

    Else
        '社会保険の情報取得
        
        '申請日は<～提出年月日>となっているタグから取得する
        Dim regexE As Object
        Set regexE = CreateObject("VBScript.RegExp")
        regexE.Pattern = ".*提出年月日$"
        regexE.Global = True
        
        ' XML内のすべてのノードを取得
        Set nodes = xmlDoc.getElementsByTagName("*")
        
        ' 特定のパターンに一致する値を検索
        For Each node In nodes
            
            '申請日
            If regexE.Test(node.tagName) Then
            
                Dim regexF As Object
                Set regexF = CreateObject("VBScript.RegExp")
                regexF.Pattern = ".*元号$"
                regexF.Global = True
                
                Dim regexG As Object
                Set regexG = CreateObject("VBScript.RegExp")
                regexG.Pattern = ".*年$"
                regexG.Global = True
                
                Dim regexH As Object
                Set regexH = CreateObject("VBScript.RegExp")
                regexH.Pattern = ".*月$"
                regexH.Global = True
                
                Dim regexI As Object
                Set regexI = CreateObject("VBScript.RegExp")
                regexI.Pattern = ".*日$"
                regexI.Global = True
                
                '<～提出年月日>のタグの中の年月日タグを取得する
                Dim childNode As Object
                Dim appEra As String             '申請日-元号
                Dim appYear As String            '申請日-年（文字列）
                Dim appYearLong As Long          '申請日-年（数値）
                Dim appMonth As String           '申請日-月
                Dim appDay As String             '申請日-日
                appEra = ""
                appYear = ""
                appMonth = ""
                appDay = ""
                
                For Each childNode In node.ChildNodes
                    If regexF.Test(childNode.tagName) Then
                        appEra = childNode.Text
                    End If
                
                    If regexG.Test(childNode.tagName) Then
                        appYear = childNode.Text
                    End If
                    
                    If regexH.Test(childNode.tagName) Then
                        appMonth = childNode.Text
                    End If
                    
                    If regexI.Test(childNode.tagName) Then
                        appDay = childNode.Text
                    End If
                Next childNode
                
                '和暦から西暦変換
                appYearLong = CInt(appYear)
                If appEra = "昭和" Then
                    appYear = appYearLong + 1925
                ElseIf appEra = "平成" Then
                    appYear = appYearLong + 1988
                ElseIf appEra = "" Or appEra = "令和" Then
                    appYear = appYearLong + 2018
                End If
                
                If Len(appMonth) = 1 Then
                    appMonth = "0" & appMonth
                End If
                
                If Len(appDay) = 1 Then
                    appDay = "0" & appDay
                End If
                
                applicationDate = appYear & appMonth & appDay
                
            End If
        Next node
        
        Set regexE = Nothing
        
        applicationType = xmlDoc.SelectSingleNode("//様式ID").Text
        On Error GoTo 0
        
        If applicationDate <> "" Then
            ws.Cells(rowIndex, 7).Value = applicationDate
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

'テキストファイルの存在確認
Function isFileExists(filePath As String) As Boolean
    If Dir(filePath) <> "" Then
        isFileExists = True
    Else
        isFileExists = False
    End If
End Function

'da_info.txtから値取得
Sub ReadDaInfo(txtPath As String, ws As Worksheet, rowIndex As Long)
    Dim FileContent As String
    Dim TextLine As String
    Dim InfoDict As Object
    Dim Key As String, Value As String
    Dim fileNum As Integer
    Dim Lines() As String
    Dim i As Integer
    Dim Stream As Object
    
    On Error Resume Next
    
    ' Dictionaryオブジェクトを作成
    Set InfoDict = CreateObject("Scripting.Dictionary")
    
    ' ADODB.Streamオブジェクトを作成
    Set Stream = CreateObject("ADODB.Stream")
    
    ' 文字列モードでストリームを開く
    Stream.Type = 2                              ' 文字列モード
    Stream.Charset = "UTF-16"                    ' UTF-16として読み込む
    Stream.Open
    Stream.LoadFromFile txtPath
    
    ' ファイル内容を一括で読み込む
    FileContent = Stream.ReadText
    
    ' ストリームを閉じる
    Stream.Close
    On Error GoTo 0
    
    If FileContent <> "" Then
        ' 改行で分割して行ごとに処理
        Lines = Split(FileContent, vbCrLf)
        
        For i = LBound(Lines) To UBound(Lines)
            TextLine = Lines(i)
            ' ":"（全角コロン）を基準に分割して項目名と値を取得
            If InStr(TextLine, "：") > 0 Then
                Key = Trim(Split(TextLine, "：")(0)) ' 項目名を取得
                Value = Trim(Split(TextLine, "：")(1)) ' 値を取得
                InfoDict(Key) = Value            ' Dictionaryに格納
            End If
        Next i
        
        ' 項目をマクロシートへ転記
        ws.Cells(rowIndex, 8).Value = InfoDict("種別")
        ws.Cells(rowIndex, 9).Value = InfoDict("手続名")
        ws.Cells(rowIndex, 10).Value = InfoDict("会社名")
        ws.Cells(rowIndex, 11).Value = InfoDict("内容")
        ws.Cells(rowIndex, 12).Value = InfoDict("状況")
        ws.Cells(rowIndex, 13).Value = InfoDict("審査完了")
        ws.Cells(rowIndex, 14).Value = InfoDict("手続完了")
        
        '内容（対象者名）が「Faith経営労務事務所」かチェック
        If checkFaith(InfoDict("内容")) Then
            ws.Cells(rowIndex, 11).Value = "Faith経営労務事務所"
        End If
    End If
End Sub

'対象者名が「Faith経営労務事務所」かどうか確認メソッド
Function checkFaith(nameStr As String)
    
    '空白を削除
    nameStr = Replace(Trim(nameStr), " ", "")
    nameStr = Replace(nameStr, "　", "")
    
    '大文字に変換
    nameStr = UCase(nameStr)
    
    '全角に変換
    nameStr = StrConv(nameStr, vbWide)
    
    If InStr(nameStr, "ＦＡＩＴＨ経営労務事務所") > 0 Then
        checkFaith = True
    Else
        checkFaith = False
    End If
    
End Function

