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
    Set folder = fso.GetFolder(folderPath) 'ここのパスは案件フォルダ名  ex.C:\work\o-td-tools\資料\01_テスト\0604384629
    hasXML = False
    
    'サブフォルダのコレクションを取得　ここは公文書フォルダ
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
    
    '各保険のリストを取得
    Dim socialInsList As Variant
    socialInsList = getIdList(1)
    
    Dim employInsList As Variant
    employInsList = getIdList(2)
    
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
                ws.Cells(wsRow, 5).Value = file.Path  'マクロシートのE列にファイルパスをセット
                If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
                    ' XMLファイルの内容を読み取る
                    Call ReadXMLFile(file.Path, ws, wsRow, socialInsList, employInsList)
                End If
            Next file
        End If
        
    Next i
End Sub

'XMLファイルのタグから情報を取得
Sub ReadXMLFile(filePath As String, ws As Worksheet, rowIndex As Long, socialInsList As Variant, employInsList As Variant)
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument")
    xmlDoc.async = False
    xmlDoc.Load filePath
    
    If xmlDoc.parseError.ErrorCode <> 0 Then
        MsgBox "XMLファイルの読み込みに失敗しました: " & filePath & vbLf & "このファイルの読み込みはスキップします。"
        Exit Sub
    End If
    
    '社会保険用
    Dim applicationDate As String   '申請日
    Dim companyName As String       '事業所名
    Dim applicantName As String     '対象者名
    Dim applicationType As String   '申請種別
    
    applicationDate = ""
    companyName = ""
    applicantName = ""
    applicationType = ""

    
    '雇用保険用
    Dim empInsFlg As Boolean        '雇用保険判定用
    Dim docNo As String             'ドキュメントNo
    Dim affiliateName As String     '事業所名
    Dim employeeName As String      '対象者名
    Dim titleName As String         '申請種別
    
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
    
    '雇用保険判定
    If InStr(docNo, "到達番号：") > 0 And affiliateName <> "" And titleName <> "" And empInsAppDate <> "" Then
        
        ws.Cells(rowIndex, 8).Value = affiliateName
        ws.Cells(rowIndex, 9).Value = employeeName
        ws.Cells(rowIndex, 10).Value = isInIdList(titleName, employInsList)
        
        Dim regex As Object
        Set regex = CreateObject("VBScript.RegExp")
        regex.Global = True
        
        Dim dateType As String          '日付種類（和暦）
        Dim dateYear As Long            '和暦→西暦変換用
        
        '日付を和暦から西暦へ変換
        Set nodes = xmlDoc.SelectSingleNode("//DATE")
        dateType = nodes.Attributes.getNamedItem("TYPE").Text
        
        dateYear = Left(empInsAppDate, 2)
        
        'dateTypeの仕様が定かではないが、以下の通り和暦から西暦へ変換する　昭和から令和まで
        If dateType = 3 Then '昭和
            empInsAppDate = dateYear + 1925 & Right(empInsAppDate, 4)
        ElseIf dateType = 4 Then '平成
            empInsAppDate = dateYear + 1988 & Right(empInsAppDate, 4)
        ElseIf dateType = 5 Then '令和
            empInsAppDate = dateYear + 2018 & Right(empInsAppDate, 4)
        End If
        
        ws.Cells(rowIndex, 7).Value = empInsAppDate
                
        Set nodes = Nothing
        hasInfo = True
    Else
    '社会保険の情報取得
    
        '会社名は<～事業所名称>となっているタグから取得する
        Dim regexA As Object
        Set regexA = CreateObject("VBScript.RegExp")
        regexA.Pattern = ".*事業所名称$"
        regexA.Global = True
        
        '対象者人数
        Dim appCnt As Long
        appCnt = 0
        
        '対象者名は<*被保険者漢字氏名>、<*被保険者氏名x漢字氏名>、<*被用者漢字氏名>となっているタグから取得する
        Dim regexB As Object
        Set regexB = CreateObject("VBScript.RegExp")
        regexB.Pattern = ".*被保険者漢字氏名$"
        regexB.Global = True
        Dim BName As String
        BName = ""
        
        Dim regexC As Object
        Set regexC = CreateObject("VBScript.RegExp")
        regexC.Pattern = ".*被保険者氏名x漢字氏名$"
        regexC.Global = True
        Dim CName As String
        CName = ""
        
        Dim regexD As Object
        Set regexD = CreateObject("VBScript.RegExp")
        regexD.Pattern = ".*被用者漢字氏名$"
        regexD.Global = True
        Dim DName As String
        DName = ""
        
        '申請日は<～提出年月日>となっているタグから取得する
        Dim regexE As Object
        Set regexE = CreateObject("VBScript.RegExp")
        regexE.Pattern = ".*提出年月日$"
        regexE.Global = True
        
        '名前を格納するコレクションを作成
        Dim names As Collection
        Set names = New Collection
        
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
                BName = node.Text
                If BName <> "" Then
                    names.Add BName
                End If
            End If
            If regexC.Test(node.tagName) Then
                CName = node.Text
                If CName <> "" Then
                    names.Add CName
                End If
            End If
            If regexD.Test(node.tagName) Then
                DName = node.Text
                If DName <> "" Then
                    names.Add DName
                End If
            End If
            
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
                Dim appEra As String            '申請日-元号
                Dim appYear As String           '申請日-年（文字列）
                Dim appYearLong As Long         '申請日-年（数値）
                Dim appMonth As String          '申請日-月
                Dim appDay As String            '申請日-日
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
                
                applicationDate = appYear & appMonth & appDay
                
                Set regexF = Nothing
                Set regexG = Nothing
                Set regexH = Nothing
                Set regexI = Nothing
                
            End If
        Next node
        
        Set regexA = Nothing
        Set regexB = Nothing
        Set regexC = Nothing
        Set regexD = Nothing
        Set regexE = Nothing
        
        '対象者0名
        If names.Count = 0 Then
            applicantName = ""
        ElseIf names.Count = 1 Then
        '対象者1名
            applicantName = names(1)
        ElseIf names.Count > 1 Then
        '対象者名2名以上
            applicantName = names.Count & "名分"
        End If
        
        applicationType = xmlDoc.SelectSingleNode("//様式ID").Text
        On Error GoTo 0
        
        If applicationDate <> "" Then
            ws.Cells(rowIndex, 7).Value = applicationDate
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
        '様式IDが取得できていればIDリストと照合
        If applicationType <> "" Then
            ws.Cells(rowIndex, 10).Value = isInIdList(applicationType, socialInsList)
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

'各種保険のリストを取得
Function getIdList(i As Long) As Variant
    '様式IDの追加削除がある場合は「社会保険_様式IDマスタ」シートを修正する
    'TITLEの追加削除がある場合は「雇用保険_申請種別マスタ」シートを修正する
    Dim maxRow As Long
    Dim DataRange As Range
    Dim ws As Worksheet
    
    If i = 1 Then
        Set ws = Worksheets("社会保険_様式IDマスタ")
    ElseIf i = 2 Then
        Set ws = Worksheets("雇用保険_申請種別マスタ")
    End If
    
    '最終行の取得
    maxRow = ws.Cells(Rows.Count, 2).End(xlUp).row
    
    'データ範囲の設定
    Set DataRange = ws.Range(ws.Cells(2, 1), ws.Cells(maxRow, 2))
    
    'データ範囲を配列に格納
    dataArray = DataRange.Value
    
    '二次元配列を返却
    getIdList = dataArray

End Function

'取得した様式IDの値がidListに存在しているかチェック　存在する場合は略称を返却、存在しない場合は略称不明を返却 → 社会保険、雇用保険共通処理
Function isInIdList(checkValue As String, arr As Variant) As String
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If arr(i, 1) = checkValue Then
            isInIdList = arr(i, 2)
            Exit Function
        End If
    Next i
    
    isInIdList = "申請種別 略称不明"  'IDが見つからない場合
End Function

'もしかしたら後で急に使うことになるかもしれないので、残しておく
Sub docNo()
    'DOCNOに「到達番号：」が含まれているので、それを取り除く処理
    regex.Pattern = "(\d+)"
    '数値以外が含まれている場合
    If Not regex.Test(docNo) Then
        '一旦文字列を半角に変換
        applicationDate = StrConv(docNo, vbNarrow)
        ' 数字のみを抽出
        For i = 1 To Len(docNo)
            If Mid(docNo, i, 1) Like "[0-9]" Then
                result = result & Mid(docNo, i, 1)
            End If
        Next i
        docNo = result
    End If
End Sub
