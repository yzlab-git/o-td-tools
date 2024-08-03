VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "organizeXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'XMLを格納するフォルダ作成と条件に該当するXMLコピー
Function getXMLInfo(outputPath As String, ws As Worksheet, logFile As Object, coll As Collection)
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
    Set docPaths = ws.Range(Cells(5, 5), Cells(maxRow, 5))
    '各種情報を詰める
    Set extractedData = New Collection
    
    '各フォルダの公文書1もしくは公文書2配下のXMLから各種情報を取得する
    For i = 5 To maxRow
        XMLPath = ws.Cells(i, 5).Value
        ' 正規表現Aチェック
        If regexA.Test(XMLPath) Then
            docPath = ""
            
            '公文書1もしくは公文書2までの文字列を含んだパスを取得
            Set matches = regexB.Execute(XMLPath)
            For Each match In matches
                docPath = match.SubMatches(0)
                Exit For
            Next match
            
            'パス一覧からdocPathが含まれるパスを取得し、コレクションに格納
            If InStr(XMLPath, docPath) > 0 And LCase(Right(XMLPath, 4)) = ".xml" Then
                data = Array(ws.Cells(i, 7).Value, ws.Cells(i, 8).Value, ws.Cells(i, 9).Value, ws.Cells(i, 10).Value)
                extractedData.Add data
            End If
        End If
    Next i
    
    'コレクション（XMLの各種情報）が取得完了したらフォルダを作成して、XMLを整理
    getXMLInfo = copyXMLFiles(outputPath, ws, extractedData, docPath, logFile, coll)
End Function

'XMLファイルコピー処理
Function copyXMLFiles(outputPath As String, ws As Worksheet, extractedData As Collection, srcPath As String, logFile As Object, coll As Collection)
    Dim searchPath As String
    Dim data As Variant
    Dim fso As Object
    Dim srcFolder As Object
    Dim destFolder As Object
    Dim file As Object
    Dim month As String
    Dim copyFlg As Boolean
    Dim collKey As String
    Dim keyExists As Boolean
    Dim fol As Object
    '実行件数確認用
    Dim OKflg As Boolean
    'フォルダ重複（suffix）確認用
    Dim originPath As String
    originPath = ""
    'フォルダ名が重複した場合のsuffix
    Dim suf As Long
    suf = 1
    'ディレクトリにsuffixを付与したフラグ
    Dim sufFlg As Boolean
    sufFlg = False
    
    '申請日付
    Dim applicationDate(1) As String
    '会社名
    Dim companyName(1) As String
    '対象者名
    Dim applicntName(1) As String
    '申請種別
    Dim applicationType(1) As String
    
    '値が異なるフラグ
    Dim isNotSameAppDate As Boolean
    Dim isNotSameCompName As Boolean
    Dim isNotSameApplicant As Boolean
    Dim isNotSameAppType As Boolean
    isNotSameAppDate = False
    isNotSameCompName = False
    isNotSameApplicant = False
    isNotSameAppType = False
    
    '取得した申請種別がNotFoundinIDListのみフラグ
    Dim isIDNotFound As Boolean
    isIDNotFound = False

    Set fso = CreateObject("Scripting.FileSystemObject")
    '未整理フォルダパス
    noInfoPath = outputPath & "\未整理"
    'すでにファイル整理済みかどうか
    copyFlg = False
    'フォルダ作成確認用 → originPathで確認した後、searchPathに代入してフォルダ作成
    searchPath = ""
    month = ""
    'コピー成功したらTrue
    OKflg = False
    
    For i = 1 To extractedData.Count
        data = extractedData(i)
        
        'data(i)に値が存在していて、かつ変数に値がない場合はdataの値を設定 + 公文書1と2のフォルダで値が異なる場合はエラーとする
        If data(0) <> "" And applicationDate(0) = "" Then
            applicationDate(0) = data(0)
        ElseIf data(0) <> "" And data(0) <> applicationDate(0) Then
            isNotSameAppDate = True
            applicationDate(1) = data(0)
        End If
        
        If data(1) <> "" And companyName(0) = "" Then
            companyName(0) = data(1)
        ElseIf data(1) <> "" And data(1) <> companyName(0) Then
            isNotSameCompName = True
            companyName(1) = data(1)
        End If
        
        If data(2) <> "" And applicntName(0) = "" Then
            applicntName(0) = data(2)
        ElseIf data(2) <> "" And data(2) <> applicntName(0) Then
            'isNotSameApplicant = True 　'現状は名前は要検討
            applicntName(1) = data(2)
        End If
        
        If data(3) <> "" And applicationType(0) = "" Then
            applicationType(0) = data(3)
        ElseIf data(3) <> "" And data(3) <> applicationType(0) Then 'applicationType(0)と異なる値であればapplicationType(1)に代入
            isNotSameAppType = True
            applicationType(1) = data(3)
        End If
    Next i
    
    '略称不明 略称不明（様式IDリストに該当しない様式IDのみ）の場合はフラグを立てる
    If applicationType(0) = "申請種別 略称不明" And applicationType(1) = "" Then
        isIDNotFound = True
    End If
    
    'XMLコピー
    Set srcFolder = fso.GetFolder(srcPath)

    '対象者名は除き、3つの変数のうち1つでも値がない場合は、フォルダが作成できないためエラーとする + 公文書1と2のXMLから取得した値が異なる場合もエラー + 申請種別が略称不明のみもエラー
    If applicationDate(0) = "" Or companyName(0) = "" Or applicationType(0) = "" Or isNotSameAppDate Or isNotSameCompName Or isNotSameAppType Or isIDNotFound Then
        
        For Each file In srcFolder.Files
            If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
                
                If applicationDate(0) <> "" Or applicationDate(1) <> "" Then
                    If applicationDate(0) <> "" Then
                        appDate = applicationDate(0)
                        month = Mid(applicationDate(0), 5, 2)
                    Else
                        appDate = applicationDate(1)
                        month = Mid(applicationDate(1), 5, 2)
                    End If
                    If Left(month, 1) = "0" Then
                        month = Right(month, 1)
                    End If
                    
                    'コピー先パス
                    searchPath = noInfoPath & "\" & Left(appDate, 4) & "年\" & month & "月\" & appDate & "\" & srcFolder.parentfolder.Name
                    '無ければ作成する
                    Call createFolder(searchPath)
                Else
                    'コピー先パス（日付が取得できなければ未整理フォルダ直下に移動する）
                    searchPath = noInfoPath & "\" & srcFolder.parentfolder.Name
                End If
                copyFlg = True
                fso.copyfolder file.parentfolder.parentfolder.Path, searchPath, True 'コピー時上書きする
                srcPath = file.parentfolder.parentfolder.Path
            End If
        Next file
        
        If copyFlg Then
            logFile.writeline "[INFO]   ：コピー元             " & srcPath
            logFile.writeline "[INFO]   ：未整理フォルダに移動 " & searchPath
            logFile.writeline "[INFO]   ：申請日付 [ " & applicationDate(0) & " ]"
            If isNotSameAppDate Then
                logFile.writeline "[INFO]   ：申請日付 [ " & applicationDate(1) & " ]"
            End If
            logFile.writeline "[INFO]   ：会社名   [ " & companyName(0) & " ]"
            If isNotSameCompName Then
                logFile.writeline "[INFO]   ：会社名   [ " & companyName(1) & " ]"
            End If
            If Not isIDNotFound Then
                '様式IDリストに存在する様式IDを取得できたが2つのXMLで異なっていた、もしくは様式ID自体が取得できなかったパターン
                logFile.writeline "[INFO]   ：申請種別 [ " & applicationType(0) & " ]"
                If isNotSameAppType Then
                    logFile.writeline "[INFO]   ：申請種別 [ " & applicationType(1) & " ]"
                End If
            Else
                '取得した様式IDが様式IDリストに存在しないものしか取得できなかったパターン
                logFile.writeline "[INFO]   ：申請種別 [ 申請種別 略称不明 ]"
            End If
            
            'エラー理由
            If applicationDate(0) = "" Then
                logFile.writeline "[ERROR]  ：申請日付が取得できませんでした。"
            End If
            If isNotSameAppDate Then
                logFile.writeline "[ERROR]  ：同じ申請で申請日付が異なっています。"
            End If
            
            If companyName(0) = "" Then
                logFile.writeline "[ERROR]  ：会社名が取得できませんでした。"
            End If
            If isNotSameCompName Then
                logFile.writeline "[ERROR]  ：同じ申請で会社名が異なっています。"
            End If
            
            If Not isIDNotFound Then
                If applicationType(0) = "" Then
                    logFile.writeline "[ERROR]  ：申請種別が取得できませんでした。"
                End If
                If isNotSameAppType Then
                    logFile.writeline "[ERROR]  ：同じ申請で申請種別が異なっています。"
                End If
            Else
                logFile.writeline "[ERROR]  ：様式IDマスタシートに存在しない様式IDしか取得できませんでした。"
            End If
            
            logFile.writeline "[INFO]   ：---------------------------------------------------"
        End If
    Else
        month = Mid(applicationDate(0), 5, 2)
        If Left(month, 1) = "0" Then
            month = Right(month, 1)
        End If
        
        '対象者名の有無でフォルダ名分岐
        Dim XMLInfoFolderName As String
        If applicntName(0) = "" Then
            XMLInfoFolderName = companyName(0) & "_" & applicationType(0)
        Else
            XMLInfoFolderName = companyName(0) & "_" & applicntName(0) & "_" & applicationType(0)
        End If
        'コピー先パス
        originPath = outputPath & "\" & Left(applicationDate(0), 4) & "年\" & month & "月\" & applicationDate(0) & "\" & XMLInfoFolderName
        
        'フォルダ名重複確認
        If Dir(originPath, vbDirectory) = "" Then
            searchPath = originPath
        Else
            Do While Dir(searchPath, vbDirectory) <> ""
                searchPath = originPath & "_" & suf
                suf = suf + 1
            Loop
            sufFlg = True
        End If
        'フォルダ作成
        Call createFolder(searchPath)
        
        For Each file In srcFolder.Files
            collKey = srcFolder.parentfolder
            If LCase(fso.GetExtensionName(file.Name)) = "xml" Then
                'ログ重複確認
                If ExistsItem(coll, collKey, sufFlg) Then
                    copyFlg = True
                    coll.Add collKey
                    'コピー
                    fso.copyfolder file.parentfolder.parentfolder.Path, searchPath
                    srcPath = file.parentfolder.parentfolder.Path
                End If
            End If
        Next file
        If copyFlg Then
            logFile.writeline "[INFO]   ：コピー元   " & srcPath
            logFile.writeline "[INFO]   ：コピー成功 " & searchPath
            If applicntName(0) = "" Then
                logFile.writeline "[CAUTION]：XMLから対象者名が取得できなかったので、対象者名無しで出力しています。"
            End If
            logFile.writeline "[INFO]   ：---------------------------------------------------"
            OKflg = True
        End If
    End If
    
    '結果返却
    copyXMLFiles = OKflg
    
End Function

'ログ重複確認処理
Function ExistsItem(objCol As Collection, varItem As Variant, sufFlg As Boolean) As Boolean
     
    Dim v As Variant
     
    '戻り値の初期値：False
    ExistsItem = True
    
    'sufFlgがTrueの場合は戻り値「True」（Collection内で重複してもXMLをコピーするため）
    If sufFlg = True Then Exit Function
     
    'Collectionのメンバー数が「0」の場合は戻り値「True」
    If objCol.Count = 0 Then Exit Function
     
    'Collectionの各メンバーと突合
    For Each v In objCol
         
        '突合結果が一致した場合：戻り値「False」にループ抜け（ログ出力しない）
       If v = varItem Then ExistsItem = False: Exit For
         
    Next
     
End Function

Sub createFolder(searchPath As String)
    
    Dim currentPath As String
    Dim t As Long
    t = 1
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(searchPath) Then
        
        'フォルダパスを分割
        folderParts = Split(searchPath, "\")
    
        'UNCパスの場合（\\～\...）
        If InStr(searchPath, "\\") = 1 Then
            currentPath = "\\" & folderParts(2) & "\"
            t = 3
        Else
            'ルートディレクトリを設定
            currentPath = folderParts(0) & "\"
        End If
        
        ' フォルダパスを再構築しながら各フォルダを作成
        For i = t To UBound(folderParts)
            currentPath = currentPath & folderParts(i) & "\"
            If Not fso.FolderExists(currentPath) Then
                fso.createFolder currentPath
            End If
        Next i
    End If
End Sub
