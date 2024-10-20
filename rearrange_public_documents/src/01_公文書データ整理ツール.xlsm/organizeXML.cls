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
    Dim srcPath As Variant                       '検索用パス
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
                data = Array(ws.Cells(i, 7).Value, ws.Cells(i, 8).Value, ws.Cells(i, 9).Value, ws.Cells(i, 10).Value, ws.Cells(i, 11).Value, _
                             ws.Cells(i, 12).Value, ws.Cells(i, 13).Value, ws.Cells(i, 14).Value, ws.Cells(i, 15).Value)
                extractedData.Add data
            End If
        End If
    Next i
    
    'コレクション（XMLの各種情報）が取得完了したらフォルダを作成して、XMLを整理
    getXMLInfo = copyXMLFiles(outputPath, ws, extractedData, docPath, logFile, coll)
End Function

'XMLファイルコピー処理
Function copyXMLFiles(outputPath As String, ws As Worksheet, extractedData As Collection, srcPath As String, logFile As Object, coll As Collection)
    Dim data As Variant
    Dim fso As Object
    Dim srcFolder As Object
    Dim destFolder As Object
    Dim copyFlg As Boolean
    Dim collKey As String
    Dim keyExists As Boolean
    Dim fol As Object
    Dim destPath As String
    Dim customPath As String
    '実行件数確認用
    Dim OKflg As Boolean
    
    '値が異なるフラグ
    Dim isNotSameAppDate As Boolean
    isNotSameAppDate = False
    '未整理フラグ
    Dim isMiseiri As Boolean
    isMiseiri = False
    
    '申請日付：XMLから取得するので、複数取得する可能性あり（STEP1と同じ）
    Dim applicationDate(1) As String
    '事業者名
    Dim companyName As String
    '対象者名
    Dim applicantName As String
    '申請種別
    Dim applicationType As String
    '手続名
    Dim applicationName As String
    '状況
    Dim situation As String
    '審査完了日時
    Dim shinsaKanryoDate As String
    '手続完了日時
    Dim tetsuzukiKanryoDate As String
    'da_comment.txt存在フラグ
    Dim isDaComment As Boolean
    isDaComment = False
    'コピー成功したらTrue
    OKflg = False
    
    'XMLコピー
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set srcFolder = fso.GetFolder(srcPath)
    
    'XMLとtxtから抽出した値を変数に格納
    For i = 1 To extractedData.Count
        data = extractedData(i)

        'data(i)に値が存在していて、かつ変数に値がない場合はdataの値を設定 + 公文書1と2のフォルダで値が異なる場合はエラーとする
        If data(0) <> "" And applicationDate(0) = "" Then
            applicationDate(0) = data(0)
        ElseIf data(0) <> "" And data(0) <> applicationDate(0) Then
            isNotSameAppDate = True
            applicationDate(1) = data(0)
        End If
        If applicationType = "" Then
            applicationType = data(1)
        End If
        If applicationName = "" Then
            applicationName = data(2)
        End If
        If companyName = "" Then
            companyName = data(3)
        End If
        If applicantName = "" Then
            applicantName = data(4)
        End If
        If situation = "" Then
            situation = data(5)
        End If
        If shinsaKanryoDate = "" Then
            shinsaKanryoDate = data(6)
        End If
        If tetsuzukiKanryoDate = "" Then
            tetsuzukiKanryoDate = data(7)
        End If
        If data(8) <> "" Then
            isDaComment = True
        End If
    Next i
    
    'フォルダ作成：申請日_種別手続名_事業者名_対象者名
    If applicantName = "" Then
        customPath = applicationDate(0) & "_" & applicationType & applicationName & "_" & companyName
    Else
        customPath = applicationDate(0) & "_" & applicationType & applicationName & "_" & companyName & "_" & applicantName
    End If
        
    '以下case分岐の処理実施フラグ
    Dim caseFlg As Boolean
    caseFlg = False
    
    'STEP1未整理パターン：各項目が取得できない場合はエラーだが、対象者名なしは除く（STEP1と同じ）/ 対象者名がフェイス / 2つのXMLから取得した申請日付が異なる
    If applicationDate(0) = "" Or companyName = "" Or applicantName = "Faith経営労務事務所" Or isNotSameAppDate = True Then
        destPath = createFolder(outputPath, "未整理", customPath, applicationDate(0)) & "\" & srcFolder.parentfolder.Name '未整理は案件フォルダの作成が必要
        isMiseiri = True
    Else
        'da_comment.txtの有無で分岐
        If Not isDaComment Then
            Select Case situation
            Case "手続終了"
                If applicationName = "算定" Then
                    '「算定」フォルダに整理
                    destPath = createFolder(outputPath, applicationName, customPath, applicationDate(0))
                    caseFlg = True
                ElseIf shinsaKanryoDate = "" Then
                    '「返戻・エラー」フォルダに整理③：状況が「手続終了」なのに、審査完了日時が入っていない
                    destPath = createFolder(outputPath, "返戻・エラー", customPath, applicationDate(0))
                    caseFlg = True
                Else
                    '「手続完了分」フォルダに整理
                    destPath = createFolder(outputPath, "手続完了分", customPath, applicationDate(0))
                    caseFlg = True
                End If
                    
            Case "審査終了"
                If applicationName = "算定" Then
                    '「算定」フォルダに整理
                    destPath = createFolder(outputPath, applicationName, customPath, applicationDate(0))
                    caseFlg = True
                ElseIf tetsuzukiKanryoDate <> "" Then
                    '「返戻・エラー」フォルダに整理④：手続完了日時が入っているのに、状況が「審査終了」
                    destPath = createFolder(outputPath, "返戻・エラー", customPath, applicationDate(0))
                    caseFlg = True
                Else
                    '「審査完了分」フォルダに整理
                    destPath = createFolder(outputPath, "審査完了分", customPath, applicationDate(0))
                    caseFlg = True
                End If
                    
            Case "審査中"
                '「審査中」フォルダに整理
                destPath = createFolder(outputPath, situation, customPath, applicationDate(0))
                caseFlg = True
                    
            Case "送信待機", "到達", "取り下げ済"
                '「その他」フォルダに整理
                destPath = createFolder(outputPath, "その他", customPath, applicationDate(0))
                caseFlg = True
                    
            Case "返戻", "エラー"
                '「返戻・エラー」フォルダに整理①：da_info.txtの「状況」が「返戻」または「エラー」
                destPath = createFolder(outputPath, "返戻・エラー", customPath, applicationDate(0))
                caseFlg = True
            
            End Select
            
        Else
            '「返戻・エラー」フォルダに整理②：da_comment.txtがある
            destPath = createFolder(outputPath, "返戻・エラー", customPath, applicationDate(0))
        End If
        
        '上記どれにも該当しない場合（STEP2未整理パターン）：「未整理」フォルダに整理
        If caseFlg = False And isDaComment = False Then
            destPath = createFolder(outputPath, "未整理", customPath, applicationDate(0)) & "\" & srcFolder.parentfolder.Name '未整理は案件フォルダの作成が必要
            isMiseiriStep2 = True
        End If
    End If
    
    'コピー 上書きする
    srcPath = srcFolder.parentfolder.path
    fso.copyfolder srcPath, destPath, True
    
    'ログ
    If isMiseiri Or isMiseiriStep2 Then
        logFile.writeline "[INFO]   ：コピー元             " & srcPath
        logFile.writeline "[INFO]   ：未整理フォルダに移動 " & destPath
        If applicationDate(0) = "" Then
            logFile.writeline "[INFO]   ：申請日付     [ " & applicationDate(0) & " ]"
        End If
        If isNotSameAppDate Then
            logFile.writeline "[INFO]   ：申請日付     [ " & applicationDate(0) & " ]"
            logFile.writeline "[INFO]   ：申請日付     [ " & applicationDate(1) & " ]"
        End If
        If applicationType = "" Then
            logFile.writeline "[INFO]   ：種別         [ " & applicationType & " ]"
        End If
        If applicationName = "" Then
            logFile.writeline "[INFO]   ：手続名       [ " & applicationName & " ]"
        End If
        If companyName = "" Then
            logFile.writeline "[INFO]   ：事業者名     [ " & companyName & " ]"
        End If
        If applicantName = "Faith経営労務事務所" Then
            logFile.writeline "[INFO]   ：対象者名     [ " & applicantName & " ]"
        End If
        'エラー理由
        If applicationDate(0) = "" Then
            logFile.writeline "[ERROR]  ：申請日付が取得できませんでした。"
        End If
        If isNotSameAppDate Then
            logFile.writeline "[ERROR]  ：同じ申請で申請日付が異なっています。"
        End If
        If applicationType = "" Then
            logFile.writeline "[ERROR]  ：種別が取得できませんでした。"
        End If
        If applicationName = "" Then
            logFile.writeline "[ERROR]  ：手続名が取得できませんでした。"
        End If
        If companyName = "" Then
            logFile.writeline "[ERROR]  ：事業者名が取得できませんでした。"
        End If
        If applicantName = "Faith経営労務事務所" Then
            logFile.writeline "[ERROR]  ：対象者名がFaith経営労務事務所でした。"
        End If
        If isMiseiriStep2 Then
            logFile.writeline "[ERROR]  ：どの条件にも当てはまりませんでした。"
        End If
        logFile.writeline "[INFO]   ：---------------------------------------------------"
        OKflg = False
    Else
        logFile.writeline "[INFO]   ：コピー元   " & srcPath
        logFile.writeline "[INFO]   ：コピー成功 " & destPath
        If applicantName = "" Then
            logFile.writeline "[CAUTION]：対象者名が取得できなかったので、対象者名無しで出力しています。"
        End If
        logFile.writeline "[INFO]   ：---------------------------------------------------"
        OKflg = True
    End If
    
    '結果返却
    copyXMLFiles = OKflg
    
End Function

'引数のパスを組み合わせてフォルダを作成、作成したフォルダパスを返却
Function createFolder(outputPath As String, sortedFolderName As String, subPath As String, applicationDate As String)
    Dim searchPath As String
    Dim currentPath As String
    Dim month As String
    '未整理フラグ
    Dim isMiseiri As Boolean
    isMiseiri = False
    'フォルダ名が重複した場合のsuffix
    Dim suf As Long
    suf = 1
    Dim t As Long
    t = 1
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    '手続完了分・審査完了分・未整理
    If sortedFolderName = "手続完了分" Or sortedFolderName = "審査完了分" Or sortedFolderName = "未整理" Then
        month = Mid(applicationDate, 5, 2)
        If Left(month, 1) = "0" Then
            month = Right(month, 1)
        End If
        If sortedFolderName <> "未整理" Then
            searchPath = outputPath & "\" & sortedFolderName & "\" & Left(applicationDate, 4) & "年\" & month & "月\" & applicationDate & "\" & subPath
        Else
            '未整理で日付あり
            If applicationDate <> "" Then
                searchPath = outputPath & "\" & sortedFolderName & "\" & Left(applicationDate, 4) & "年\" & month & "月\" & applicationDate
            Else
                '未整理で日付なし
                searchPath = outputPath & "\" & sortedFolderName
            End If
            isMiseiri = True
        End If
        
    Else
        '上記以外
        searchPath = outputPath & "\" & sortedFolderName & "\" & subPath
    End If
    
    'フォルダ名重複確認：未整理フォルダは重複確認しない
    If Not isMiseiri Then
        If Dir(searchPath, vbDirectory) <> "" Then
            Do While Dir(searchPath, vbDirectory) <> ""
                searchPath = searchPath & "_" & suf
                suf = suf + 1
            Loop
        End If
    End If
    
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
    createFolder = searchPath
End Function


