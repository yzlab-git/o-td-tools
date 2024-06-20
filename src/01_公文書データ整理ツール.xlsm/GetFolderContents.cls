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

    Set ws = Worksheets("マクロ")
    Set ws2 = Worksheets("フォルダツリー")
    
    ' フォルダパスAを"マクロ"シートのB2セルから取得
    folderPath = ws.Range("B2").Value
    
    '＜フォルダパスが無い場合はエラー＞
    
    
    
    startRow = 1 ' フォルダツリー出力を開始する行番号
    wsRow = 1 '一覧出力を開始する行番号
    
    Application.ScreenUpdating = True
    ws2.Cells.Clear
    ws2.Cells(startRow, 1).Value = folderPath
    Call ListSubFolders(folderPath, startRow, wsRow, 1, ws, ws2)
    
    ' 取得日時を"マクロ"シートのH1にセット
    ws.Range("H1").Value = Now
    
    Application.ScreenUpdating = True
End Sub

Sub ListSubFolders(folderPath As String, row As Long, wsRow As Long, level As Integer, ws As Worksheet, ws2 As Worksheet)
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
        ws.Cells(wsRow, 8).Value = subFolder.Path ' "マクロ"シートのH列にフォルダパスをセット
        Call ListSubFolders(subFolder.Path, row, wsRow, level + 1, ws, ws2)
    Next subFolder
    
    ' ファイルのリスト
    For Each file In folder.Files
        row = row + 1
        wsRow = wsRow + 1
        ws2.Cells(row, level).Value = file.Name
        ws.Cells(wsRow, 8).Value = file.Path ' "マクロ"シートのH列にファイルパスをセット
    Next file
End Sub
