VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Sub fatalError(msg As String)
    MsgBox "！！！致命的なエラーが発生したので、マクロの実行を中断します！！！" & vbLf & vbLf & msg
    End
End Sub

'startRowから最終行までを削除
Sub clearRow(ws As Worksheet, ByVal startRow As Long)
    ws.Range(ws.Rows(startRow), ws.Rows(Rows.Count)).ClearContents
End Sub

'startColumnから最終列までを削除
Sub clearColumn(ws As Worksheet, ByVal startCol As Long)
    ws.Range(ws.Columns(Columns.Count), ws.Columns(startCol)).ClearContents
End Sub

Function importCsv(ws As Worksheet, filename As String, countCol As Long) As Long
  ws.Cells.ClearContents
  With ws.QueryTables.Add(Connection:="text;" & filename, Destination:=ws.Range("A1"))
    .TextFilePlatform = 932 'Shift_JIS
    .TextFileCommaDelimiter = True
    .Refresh BackgroundQuery:=False
  End With
  importCsv = ws.Cells(Rows.Count, countCol).End(xlUp).row
End Function

Sub exportCsv(ws As Worksheet, fName As String)
  ws.Copy
  ActiveWorkbook.SaveAs filename:=fName, FileFormat:=xlCSV
  ActiveWorkbook.Close
End Sub

'特定の列の値を文字列から数値に変換する（1行目はタイトル行として無視）
Sub toValue(ws As Worksheet, col As Long)
    lastRow = ws.Cells(Rows.Count, col).End(xlUp).row
    For i = 2 To lastRow
        With ws.Cells(i, col)
            .NumberFormat = ""
            .Value = .Value
        End With
    Next i
End Sub

'特定の列の値を数値から文字列に変換する（1行目はタイトル行として無視）
Sub toText(ws As Worksheet, col As Long)
    lastRow = ws.Cells(Rows.Count, col).End(xlUp).row
    For i = 2 To lastRow
        With ws.Cells(i, col)
            .NumberFormat = "@"
            .Value = .Value
        End With
    Next i
End Sub
