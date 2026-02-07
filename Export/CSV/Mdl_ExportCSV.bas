Attribute VB_Name = "Mdl_ExportCSV"
Option Explicit


' --- CSV出力コード ---
Public Sub ExportSheetToCsv(ByVal csvPath As String)
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    '実データ範囲を取得（空シート対策）
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, Searchorder:=xlByRows, SearchDirection:=xlPrevious)
    If lastCell Is Nothing Then Exit Sub
    
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), lastCell)
    
    '配列に格納
    Dim data As Variant
    data = rng.Value
    
    '念のため１セル対策
    If Not IsArray(data) Then
        Dim oneCell(1 To 1, 1 To 1) As Variant
        oneCell(1, 1) = data
        data = oneCell
    End If
    
    Dim rowStart As Long, rowEnd As Long
    Dim colStart As Long, colEnd As Long
    rowStart = LBound(data, 1): rowEnd = LBound(data, 1)
    colStart = LBound(data, 2): colEnd = LBound(data, 2)
    
    Dim rowIdx As Long, colIdx As Long
    Dim line As String
    Dim output As String
    
    For rowIdx = rowStart To rowEnd
        line = ""
        For colIdx = colStart To colEnd
              line = line & CsvEscape(data(rowIdx, colIdx))
              If colIdx < colEnd Then line = line & ","
        Next colIdx
        output = output & line & vbCrLf
    Next rowIdx
    
    WriteTextUtf8 csvPath, output   'CSV出力
    
    Set data = Nothing
    
End Sub


' --- CSV用エスケープ処理 ---
Private Function CsvEscape(ByVal v As Variant) As String
    Dim s As String
    
    If IsError(v) Or IsEmpty(v) Then
        CsvEscape = ""
        Exit Function
    End If
    
    s = CStr(v)
    s = Replace(s, """", """""""")  ' " → ""
    
    If InStr(s, ",") > 0 Or InStr(s, vbCr) > 0 Or InStr(s, vbCrLf) > 0 Then
        s = """" & s & """"
    End If
    
    CsvEscape = s
End Function


' --- UTF-8 (BOMなし) でファイル出力 ---
Private Sub WriteTextUtf8(ByVal path As String, ByVal text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    
    stm.Type = 2
    stm.Charset = "UTF-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile path, 2
    stm.Close
End Sub

'実行
Sub Sample_Run()
    ExportSheetToCsv ThisWorkbook.path & "output.csv"
    MsgBox "CSVを出力しました。", vbInformation
End Sub
