Attribute VB_Name = "Mdl_VBL_Timesheet"
Option Explicit

'=============================
' 勤怠表を作成（テンプレ）
'=============================
Public Sub VBL_TS_CreateMonthlySheet( _
    ByVal ymYear As Long, _
    ByVal ymMonth As Long, _
    Optional ByVal weekStartMonday As Boolean = False, _
    Optional ByVal holidayDict As Object = Nothing _
)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim sheetName As String
    Dim firstDate As Date, lastDate As Date
    Dim d As Date
    Dim r As Long
    Dim title As String

    Set wb = ThisWorkbook

    If ymYear < 1900 Or ymYear > 9999 Then Err.Raise vbObjectError + 6101, "VBL_TS_CreateMonthlySheet", "ymYear が範囲外です。"
    If ymMonth < 1 Or ymMonth > 12 Then Err.Raise vbObjectError + 6102, "VBL_TS_CreateMonthlySheet", "ymMonth が範囲外です。"

    firstDate = DateSerial(ymYear, ymMonth, 1)
    lastDate = DateSerial(ymYear, ymMonth + 1, 0)

    sheetName = "Timesheet_" & Format$(firstDate, "yyyymm")
    Set ws = VBL_TS_GetOrCreateSheet(wb, sheetName)
    ws.Cells.Clear

    ' ---- タイトル ----
    title = "勤怠表 " & Format$(firstDate, "yyyy年m月")
    With ws.Range("A1:I1")
        .Merge
        .Value = title
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' ---- ヘッダ ----
    ' A:日付 B:曜 C:区分 D:開始 E:終了 F:休憩(h) G:労働(h) H:残業(h) I:備考
    ws.Range("A3").Value = "日付"
    ws.Range("B3").Value = "曜"
    ws.Range("C3").Value = "区分"
    ws.Range("D3").Value = "開始"
    ws.Range("E3").Value = "終了"
    ws.Range("F3").Value = "休憩(h)"
    ws.Range("G3").Value = "労働(h)"
    ws.Range("H3").Value = "残業(h)"
    ws.Range("I3").Value = "備考"

    With ws.Range("A3:I3")
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Interior.Color = &HE0E0E0 ' 薄グレー（必要なら色定数へ）
    End With

    ' ---- 本体 ----
    r = 4
    d = firstDate
    Do While d <= lastDate
        ws.Cells(r, "A").Value = d
        ws.Cells(r, "A").NumberFormatLocal = "yyyy/mm/dd"

        ws.Cells(r, "B").Value = VBL_TS_WeekName(d, weekStartMonday)
        ws.Cells(r, "B").HorizontalAlignment = xlCenter

        ' 区分：平日=出勤、土日祝=休日（初期値）
        ws.Cells(r, "C").Value = IIf(VBL_TS_IsWeekend(d), "休日", "出勤")
        If VBL_TS_IsHoliday(d, holidayDict) Then ws.Cells(r, "C").Value = "祝日"

        ' 入力欄の書式
        ws.Cells(r, "D").NumberFormatLocal = "hh:mm"
        ws.Cells(r, "E").NumberFormatLocal = "hh:mm"
        ws.Cells(r, "F").NumberFormatLocal = "0.0"

        ' 計算式
        ' 労働(h) = (終了-開始)*24 - 休憩(h)（開始/終了が未入力なら空）
        ws.Cells(r, "G").Formula = _
            "=IF(OR(D" & r & "="""",E" & r & "=""""),"""",(E" & r & "-D" & r & ")*24-F" & r & ")"
        ws.Cells(r, "G").NumberFormatLocal = "0.00"

        ' 残業(h) = MAX(労働-8, 0)
        ws.Cells(r, "H").Formula = "=IF(G" & r & "="""","""",MAX(G" & r & "-8,0))"
        ws.Cells(r, "H").NumberFormatLocal = "0.00"

        ' 罫線
        With ws.Range("A" & r & ":I" & r)
            .Borders.LineStyle = xlContinuous
            .VerticalAlignment = xlCenter
        End With

        ' 土日祝の背景色（必要なら Mdl_VBL_Color の定数へ置換）
        If VBL_TS_IsHoliday(d, holidayDict) Then
            ws.Range("A" & r & ":I" & r).Interior.Color = &HCEC7FF ' 祝日
        ElseIf VBL_TS_IsWeekend(d) Then
            ws.Range("A" & r & ":I" & r).Interior.Color = &HCCE5FF ' 週末
        End If

        r = r + 1
        d = DateAdd("d", 1, d)
    Loop

    ' ---- 合計行 ----
    ws.Cells(r + 1, "F").Value = "合計"
    ws.Cells(r + 1, "G").Formula = "=SUM(G4:G" & (r - 1) & ")"
    ws.Cells(r + 1, "H").Formula = "=SUM(H4:H" & (r - 1) & ")"
    ws.Range("F" & (r + 1) & ":I" & (r + 1)).Font.Bold = True
    ws.Range("F" & (r + 1) & ":I" & (r + 1)).Borders.LineStyle = xlContinuous

    ' ---- 見た目調整 ----
    ws.Columns("A").ColumnWidth = 12
    ws.Columns("B").ColumnWidth = 4
    ws.Columns("C").ColumnWidth = 6
    ws.Columns("D").ColumnWidth = 7
    ws.Columns("E").ColumnWidth = 7
    ws.Columns("F").ColumnWidth = 7
    ws.Columns("G").ColumnWidth = 7
    ws.Columns("H").ColumnWidth = 7
    ws.Columns("I").ColumnWidth = 28

    ws.Range("A4:A" & (r - 1)).HorizontalAlignment = xlCenter
    ws.Range("C4:C" & (r - 1)).HorizontalAlignment = xlCenter

    ws.Range("A3:I3").AutoFilter

    ' 入力しやすいように開始/終了/休憩/備考だけ白、それ以外薄く
    ws.Range("D4:F" & (r - 1)).Interior.Color = vbWhite
    ws.Range("I4:I" & (r - 1)).Interior.Color = vbWhite

End Sub

'=============================
' ヘルパ
'=============================
Private Function VBL_TS_GetOrCreateSheet(ByVal wb As Workbook, ByVal name As String) As Worksheet
    On Error Resume Next
    Set VBL_TS_GetOrCreateSheet = wb.Worksheets(name)
    On Error GoTo 0
    If VBL_TS_GetOrCreateSheet Is Nothing Then
        Set VBL_TS_GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        VBL_TS_GetOrCreateSheet.Name = name
    End If
End Function

Private Function VBL_TS_IsWeekend(ByVal d As Date) As Boolean
    Dim wd As Long
    wd = Weekday(d, vbMonday) ' 1=Mon ... 7=Sun
    VBL_TS_IsWeekend = (wd >= 6)
End Function

Private Function VBL_TS_IsHoliday(ByVal d As Date, ByVal holidayDict As Object) As Boolean
    If holidayDict Is Nothing Then
        VBL_TS_IsHoliday = False
    Else
        VBL_TS_IsHoliday = holidayDict.Exists(CLng(d))
    End If
End Function

Private Function VBL_TS_WeekName(ByVal d As Date, ByVal weekStartMonday As Boolean) As String
    ' 表示は日本語1文字
    ' weekStartMonday は表示には不要ですが、引数として保持しておくと拡張しやすいです
    Select Case Weekday(d, vbSunday) ' 1=日..7=土
        Case 1: VBL_TS_WeekName = "日"
        Case 2: VBL_TS_WeekName = "月"
        Case 3: VBL_TS_WeekName = "火"
        Case 4: VBL_TS_WeekName = "水"
        Case 5: VBL_TS_WeekName = "木"
        Case 6: VBL_TS_WeekName = "金"
        Case 7: VBL_TS_WeekName = "土"
    End Select
End Function
