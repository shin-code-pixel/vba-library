Attribute VB_Name = "Mdl_Calendar"
Option Explicit

'============================================================
' 月間カレンダーを描画
' - year, month の月間カレンダーを ws 上に出力
' - anchorCell: 左上セル（例: "B2"）
' 戻り値: カレンダー全体範囲（タイトル～日付欄）
' 祝日リストは内閣府サイトでダウンロードしExcelの [Holidays] シート のA列に反映
'============================================================
Public Enum eVBL_WeekStart
        eVBL_WeekStart_Sun = 0      '日曜始まり
        eVBL_WeekStart_Mon = 1      '月曜始まり
End Enum

'祝日リスト(Range)から辞書を作る:Key=CLng(日付)
Public Function VBL_Cal_BuildHolidayDict(ByVal rngDates As Range) As Object
    Dim v As Variant
    Dim r As Variant
    Dim c As Variant
    Dim d As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    If rngDates Is Nothing Then
        Set VBL_Cal_BuildHolidayDict = dict
        Exit Function
    End If
    
    v = rngDates.Value2     '日付がシリアル（数値）になる
    '辞書キーは数値(Long)のため、日付データの場合は「CLng」で変換しないと一致しない
    
    If IsArray(v) Then
        For r = LBound(v, 1) To UBound(v, 1)
            For c = LBound(v, 2) To UBound(v, 2)
                If IsNumeric(v(r, c)) Then
                    dict(CLng(v(r, c))) = True
                ElseIf IsDate(v(r, c)) Then
                    dict(CLng(CDate(v(r, c)))) = True
                End If
            Next c
        Next r
    Else
        '単一セル
         If IsNumeric(v) Then
            dict(CLng(v)) = True
        ElseIf IsDate(v) Then
            dict(CLng(CDate(v))) = True
        End If
    End If
    
    Set VBL_Cal_BuildHolidayDict = dict
End Function

Private Function VBL_Cal_IsHoliday(ByVal d As Date, ByVal holidayDict As Object) As Boolean
    If holidayDict Is Nothing Then
        VBL_Cal_IsHoliday = False
    Else
        VBL_Cal_IsHoliday = holidayDict.Exists(CLng(d))
    End If
End Function

Public Function VBL_Cal_DrawMonth( _
    ByVal ws As Worksheet, _
    ByVal ymYear As Long, _
    ByVal ymMonth As Long, _
    ByVal anchorCell As Range, _
    Optional ByVal weekStart As eVBL_WeekStart = eVBL_WeekStart_Sun, _
    Optional ByVal holidayDict As Object = Nothing) As Range

    Dim firstDate As Date
    Dim lastDate As Date
    Dim firstCellDate As Date
    Dim startDow As Long
    Dim weeks As Long

    Dim r0 As Long, c0 As Long
    Dim r As Long, c As Long
    Dim d As Date
    Dim cur As Date

    If ws Is Nothing Then Err.Raise vbObjectError + 5001, "VBL_Cal_DrawMonth", "ws が Nothing です。"
    If anchorCell Is Nothing Then Err.Raise vbObjectError + 5002, "VBL_Cal_DrawMonth", "anchorCell が Nothing です。"
    If ymYear < 1900 Or ymYear > 9999 Then Err.Raise vbObjectError + 5003, "VBL_Cal_DrawMonth", "year が範囲外です。"
    If ymMonth < 1 Or ymMonth > 12 Then Err.Raise vbObjectError + 5004, "VBL_Cal_DrawMonth", "month が範囲外です。"

    r0 = anchorCell.row
    c0 = anchorCell.Column

    firstDate = DateSerial(ymYear, ymMonth, 1)
    lastDate = DateSerial(ymYear, ymMonth + 1, 0)

    ' startDow: 0..6（週の開始曜日からのオフセット）
    startDow = VBL_Cal_OffsetFromWeekStart(Weekday(firstDate, vbSunday), eVBL_WeekStart_Sun)

    ' カレンダー左上（週開始日のセル）の日付
    firstCellDate = DateAdd("d", -startDow, firstDate)

    weeks = VBL_Cal_WeeksInGrid(firstCellDate, lastDate)

    ' --- 描画（最低限） ---
    ' タイトル
    With ws.Cells(r0, c0).Resize(1, 7)
        .Merge
        .Value = Format$(firstDate, "yyyy-mm") & " Calendar"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
    End With

    ' 曜日行
    VBL_Cal_WriteWeekHeader ws, r0 + 1, c0, eVBL_WeekStart_Sun

    ' 日付欄（weeks 行 x 7列）
    cur = firstCellDate
    For r = 0 To weeks - 1
        For c = 0 To 6
            With ws.Cells(r0 + 2 + r, c0 + c)
                .Value = Day(cur)
                .NumberFormatLocal = "0"
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlTop

                ' 当月以外は薄く
                If month(cur) <> ymMonth Then
                    .Font.ColorIndex = 15 ' 薄いグレー
                Else
                    .Font.ColorIndex = xlAutomatic
                End If

                ' 土日を薄く区別（色は最小限）
                If VBL_Cal_IsWeekend(cur, eVBL_WeekStart_Sun) Then
                    .Font.Bold = True
                    .Font.Color = C_CLR_LRed2
                Else
                    .Font.Bold = False
                End If

                .Borders.LineStyle = xlContinuous
                
                '祝日
                If VBL_Cal_IsHoliday(cur, holidayDict) Then
                    .Font.Bold = True
                    .Font.Color = C_CLR_LRed2
                End If
                
            End With
            cur = DateAdd("d", 1, cur)
        Next c
    Next r

    ' サイズ調整（好みで）
    ws.Columns(c0).Resize(, 7).ColumnWidth = 4.2
    ws.Rows(r0 + 2).Resize(weeks).RowHeight = 18

    Set VBL_Cal_DrawMonth = ws.Range(ws.Cells(r0, c0), ws.Cells(r0 + 1 + weeks, c0 + 6))
End Function

'============================================================
' 実行Sub
'============================================================
Public Sub Demo_Calendar()
    Dim ws As Worksheet
    Dim wsHol As Worksheet
    Dim hol As Object
    
    Set ws = ActiveSheet
    Set wsHol = ThisWorkbook.Worksheets("Holidays")
    
    Set hol = VBL_Cal_BuildHolidayDict(wsHol.Range("A1:A200"))
    
    ws.Cells.Clear
    Call VBL_Cal_DrawMonth(ws, year(Date), month(Date), ws.Range("B2"), eVBL_WeekStart_Sun, hol)
End Sub


'-------------------------
' 内部ヘルパ
'-------------------------

Private Sub VBL_Cal_WriteWeekHeader(ByVal ws As Worksheet, ByVal row As Long, ByVal col As Long, ByVal weekStartMonday As Boolean)
    Dim i As Long
    Dim namesMon() As String
    Dim namesSun() As String

    namesMon = Split("Mon,Tue,Wed,Thu,Fri,Sat,Sun", ",")
    namesSun = Split("Sun,Mon,Tue,Wed,Thu,Fri,Sat", ",")

    For i = 0 To 6
        With ws.Cells(row, col + i)
            .Value = IIf(weekStartMonday, namesMon(i), namesSun(i))
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Font.Color = C_CLR_Black
            .Interior.Color = C_CLR_LBlue
            .Borders.LineStyle = xlContinuous
        End With
    Next i
End Sub

' Weekday(vbSunday) = 1..7 を、週開始（月/日）基準で 0..6 のオフセットへ
Private Function VBL_Cal_OffsetFromWeekStart(ByVal weekdaySunBase As Long, ByVal weekStartMonday As Boolean) As Long
    ' weekdaySunBase: 1=Sun ... 7=Sat
    If weekStartMonday Then
        ' Monを0にする：Mon(2)->0, Tue(3)->1, ... Sun(1)->6
        VBL_Cal_OffsetFromWeekStart = (weekdaySunBase + 5) Mod 7
    Else
        ' Sunを0にする：Sun(1)->0 ... Sat(7)->6
        VBL_Cal_OffsetFromWeekStart = weekdaySunBase - 1
    End If
End Function

Private Function VBL_Cal_WeeksInGrid(ByVal firstCellDate As Date, ByVal lastDate As Date) As Long
    Dim days As Long
    days = DateDiff("d", firstCellDate, lastDate) + 1
    VBL_Cal_WeeksInGrid = (days + 6) \ 7
End Function

Private Function VBL_Cal_IsWeekend(ByVal d As Date, ByVal weekStartMonday As Boolean) As Boolean
    Dim wd As Long
    wd = Weekday(d, vbMonday) ' 1=Mon ... 7=Sun
    ' weekend: Sat(6), Sun(7)
    VBL_Cal_IsWeekend = (wd >= 6)
End Function


