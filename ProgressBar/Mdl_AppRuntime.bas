Attribute VB_Name = "Mdl_AppRuntime"
Option Explicit
Option Private Module

Private mTimeoutMs As Long  'ミリ秒単位のタイムアウト時間

' 高速化
Public Type T_AppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    DisplayAlerts As Boolean
    Calculation As XlCalculation
End Type

Public Sub VBL_PerfEnter(ByRef st As T_AppState)
'■高速化開始
    With Application
            st.ScreenUpdating = .ScreenUpdating
            st.EnableEvents = .EnableEvents
            st.DisplayAlerts = .DisplayAlerts
            st.Calculation = .Calculation

            .ScreenUpdating = False                       '画面描画を一時停止
            .EnableEvents = False                          'イベント処理を無視する
            .DisplayAlerts = False                           '警告を非表示
            .Calculation = xlCalculationManual        '手動計算に切り替える
     End With
End Sub
    
Public Sub VBL_PerfLeave(ByRef st As T_AppState)
'■高速化終了
    With Application
            .ScreenUpdating = st.ScreenUpdating
            .EnableEvents = st.EnableEvents
            .DisplayAlerts = st.DisplayAlerts
            .Calculation = st.Calculation
     End With
End Sub

' タイムアウト
Public Property Get VBL_TimeoutMs() As Long
    VBL_TimeoutMs = mTimeoutMs
End Property

Public Property Let VBL_TimeoutMs(ByVal v As Long)
    '10分まで待機可能。それ以降はエラー
    If v < 0 Or v > 600000 Then
        Err.Raise vbObjectError + 1001, "VBL_TimeoutMs", _
                        "Timeout は　0～600000 の範囲で指定してください。"
    End If
    mTimeoutMs = v
End Property

' 処理時間計測
Public Function VBL_Tick() As Double
        VBL_Tick = Timer
End Function

Public Function VBL_Tock(ByVal t0 As Double) As Double
        Dim ti As Double
        ti = Timer
        
        If ti < t0 Then
            '日付またぎ(Timerは24時間でリセットされるためマイナス防止)
            VBL_Tock = (86400 - t0) + ti
        Else
            VBL_Tock = ti - t0
        End If
End Function





