Attribute VB_Name = "Mdl_ProgressBar"
Option Explicit
Sub ProgressBarTask()
    '■プロシージャ単位のスコープ
    Dim i As LongLong
    Dim maxN As Long
    Dim st As T_AppState
    Dim StartDt As Double
    
    On Error GoTo EH
    StartDt = VBL_Tick()    '処理計測
    
    VBL_PerfEnter st    '高速開始
    
    frmProgress.Show vbModeless
    maxN = 10
    For i = 1 To maxN
        Application.Wait Now + TimeValue("0:00:01")
        frmProgress.SetProgress i / maxN
    Next i
    
    Unload frmProgress
    
    VBL_PerfLeave st    '高速解除
    
    MsgBox "処理時間:" & Format(VBL_Tock(StartDt), "0.000")
    Exit Sub

EH:
    On Error Resume Next
    VBL_PerfLeave st
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

