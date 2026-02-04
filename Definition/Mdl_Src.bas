Attribute VB_Name = "Mdl_Src"
Option Explicit

'■モジュール単位のスコープ
Private i As LongLong  'ループ処理
Private cnt As LongLong  'カウント【count】
Private ret As Variant   '戻り値【return value】
Private res As Variant   '結果【result】
Private flag As Boolean  'フラグ【flag】
Private isOK As Boolean  'OK結果
Private hasError As Boolean  'エラー発生
Private fileExists As Boolean    '存在チェック

Sub Src()

    '■プロシージャ単位のスコープ
    Dim row  As LongLong '行
    Dim startRow As LongLong '開始行
    Dim endRow As LongLong '最終行
    Dim col As LongLong '列
    Dim startCol As LongLong '開始列
    Dim endCol As LongLong   '最終列
    Dim maxN As Long
    
    Dim st As T_AppState
    Dim ctx As T_Ctx
    On Error GoTo EH
    
    ctx.StartDt = VBL_Tick()
    ' --- 計測したい処理 ---
    
    VBL_PerfEnter st    '高速開始
    
    ' ...本処理...
    frmProgress.Show vbModeless
    
    maxN = 100
    
    For i = 1 To maxN
    
        Application.Wait Now + TimeValue("0:00:01")
        
        frmProgress.SetProgress i / maxN
    
    Next i
    
    Unload frmProgress
    
    VBL_PerfLeave st    '高速解除
    
    MsgBox "処理時間:" & Format(VBL_Tock(ctx.StartDt), "0.000")
    Exit Sub

    '待機時間設定
    VBL_TimeoutMs = 30000

EH:
    On Error Resume Next
    VBL_PerfLeave st
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
