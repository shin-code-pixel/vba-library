Attribute VB_Name = "Mdl01_Def"
Option Explicit
Option Private Module

Public Const C_Ver As String = "1.0.0"

Public Enum E_RunMode
            E_RunMode_Interactive = 1
            E_RunMode_Batch = 2
End Enum

Public Type T_Ctx
    '--- Excel ---
    wb As Workbook
    ws As Worksheet
    Rng As Range
    Tbl As ListObject   'テーブル操作
    
    '--- Object(late binding) ---
    dict As Object  ' Scripting.Dictionary
    Re As Object    ' VBScript.RegExp
    Fso As Object   'Scripting.FileSystemObject
    Obj As Object   '汎用
    
    '--- Data ---
    Arr As Variant
    ColArr As Variant
    ResultArr As Variant
    Co As Collection    '項目(item)とキー(key)をセットで格納するオブジェクト
    ListVal As Variant  'リストボックス【ListBox】
    Data As Variant
    RowData As Variant
    ColData As Variant
    
    '--- Position / Index ---
    Idx As String   '索引【index】
    Pos As Variant  '位置【position】
    
    '--- File / String ---
    StrVal As String
    Path As String
    FileName As String
    folderPath As String
    Ext As String   '拡張子
    
    '--- Date / Time ---
    Dt As Date
    NowDt As Date
    StartDt As Double
    EndDt As Double
    Elapsed As Date '経過日時
    
    '--- Flow ---
    Src As Variant  'コピー元【source】
    Dst As Variant  'コピー先【destination】
    InputVal As Variant
    OutputVal As Variant
    Tmp As Variant
    Buf As Variant  'バッファ【buffer】
    
    '--- Error / Log ---
    ErrNum As Variant
    ErrMsg As Variant
    LogVal As Variant
    StatusVal As Variant
End Type
