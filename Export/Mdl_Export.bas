Attribute VB_Name = "Mdl_Export"
Option Explicit

'「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」にチェック

Public Sub Export_VBA_All()
    Dim root As String
    root = ThisWorkbook.Path & "/src"
    
    ' 1) 保存先が無いとフォルダが作れない（Pathが空）
    If Len(ThisWorkbook.Path) = 0 Then
        MsgBox "先にブックを保存してください。(ThisWorkbook.Pathが空です)", vbExclamation
        Exit Sub
    End If
    
    Ensurefolder root
    Ensurefolder root & "/bas"
    Ensurefolder root & "/cls"
    Ensurefolder root & "/frm"
    
    Dim vbProj As Object
    Dim comp As Object
    
    ' 2) VBProjectが取れるか（ここが失敗すると後で91になりがち）
    On Error Resume Next
    Set vbProj = ThisWorkbook.VBProject
    On Error GoTo 0
    
    If vbProj Is Nothing Then
        MsgBox _
            "VBProject にアクセスできません。" & vbCrLf & _
            "Excel の設定で「VBAプロジェクトオブジェクトモデルへのアクセスを信頼する」をONにしてください。" & vbCrLf & _
            "(会社PCだとポリシーで不可の場合があります。", vbCritical
        Exit Sub
    End If
    
    Debug.Print vbProj Is Nothing
    
    Dim outPath As String
    
    For Each comp In vbProj.VBComponents
        Select Case comp.Type
            Case 1
                outPath = root & "/bas/" & comp.name & ".bas"
                safeKill outPath
                comp.Export outPath
            
            Case 2
                outPath = root & "/cls/" & comp.name & ".cls"
                safeKill outPath
                comp.Export outPath
            
            Case 3
                outPath = root & "/frm/" & comp.name & ".frm"
                safeKill outPath
                comp.Export outPath
                
                safeKill root & "/frm/" & comp.name & ".frx"
                comp.Export outPath
            
            Case Else
        
        End Select
    Next
    
    MsgBox "エクスポート完了：" & vbCrLf & root, vbInformation
            
End Sub

Private Sub Ensurefolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        MkDir folderPath
    End If
End Sub

Private Sub safeKill(ByVal filePath As String)
    If Len(Dir(filePath)) > 0 Then
        Kill filePath
    End If
End Sub
