Attribute VB_Name = "Mdl_TestState"
Option Explicit

Public gCaseFailed As Boolean
Public gCaseFailMsg As String

Public Sub VBL_TestCaseReset()
    gCaseFailed = False
    gCaseFailMsg = vbNullString
End Sub

Public Sub VBL_TestFail(ByVal msg As String)
    gCaseFailed = True
    If Len(gCaseFailMsg) > 0 Then gCaseFailMsg = gCaseFailMsg & " | "
    gCaseFailMsg = gCaseFailMsg & msg
End Sub
