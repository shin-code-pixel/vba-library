Attribute VB_Name = "Mdl_Assert"
Option Explicit

Public Sub AssertTrue(ByVal condition As Boolean, Optional ByVal message As String = "")
    If Not condition Then
        Mdl_TestState.VBL_TestFail NzMsg(message, "Expected True, but was False.")
    End If
End Sub

Public Sub AssertFalse(ByVal condition As Boolean, Optional ByVal message As String = "")
    If condition Then
        Mdl_TestState.VBL_TestFail NzMsg(message, "Expected False, but was True.")
    End If
End Sub

Public Sub AssertEquals(ByVal expected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
    If Not VariantEquals(expected, actual) Then
        Mdl_TestState.VBL_TestFail NzMsg(message, _
            "Not equal. expected=[" & ToDebugText(expected) & "] actual=[" & ToDebugText(actual) & "]")
    End If
End Sub

Public Sub AssertNotEquals(ByVal notExpected As Variant, ByVal actual As Variant, Optional ByVal message As String = "")
    If VariantEquals(notExpected, actual) Then
        Mdl_TestState.VBL_TestFail NzMsg(message, "Should not equal. value=[" & ToDebugText(actual) & "]")
    End If
End Sub

Public Sub AssertNotNothing(ByVal obj As Object, Optional ByVal message As String = "")
    If obj Is Nothing Then
        Mdl_TestState.VBL_TestFail NzMsg(message, "Object is Nothing.")
    End If
End Sub

Public Sub AssertContains(ByVal haystack As String, ByVal needle As String, Optional ByVal message As String = "")
    If InStr(1, haystack, needle, vbBinaryCompare) = 0 Then
        Mdl_TestState.VBL_TestFail NzMsg(message, _
            "Not contains. needle=[" & needle & "] haystack=[" & haystack & "]")
    End If
End Sub

' --- 以下は今のあなたのままでOK（内部ユーティリティ） ---
Private Function VariantEquals(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsNull(a) And IsNull(b) Then VariantEquals = True: Exit Function
    If IsEmpty(a) And IsEmpty(b) Then VariantEquals = True: Exit Function
    If IsDate(a) And IsDate(b) Then VariantEquals = (CDbl(CDate(a)) = CDbl(CDate(b))): Exit Function
    If IsNumeric(a) And IsNumeric(b) Then VariantEquals = (CDbl(a) = CDbl(b)): Exit Function
    VariantEquals = (CStr(a) = CStr(b))
End Function

Private Function ToDebugText(ByVal v As Variant) As String
    If IsNull(v) Then
        ToDebugText = "<Null>"
    ElseIf IsEmpty(v) Then
        ToDebugText = "<Empty>"
    Else
        ToDebugText = CStr(v)
    End If
End Function

Private Function NzMsg(ByVal msg As String, ByVal fallback As String) As String
    If Len(Trim$(msg)) = 0 Then
        NzMsg = fallback
    Else
        NzMsg = msg
    End If
End Function

