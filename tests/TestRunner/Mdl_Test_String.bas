Attribute VB_Name = "Mdl_Test_String"
Option Explicit

Public Sub Test_InStr_Contains()
    Dim s As String: s = "ABCDEF"
    Mdl_Assert.AssertTrue InStr(1, s, "CD") > 0, "CDが含まれているべき"
End Sub

Public Sub Test_InStr_NotContains()
    Dim s As String: s = "ABCDEF"
    Mdl_Assert.AssertTrue InStr(1, s, "ZZ") = 0, "ZZは含まれないべき"
End Sub
