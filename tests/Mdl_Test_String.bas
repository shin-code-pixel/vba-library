Attribute VB_Name = "Mdl_Test_String"
Option Explicit

Public Sub Test_InStr_Contains()
    Dim s As String: s = "ABCDEF"
    Mdl_Assert.AssertTrue InStr(1, s, "CD") > 0, "CD‚ªŠÜ‚Ü‚ê‚Ä‚¢‚é‚×‚«"
End Sub

Public Sub Test_InStr_NotContains()
    Dim s As String: s = "ABCDEF"
    Mdl_Assert.AssertTrue InStr(1, s, "ZZ") = 0, "ZZ‚ÍŠÜ‚Ü‚ê‚È‚¢‚×‚«"
End Sub
