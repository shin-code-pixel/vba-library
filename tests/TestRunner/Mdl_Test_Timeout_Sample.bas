Attribute VB_Name = "Mdl_Test_Timeout_Sample"
Option Explicit

Public Sub Test_Timeout_Sample()
    Dim i As LongLong
    Dim maxN As Long
    maxN = 100
    
    For i = 1 To maxN
        Application.Wait Now + TimeValue("0:00:01")
        frmProgress.SetProgress i / maxN

        VBL_TimeoutMs = 30000    '‘Ò‹@ŽžŠÔÝ’è
    Next i
   
    Mdl_Assert.AssertTrue True, "ƒTƒ“ƒvƒ‹"
End Sub
