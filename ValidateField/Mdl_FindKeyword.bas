Attribute VB_Name = "Mdl_FindKeyword"
Option Explicit

Sub •¶ŽšŒŸ’m()
    Dim res As Variant   'Œ‹‰Êyresultz
    Dim TextString As Range
    Dim WordList As Range
    
    Set TextString = Range("A1")  'ŒŸõ‘ÎÛ
    Set WordList = Range("C1:C3")    '’T‚·•¶Žš
    res = FindKeyword(WordList, TextString)

End Sub

Function FindKeyword(WordList As Range, TextString As Range)
    Dim Word As Range
    For Each Word In WordList
        'Žw’è‚µ‚½•¶Žš‚ª‰½•¶Žš–Ú‚É‚ ‚é‚©‚ð•Ô‚·ŠÖ”(ŠJŽnˆÊ’uAŒŸõ‘ÎÛA’T‚·•¶ŽšAvbTextCompare)
        If InStr(1, TextString, Word, 1) > 0 Then
            If IsEmpty(FindKeyword) Then
                FindKeyword = Word
            Else
                FindKeyword = FindKeyword & ", " & Word
            End If
        End If
    Next Word
    
    If FindKeyword = 0 Then
        MsgBox "ŠY“–‚È‚µ", vbExclamation
    Else
        Range("A5") = FindKeyword
    End If
End Function
