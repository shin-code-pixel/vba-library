VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgress 
   Caption         =   "処理中"
   ClientHeight    =   850
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6690
   OleObjectBlob   =   "frmProgress.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub SetProgress(ByVal percent As Double)
    lblBar.Width = lblBack.Width * percent
    DoEvents
End Sub
