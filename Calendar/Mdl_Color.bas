Attribute VB_Name = "Mdl_Color"
Option Explicit
Option Private Module

'カラーコード
Public Const C_CLR_Black As Long = &H0              '黒
Public Const C_CLR_LGray As Long = &HF0F0F0     '薄グレー
Public Const C_CLR_Gray As Long = &HC8C8C8      'グレー
Public Const C_CLR_DGray As Long = &HA0A0A0    '濃グレー
Public Const C_CLR_White As Long = &HFFFFFF     '白
Public Const C_CLR_Hada As Long = &HDAE9F8     '肌色
Public Const C_CLR_LYellow As Long = &HB4FFFF   '薄黄色
Public Const C_CLR_Yellow As Long = &HFFFF         '黄色
Public Const C_CLR_LOrange As Long = &HCCE5FF  '薄橙色
Public Const C_CLR_Orange As Long = &H99E6FF     '橙色
Public Const C_CLR_LRed As Long = &HCEC7FF       '薄赤
Public Const C_CLR_LRed2 As Long = &HB4B4FF     '薄赤2
Public Const C_CLR_Red As Long = &HFF                 '赤
Public Const C_CLR_Magenta As Long = &HFF00FF    'マゼンタ
Public Const C_CLR_LBlue As Long = &HFFE5CC       '薄青
Public Const C_CLR_Blue As Long = &HFF0000         '青
Public Const C_CLR_Cyan As Long = &HFFFF00         'シアン
Public Const C_CLR_LGreen As Long = &HCEEFC6     '薄緑
Public Const C_CLR_Green As Long = &HFF00           '緑

'カラー抽出(抽出対象のセルを選択しイミディエイトで入力)
'? Hex(Selection.Interior.Color)

