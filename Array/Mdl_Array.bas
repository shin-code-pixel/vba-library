Attribute VB_Name = "Mdl10_Array"
Option Explicit

Sub 配列()
Arr = rng.Value

For i = LBound(Arr, 1) To UBound(Arr, 1)
    For j = LBound(Arr, 2) To UBound(Arr, 2)
        'arr(i, j)を処理
    Next j
Next i

rng.Value = Arr

'1次元配列か？2次元か？
If IsArray(Arr) Then
End If

'配列サイズ取得
Tmp = UBound(Arr) - LBound(Arr) + 1

'配列初期化
Erase Arr

'動的配列拡張
ReDim Preserve Arr(1 To newSize)

'列を1次元に
Arr = rng.Value
ReDim ColArr(1 To UBound(Arr, 1))

For i = 1 To UBound(Arr, 1)
    ColArr(i) = Arr(i, 1)
Next i

'配列フィルタ（条件抽出）
For i = 1 To UBound(Arr)
    If Arr(i) <> "" Then
        cnt = cnt + 1
        ReDim Preserve result(1 To cnt)
        result(cnt) = Arr(i)
    End If
Next i

'配列コピー
destArr = srcArr

'配列　→　Join （文字列化）
str = Join(Arr, ",")

'Split → 配列
Arr = Split(str, ",")

End Sub

'拡張用関数
'・結合セル考慮
'・1セルだけの場合は1次元
'・エラー値を除外
'・空行をトリム
Function GetRangeToArray(rng As Range) As Variant
    '前処理
    GetRangeToArray = rng.Value
    '後処理
End Function

'「配列を書き戻す」という操作に名前を付けて、実装詳細と将来の変更点を集約
'配列 arr のサイズに合わせて Range を拡張し、その範囲に一括で書き戻す
'・サイズ調整
'・書き込み前にシートクリア
'・既存データを消すか確認
'・書き込み形式を Value2 にする
'・書き込みログをとる
Sub WriteArrayToRange(Arr As Variant, rng As Range)
    '前処理
    rng.ClearContents
    rng.Resize(UBound(Arr, 1), UBound(Arr, 2)).Value = Arr
    '後処理
End Sub
