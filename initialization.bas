Attribute VB_Name = "initialization"
Option Explicit

'------------------------------------------------------------------------------
'■エディター設定
'フォント名:メイリオ（日本語）
'サイズ：9
'インジケーターバー：ON
'標準コード     　：前景(白)　・背景(黒)
'選択された文字  ：前景(自動)・背景(青)
'構文エラー文字  ：前景(赤)　・背景(白)
'次ステートメント：前景(自動)・背景(黄色)
'ブレークポイント：前景(白)　・背景(茶色)
'コメント　　　　：前景(黄緑)・背景(黒)
'キーワード　　　：前景(水色)・背景(黒)
'識別子　　　　　：前景(黄色)・背景(黒)
'ブックマーク　　：前景(ピンク)・背景(黒)・インジケータ(ピンク)
'呼び出し元　　　：前景(赤)・背景(黒)・インジケータ(赤)


Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As LongLong) As Long
Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As LongLong) As Long

'1)基本・制御系
Public i, j, k        As LongLong      'ループ処理
Public idx           As String           '索引【index】
Public cnt           As LongLong      'カウント【count】
Public ret           As Variant          '戻り値【return value】
Public res           As Variant          '結果【result】
Public flag          As Boolean         'フラグ【flag】
Public isOK         As Boolean         'OK結果
Public hasError    　As Boolean         'エラー発生
Public fileExists    As Boolean         '存在チェック

'2)オブジェクト
Public fso          As Object           'オブジェクト
Public obj          As Object           'オブジェクト
Public dict          As Object           'Dictionaryオブジェクト
Public re            As Object           'RegExpオブジェクト
Public wb           As Workbook      'ブック
Public ws           As Worksheet      'シート
Public rng           As Range           'セル
Public tbl            As ListObject      'テーブル操作

'3)データ系
Public arr           As Variant          '配列【array】
Public colArr       　As Variant          '配列【array】
Public result()    　 As Variant          '配列【array】
Public co           　As Collection       '項目(item)とキー(key)をセットで格納するオブジェクト
Public list           As Variant          'リストボックス【ListBox】
Public data         　As Variant          '汎用データ
Public rowData   　　　As Variant          '行データ
Public colData    　　As Variant            '列データ

'4)数値・位置
Public row               As LongLong       '行
Public startRow        As LongLong       '開始行
Public endRow         As LongLong       '最終行
Public col                 As LongLong       '列
Public startCol          As LongLong       '開始列
Public endCol           As LongLong       '最終列
Public pos                As Variant          '位置【position】

'5)文字列・ファイル
Public str                 As String             '文字列
Public path              As String             'ファイルパス
Public fileName        As String             'ファイル名
Public folderPath      As String             'フォルダパス
Public ext                As String             '拡張子

'6)日付・時間
Public dt                 As Date               '日付
Public nowDt           As Date               '今の日時
Public startDt           As Date              '開始日時
Public endDt            As Date               '終了日時
Public elapsed          As Date               '経過日時

'7)処理フロー
Public src                As Variant            'コピー元【source】
Public dst                As Variant            'コピー先【destination】
Public inputVal         As Variant            '入力値
Public outputVal       As Variant            '出力値
Public tmp               As Variant            '一時変数
Public buf                As Variant            'バッファ【buffer】

'8)エラー・ログ
Public errNum          As Variant            'エラー名
Public errMsg           As Variant            'エラーメッセージ
Public log                As Variant            'ログ
Public status           As Variant             '状況：ステータス

'------------------------------------------------------------------------------
'■参照設定
'Microsoft Scripting Runtime
'Microsoft VBScript Regular Expressions X.X
'------------------------------------------------------------------------------
Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Public Function NewRegExp(pattern As String) As Object
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.Global = True
    re.IgnoreCase = True
    Set NewRegExp = re
End Function

Sub library()

path = "\\////"
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.fileExists(path) Then
    Debug.Print "存在します。"
End If

Set fso = Nothing


'------------------------------------------------------------------------------
'■ツール配布リスク
'・参照設定で欠落/バージョン差
'・32/64bit差
'・Officeバージョン差
'・利用禁止のCOM（セキュリティポリシー）

'対策A：参照設定不要の範囲で完結させる
'・Excel標準機能（配列・Dictionaryなしでも書ける範囲）
'・WorksheetFunction、Collection 等

'対策B：CreateObject（遅延バインディング）で吸収する
'Dictionary(Scripting)
Set dict = CreateObject("Scripting.Dictionary")

'高速な「キー→値」検索用のデータ構造
'「この値、もう出てきた？」
'「IDから名前引きたい」
'「重複チェックしたい」
'・重複削除
'・マスタ参照
'・集計
'・グルーピング
'・フラグ管理

dict("A001") = "山田"
dict("A002") = "佐藤"
If dict.Exists("A001") Then
    Debug.Print dict("A001")
End If

'Dictionaryは **0(1)アクセス（ほぼ定数時間）**のため
'大量データになると体感速度が別次元


'RegExp(VBScript)
Set re = CreateObject("VBScript.RegExp")

'正規表現（パターンマッチ）
'・「この文字列、メール形式？」
'・「数字だけ抜きたい」
'・「特定フォーマットだけ抽出」
'10行で100行分の文字列処理を消せる系の道具

re.pattern = "/d+"
re.Global = True

If re.test("ID=12345") Then
    Debug.Print re.Execute("ID=12345")(0)
End If

'------------------------------------------------------------------------------
'■オブジェクト設定
Set wb = ThisWorkbook
Set ws = ThisWorkbook.ActiveSheet

Set ws = Nothing
Set wb = Nothing


'------------------------------------------------------------------------------
'■空の判定
'Null       :  DB由来の欠損値　        ：　If IsNull(x) Then
'Empty    :  未初期化のVariant　    ：　If IsEmpty(x) Then
'Nothing  :  オブジェクトの未生成　：　If x Is Nothing Then
'ブランク  :  ""(空文字)　　　　　　：　If x = "" Then

'Variantが「何も入ってないか？」
If IsNull(tmp) Or IsEmpty(tmp) Or tmp = "" Then
End If

'Rangeセル用
If IsNull(rng.Value) Or rng.Value = "" Then
End If

If obj Is Nothing Then
End If

'------------------------------------------------------------------------------
'■値チェック
tmp = rng.Value

If IsError(tmp) Then
    'エラー
ElseIf IsEmpty(tmp) Or tmp = "" Then
    '空
ElseIf IsDate(tmp) Then
    '日付
ElseIf IsNumeric(tmp) Then
    '数値
ElseIf VarType(tmp) = vbString Then
    '文字列
End If


'------------------------------------------------------------------------------
'■文字列操作
'関数名の後ろに「＄」を付けるとString型を返し型判定の処理が不要となる

str = Replace$("test", "t", "")
str = Left$("090-1234-5678", 3)
str = Mid$("090-1234-5678", 5, 4)
str = Right$("090-1234-5678", 3)

'------------------------------------------------------------------------------
'■配列
arr = rng.Value

For i = LBound(arr, 1) To UBound(arr, 1)
    For j = LBound(arr, 2) To UBound(arr, 2)
        'arr(i, j)を処理
    Next j
Next i

rng.Value = arr

'1次元配列か？2次元か？
If IsArray(arr) Then
End If

'配列サイズ取得
tmp = UBound(arr) - LBound(arr) + 1

'配列初期化
Erase arr

'動的配列拡張
ReDim Preserve arr(1 To newSize)

'列を1次元に
arr = rng.Value
ReDim colArr(1 To UBound(arr, 1))

For i = 1 To UBound(arr, 1)
    colArr(i) = arr(i, 1)
Next i

'配列フィルタ（条件抽出）
For i = 1 To UBound(arr)
    If arr(i) <> "" Then
        cnt = cnt + 1
        ReDim Preserve result(1 To cnt)
        result(cnt) = arr(i)
    End If
Next i

'配列コピー
destArr = srcArr

'配列　→　Join （文字列化）
str = Join(arr, ",")

'Split → 配列
arr = Split(str, ",")

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
Sub WriteArrayToRange(arr As Variant, rng As Range)
    '前処理
    rng.ClearContents
    rng.Resize(UBound(arr, 1), UBound(arr, 2)).Value = arr
    '後処理
End Sub



Sub library2()

'------------------------------------------------------------------------------
'■呼び出し値の処理速度
'「値渡し（ByVal）」 の Long型が最速
'※「参照渡し（ByRef）」のString型が最も遅い

'------------------------------------------------------------------------------
'■イミディエイトウィンドウ
Debug.Print tStart

'------------------------------------------------------------------------------
'■エラー処理
On Error Resume Next        'エラー発生後も次の処理を実行
On Error GoTo 0                 'エラーリセット


'------------------------------------------------------------------------------
'■Excel関数使用
buf1 = Application.WorksheetFunction.VLookup(searchData, _
                                                Range("B1:C2000"), 2, False)    'VLookup
    



'------------------------------------------------------------------------------
'警告を制御「ブックを閉じるときに保存確認を出さない」
Application.DisplayAlerts = False   '警告を非表示
Application.DisplayAlerts = True    '警告を表示




'------------------------------------------------------------------------------
'■高速化開始
Application.ScreenUpdating = False                       '画面描画を一時停止
Application.EnableEvents = False                          'イベント処理を無視する
Application.Calculation = xlCalculationManual       '手動計算に切り替える


'------------------------------------------------------------------------------
'■高速化終了
Application.ScreenUpdating = True                          '画面描画を開始
Application.EnableEvents = True                             'イベント処理を有効にする
Application.Calculation = xlCalculationAutomatic     '自動計算に切り替える
Application.Calculate                                                'Excel関数を再計算



End Sub

