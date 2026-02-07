Attribute VB_Name = "Mdl00_Initialization"
Option Explicit
Option Private Module

'■Public変数の懸念点
'状態(State)がどこからでも書き換え可能になり、原因追跡が難しくなる
'初期化順序(Auto_Open / Workbook_Open 等)や、エラー中断で半端な値が残る
'将来の改修でいつの間にか依存箇所が増える（スパゲッティ化）
'(アドイン/参照される形だと）外部プロジェクトから触れる可能性が出る

'1) 定数・型・列挙は Public 推奨

'2) 可変の値は「Private + Property」で一本化 推奨

'------------------------------------------------------------------------------

'■Property
'1. 設定値（Config）：環境依存・運用で変わり得るもの
'2. 状態（State）　  ：処理モードや実行中フラグなど、壊れる厄介なもの
'3. 入口（Facade)  　: 外部から触っていい窓口を１つに絞る

'Public 可変変数を大量に置く代わりに、「入口だけ Public」「中身はPrivate」に寄せる

'------------------------------------------------------------------------------

'■Const (定数)
' ・単独の不変値（意味がブレない）
'       (例:タイムアウト秒、固定パス名、固定文字列、シート名、列番号、正規表現パターン等）
' ・ビットフラグ（&H01 など）を組み合わせたい場合

'■Enum (列挙)
' ・ 番号のみ設定可（文字列を設定する場合はDictionaryを使用）
' ・「取りうる値の集合」が決まっている状態・種別
' ・「数値そのもの」より名前で意味を渡したい（引数・戻り値で特に効く）
' ・選択肢を型（名前）として固定する

        'Enum enmColumnsNo
        '  No = 1
        '  Name = 2
        '  GoodLang = 3
        'End Enum
        '
        'Sub Test()
        '  Cells(5, enmColumnsNo.No).Value = "１（ナンバー）"
        '  Cells(5, enmColumnsNo.Name).Value = "２（名前）"
        '  Cells(5, enmColumnsNo.GoodLang).Value = "３（得意言語）"
        'End Sub

'■Type
' ・このプロジェクトにおける "1件分のデータの形" を定義するための道具
' ・複数の変数をひとまとめにして、１つの "構造体" として扱うための型定義
' ・Cでいう struct、他言語でいう「レコード型」に相当

        'Public Type TUser
        '       Id As Long
        '       Name As String
        '       Age As Long
        'End Type

        'Dim u As TUser
        'u.Id = 1
        'u.Name = "山田"
        'u.Age = 30

' ・バラバラの変数を、意味的に一つのまとまりとして扱える
' ・「１レコード」を表にしたい
  '（例：ユーザー情報、設定１件分、CSVの１行、DBの１行）
    '　→　「設定」という概念を1つの変数で扱える。

' ・関数の引数が増えすぎたときに使用
    ' →　引数の意味が崩れない、順番ミスも消える

' ・「関連する値の集合」を表したいとき
   ' →　配列やDictionaryより型安全で自己説明的。
   
' Type と Class の違い
'---------------- Type  ------｜-------- Class --------
'[性　　質] ただのデータ箱　|　データ＋ロジック
'[メソッド] 持てない　　　　|　持てる
'[初 期 化 ] 自動ゼロ初期化　|　Initialize
'[参　　照] 値渡し　　　　　|　参照渡し
'[用　　途] 構造体　　　　　|　オブジェクト
'[判断基準] データだけ　　　|　振る舞いも必要

'Type を使わないほうがいい場面
'・要素が頻繁に変わる
'・フィールドが動的
'・ロジックを持たせたい
'・辞書構造が必要

'■固定値のデータ管理方法
    '▼シートに配置
        '・業務ルールやマスタデータで、現場が変える可能性があるもの
        '　(例：部門コード表、商品区分、メッセージ文言、閾値、帳票設定、表示順など
        '・運用者がVBAを触れない前提で、データとして更新したいもの

    '▼コードに配置
        '・変えると動作が変わる "仕様そのもの" で、実装の整合性が必要なもの
        '　（例：処理モード、内部状態、分岐キー、エラー種別、APIの戻り値分類など
        '・変更が入るなら必ずテストやレビューが必要なもの


'■Class
'・ "もの" を表現したいとき（実体がある）
    '（例：印鑑、帳票、ジョブ、ユーザ、セッション）
'・状態＋振る舞いをセットで持たせたい
    '例：Sign()、Validate()、Render()
'・同じロジックを複数インスタンスで使う
    '例：複数ファイル、複数帳票を同時処理
'・責務を分離したい（依存関係を切りたい）
    'UI層/ロジック層/I/O層 など

'単位別変数名
'timeoutMs      :ミリ秒
'timeoutSec     :秒
'intervalMin     :分
'sizeKb            :キロバイト
'sizeMb           :メガバイト
'lengthPx        :ピクセル
'angleDeg       :度
'angleRad       :ラジアン

'------------------------------------------------------------------------------
'「C + 型名」= 型変換
'関数       変換先
'CInt       Inreger
'CLng      Long
'CDbl      Double
'CStr       String
'CBool     Boolean
'CDate     Date

Public Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As LongLong) As Long
Public Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As LongLong) As Long

'------------------------------------------------------------------------------
'■参照設定
'Microsoft Scripting Runtime
'Microsoft VBScript Regular Expressions X.X
'------------------------------------------------------------------------------
Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

Public Function NewRegExp(pattern As String) As Object
    Set Re = CreateObject("VBScript.RegExp")
    Re.pattern = pattern
    Re.Global = True
    Re.IgnoreCase = True
    Set NewRegExp = Re
End Function

Sub library()

Path = "\\////"
Set Fso = CreateObject("Scripting.FileSystemObject")
If Fso.fileExists(Path) Then
    Debug.Print "存在します。"
End If

Set Fso = Nothing


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
Set Re = CreateObject("VBScript.RegExp")

'正規表現（パターンマッチ）
'・「この文字列、メール形式？」
'・「数字だけ抜きたい」
'・「特定フォーマットだけ抽出」
'10行で100行分の文字列処理を消せる系の道具

Re.pattern = "/d+"
Re.Global = True

If Re.Test("ID=12345") Then
    Debug.Print Re.Execute("ID=12345")(0)
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
If IsNull(Tmp) Or IsEmpty(Tmp) Or Tmp = "" Then
End If

'Rangeセル用
If IsNull(rng.Value) Or rng.Value = "" Then
End If

If obj Is Nothing Then
End If

'------------------------------------------------------------------------------
'■値チェック
Tmp = rng.Value

If IsError(Tmp) Then
    'エラー
ElseIf IsEmpty(Tmp) Or Tmp = "" Then
    '空
ElseIf IsDate(Tmp) Then
    '日付
ElseIf IsNumeric(Tmp) Then
    '数値
ElseIf VarType(Tmp) = vbString Then
    '文字列
End If


'------------------------------------------------------------------------------
'■文字列操作
'関数名の後ろに「＄」を付けるとString型を返し型判定の処理が不要となる

str = Replace$("test", "t", "")
str = Left$("090-1234-5678", 3)
str = Mid$("090-1234-5678", 5, 4)
str = Right$("090-1234-5678", 3)

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
    












End Sub

