Attribute VB_Name = "Mdl02_Config"
Option Explicit


'1)基本・制御系
Public i, j, k As LongLong  'ループ処理
Public idx As String    '索引【index】
Public cnt As LongLong  'カウント【count】
Public ret As Variant   '戻り値【return value】
Public res As Variant   '結果【result】
Public flag As Boolean  'フラグ【flag】
Public isOK As Boolean  'OK結果
Public hasError As Boolean  'エラー発生
Public fileExists As Boolean    '存在チェック

'2)オブジェクト
Public fso As Object    'オブジェクト
Public obj As Object    'オブジェクト
Public dict As Object   'Dictionaryオブジェクト
Public re As Object 'RegExpオブジェクト
Public wb As Workbook   'ブック
Public ws As Worksheet  'シート
Public rng As Range 'セル
Public tbl As ListObject    'テーブル操作

'3)データ系
Public arr As Variant   '配列【array】
Public colArr As Variant    '配列【array】
Public result() As Variant  '配列【array】
Public co As Collection '項目(item)とキー(key)をセットで格納するオブジェクト
Public list As Variant  'リストボックス【ListBox】
Public data As Variant  '汎用データ
Public rowData As Variant   '行データ
Public colData As Variant   '列データ

'4)数値・位置
Public row  As LongLong '行
Public startRow As LongLong '開始行
Public endRow As LongLong '最終行
Public col As LongLong '列
Public startCol As LongLong '開始列
Public endCol As LongLong   '最終列
Public pos As Variant   '位置【position】

'5)文字列・ファイル
Public str As String    '文字列
Public path As String   'ファイルパス
Public fileName As String   'ファイル名
Public folderPath As String 'フォルダパス
Public ext As String    '拡張子

'6)日付・時間
Public dt As Date   '日付
Public nowDt As Date    '今の日時
Public startDt As Date  '開始日時
Public endDt As Date    '終了日時
Public elapsed As Date  '経過日時

'7)処理フロー
Public src As Variant   'コピー元【source】
Public dst As Variant   'コピー先【destination】
Public inputVal As Variant  '入力値
Public outputVal As Variant '出力値
Public tmp As Variant   '一時変数
Public buf As Variant   'バッファ【buffer】

'8)エラー・ログ
Public errNum As Variant    'エラー名
Public errMsg As Variant    'エラーメッセージ
Public log As Variant 'ログ
Public status As Variant  '状況：ステータス
