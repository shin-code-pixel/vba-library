Attribute VB_Name = "Mdl_TestCaseSheet"
Option Explicit

Public Sub VBL_GenerateReportTestCaseTable()
    Dim ws As Worksheet
    Dim r As Long
    Dim i As Long
    Dim cases As Collection

    ' 出力シート
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("TestCases_Report")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = "TestCases_Report"
    Else
        ws.Cells.Clear
    End If

    ' ヘッダ（実務向け固定）
    ws.Range("A1:I1").Value = Array( _
        "ID", "カテゴリ", "チェック観点", "テスト手順", "入力/条件例", _
        "期待結果", "結果", "実施日", "備考" _
    )
    ws.Rows(1).Font.Bold = True
    ws.Rows(1).WrapText = True

    '=========================
    ' 縦軸（網羅）: 叩き台ケース
    '=========================
    Set cases = New Collection

    ' --- 入力バリデーション／境界値 ---
    cases.Add Array("REP-VAL-001", "VAL/BND", "必須項目：空", "必須項目を未入力で実行", "必須=空", "エラー表示。対象項目が特定できる。処理は中断される。")
    cases.Add Array("REP-VAL-002", "VAL/BND", "境界値：最小値", "最小値を入力して実行", "数値=最小", "正常終了し、出力値が仕様通り。")
    cases.Add Array("REP-VAL-003", "VAL/BND", "境界値：最大値", "最大値を入力して実行", "数値=最大", "正常終了し、出力値が仕様通り。")
    cases.Add Array("REP-VAL-004", "VAL/BND", "数値：0", "0を入力して実行", "数値=0", "仕様通り（許可/拒否）が成立。")
    cases.Add Array("REP-VAL-005", "VAL/BND", "数値：負数", "負数を入力して実行", "数値=-1", "仕様通り（拒否ならエラー）。")
    cases.Add Array("REP-VAL-006", "VAL/BND", "数値：桁あふれ", "最大桁+1で実行", "数値=桁超過", "入力段階または実行時に拒否できる。")
    cases.Add Array("REP-VAL-007", "VAL/BND", "文字列：最大長", "最大長ちょうどで実行", "文字列=MAXLEN", "正常終了し、欠け/崩れがない。")
    cases.Add Array("REP-VAL-008", "VAL/BND", "文字列：最大長超過", "最大長+1で実行", "文字列=MAXLEN+1", "拒否できる（エラー/案内）。")
    cases.Add Array("REP-VAL-009", "VAL/BND", "文字列：全角/半角混在", "全角半角混在で実行", "例：ＡB１２3", "仕様通りに保持/変換される。")
    cases.Add Array("REP-VAL-010", "VAL/BND", "文字列：先頭末尾スペース", "前後スペース付きで実行", "例：' ABC '", "仕様通り（trim有無）が成立。")
    cases.Add Array("REP-VAL-011", "VAL/BND", "日付：不正", "不正日付で実行", "例：2026/02/30", "拒否できる（エラー/案内）。")
    cases.Add Array("REP-VAL-012", "VAL/BND", "日付：うるう日", "うるう日で実行", "例：2024/02/29", "正常終了し、出力が仕様通り。")
    cases.Add Array("REP-VAL-013", "VAL/BND", "期間：開始>終了", "開始日>終了日で実行", "開始=翌日, 終了=当日", "拒否できる（相関チェック）。")
    cases.Add Array("REP-VAL-014", "VAL/BND", "選択肢：未選択", "未選択で実行", "選択=空", "拒否できる（必須なら）。")
    cases.Add Array("REP-VAL-015", "VAL/BND", "選択肢：想定外値（貼付け）", "貼付けで不正値を入れて実行", "選択='XXX'", "拒否できる（入力規則/実行時チェック）。")

    ' --- 書式・表示（帳票としての品質） ---
    cases.Add Array("REP-FMT-001", "FMT", "罫線：崩れ", "通常入力で出力生成", "標準データ", "罫線が欠けない/ズレない。")
    cases.Add Array("REP-FMT-002", "FMT", "フォント：種類/サイズ", "通常入力で出力生成", "標準データ", "フォント指定が保持される。")
    cases.Add Array("REP-FMT-003", "FMT", "折り返し：表示欠け", "長文入力で生成", "長文（2～3行）", "欠けずに表示される（行高含む）。")
    cases.Add Array("REP-FMT-004", "FMT", "表示形式：日付/数値/0埋め", "代表項目を生成して確認", "日付/数値/0埋め", "表示形式が仕様通り。")
    cases.Add Array("REP-FMT-005", "FMT", "列幅/行高：崩れ", "通常入力で生成", "標準データ", "列幅行高が意図通り。")
    cases.Add Array("REP-FMT-006", "FMT", "印刷範囲", "印刷プレビュー/印刷設定確認", "標準データ", "印刷範囲が適切。")
    cases.Add Array("REP-FMT-007", "FMT", "改ページ：複数枚", "複数ページになるデータで生成", "長データ", "分割が意図通り。")

    ' --- ファイル入出力 ---
    cases.Add Array("REP-IO-001", "IO", "保存：新規", "新規保存で生成", "保存先=既存フォルダ", "指定先に保存される。")
    cases.Add Array("REP-IO-002", "IO", "保存：上書き", "同名ファイルありで生成", "同名ファイルあり", "仕様通り（確認/リネーム/上書き）。")
    cases.Add Array("REP-IO-003", "IO", "ファイル名：禁止文字", "禁止文字含む名前で生成", "例：A:B", "拒否または自動置換される。")
    cases.Add Array("REP-IO-004", "IO", "ファイルロック", "出力先ファイルを開いたまま生成", "対象ファイル=開く", "仕様通り（失敗時メッセージ明確）。")

    ' --- フロー／前提不備 ---
    cases.Add Array("REP-FLOW-001", "FLOW", "テンプレ不在", "テンプレ削除/リネームして実行", "テンプレ無し", "適切に停止し案内。")
    cases.Add Array("REP-FLOW-002", "FLOW", "同名シート存在", "同名シートがある状態で生成", "既存シートあり", "仕様通り（上書き/新規名/停止）。")
    cases.Add Array("REP-FLOW-003", "FLOW", "連続生成：取り違え無し", "2件以上連続で生成", "顧客A→顧客B", "出力が混ざらない。")

    ' --- エラー処理 ---
    cases.Add Array("REP-ERR-001", "ERR", "入力不備：原因特定", "意図的に入力不備で実行", "必須空など", "どの項目が原因か分かる。")
    cases.Add Array("REP-ERR-002", "ERR", "例外：最低限ログ", "例外を発生させる操作", "異常系", "Err.Number/Description等が残る。")

    ' --- 環境差／権限 ---
    cases.Add Array("REP-CFG-001", "CFG/CON", "参照設定差", "参照無し/有りで実行", "参照状態差", "動作条件が崩れない/明記通り。")
    cases.Add Array("REP-CFG-002", "CFG/CON", "権限なし保存先", "権限なしフォルダで生成", "保存先=権限なし", "失敗し、案内が明確。")

    ' --- 性能 ---
    cases.Add Array("REP-PERF-001", "PERF", "件数増：時間許容", "件数増で生成", "100/1000件", "許容時間内 or 制約が明記。")

    '=========================
    ' 出力
    '=========================
    r = 2
    For i = 1 To cases.Count
        ws.Cells(r, 1).Value = cases(i)(0) ' ID
        ws.Cells(r, 2).Value = cases(i)(1) ' カテゴリ
        ws.Cells(r, 3).Value = cases(i)(2) ' 観点
        ws.Cells(r, 4).Value = cases(i)(3) ' 手順
        ws.Cells(r, 5).Value = cases(i)(4) ' 入力例
        ws.Cells(r, 6).Value = cases(i)(5) ' 期待結果
        r = r + 1
    Next i

    ' 結果列：PASS/FAIL/SKIP を選択
    With ws.Range("G2:G" & r - 1)
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
            Formula1:="PASS,FAIL,SKIP"
    End With

    ' 見た目（最低限）
    ws.Columns("A:I").ColumnWidth = 18
    ws.Columns("D:F").ColumnWidth = 34
    ws.Rows(1).RowHeight = 24
    ws.Range("A1:I1").Interior.ColorIndex = 15
    ws.Range("A1:I" & r - 1).Borders.LineStyle = xlContinuous
    ws.Columns("A:I").VerticalAlignment = xlVAlignTop
    ws.Columns("D:F").WrapText = True

    MsgBox "テストケース表（叩き台）を生成しました：TestCases_Report", vbInformation
End Sub

