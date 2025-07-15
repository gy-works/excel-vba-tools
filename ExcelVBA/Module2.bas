Option Explicit
'ボタンから以下を呼び出してマクロを実行
Sub Module2のプロシージャ()
 Call SetHyperLinks
End Sub

Sub SetHyperLinks()
'バージョンアップシートとバージョンアップ仕様シートの相互リンク

'画面更新停止
 Application.ScreenUpdating = False
'計算を手動に
 Application.Calculation = xlCalculationManual
'マクロを実行しているブックをアクティプにする
 ThisWorkbook.Activate

Call 事前処理
Call ハイパーリンク設定1
Call ハイパーリンク設定2
Call 事後処理

  MsgBox "ハイパーリンク設定完了"

'計算を自動へ戻す
 Application.Calculation = xlCalculationAutomatic
'画面停止再開
 Application.ScreenUpdating = True
End Sub

Private Sub 事前処理()
'ハイパーリンク設定の有無を確認し有の場合リセットする
'バージョンアップシートにハイパーリンク設定が無い場合はスキップして次へ
 If Worksheets("バージョンアップ").Hyperlinks.Count < 1 Then
 MsgBox "ハイパーリンクを設定します"
   Exit Sub
  Else
 MsgBox "ハイパーリンク設定リセット"
 End If
  Call ハイパーリンククリア
End Sub

Private Sub ハイパーリンククリア()
'エラー行は無視して次へ
   On Error Resume Next

'バージョンアップシートのハイパーリンクをクリア
Worksheets("バージョンアップ").ClearHyperlinks
'バージョンアップ仕様シートのハイパーリンクをクリア
Call Find_テーブル名
End Sub

Private Sub Find_テーブル名()
'「テープル名」というワードが合まれるセルを選択し、
'そのセルをRange型で取得して次の関数に渡す
Sheet1.Select

Dim r As Range

Set r = Range("T29").CurrentRegion.Find(what:="テーブル名")

'テーブル名でないセルはスキップして次のセルへ
 If r Is Nothing Then
   MsgBox "「テープル名」のセルを特定しています"
   Exit Sub
 Else
  r.Select
 End If

 Call テーブル名ハイパーリンクのみクリア(r)

End Sub

Private Sub テーブル名ハイパーリンクのみクリア(ByRef r As Range)
'「テープル名」というワードが合まれるセル行のX列のセルのみを指定
'ハイパーリンクをクリアする処理

Dim CNT2 As Long
Dim MaxRow As Long
'バージョンアップ仕様書シート最終セルの行番号
Dim TNCell As Variant

'バージョンアップ仕様シートの全体行数取得
  MaxRow = Worksheets("バージョンアップ仕様").Range("A1").SpecialCells(xlLastCell).Row
  TNCell = r.Column + 5

エラー行は無視して次へ
 On Error Resume Next

'29行目から最終行まで実行
 For CNT2 = 29 To MaxRow

'T列が「テーブル名」でないセルはスキップ
 If Not Cells(CNT2, 20) = "テーブル名" Then
 Resume Next
Else 'T列が「テーブル名」のセルはハイパーリンククリア
  Cells(CNT2, TNCell).ClearHyperlinks
  End If
  Next CNT2
End Sub

Private Sub ハイパーリンク設定1()
Sheet1.Select
  Call バージョンアップシートでのリンク設定
End Sub

Private Sub ハイパーリンク設定2()

Sheet2.Select
  Call バージョンアップ仕様シートでのリンク設定
  Call 管理クリア処理
 End Sub

Private Sub バージョンアップシートでのリンク設定()

Dim CNT1 As Long
Dim END1 As Long 'バージョンアップシートの最終行
Dim MaxRow As Long 'バージョンアップ仕様書シート最終セルの行番号
Dim FoundCell As Variant
Dim i As Variant '位置
Dim j As Variant '検索値
Dim ShV As Worksheet
Dim ShSpc As Worksheet

Set ShV = Worksheets("バージョンアップ")
Set ShSpc = Worksheets("バージョンアップ仕様")

'バージョンアップシートに設定
 With ShV


'バージョンアップシート全体の最終行を求める
    END1 = ShV.Range("D" & Rows.Count).End(xlUp).Row
'バージョンアップ仕様シートの全体行数取得
    MaxRow = ShSpc.Range("A1").SpecialCells(xlLastCell).Row

'エラー行は無視して次へ
    On Error Resume Next

'3行目から最終行までの行数分実行
 For CNT1 = 3 To END1
'バージョンアップシートD列最終行までの検索値
 j = ShV.Range("D" & CNT1)
'検索値をバージョンアップ仕様シートから探す
 Set FoundCell = ShSpc.Range("X28:X" & MaxRow).Find(j)
'B列が「決算期別」でない場合は無視して次へ
 If Not Cells(CNT1, 2) = "決算期別" Then
 Resume Next

 Else 'ハイパーリンク設定
 i = FoundCell.Address(False, False) ' 相対参照で取得
 ShV.Range("D" & CNT1).Hyperlinks.Add Anchor:=ShV.Range("D" & CNT1), Address:="", SubAddress:="'バージョンアップ仕様'!" & i
 End If
continue:
 Next CNT1
 End With
End Sub

Private Sub バージョンアップ仕様シートでのリンク設定()

Dim CNT2 As Long
Dim MaxRow As Long  'バージョンアップ仕様書シート行数
Dim END1 As Long   'バージョンアップシートの最終行
Dim FoundCell2 As Variant
Dim K As Variant  '検索値
Dim h As Variant
Dim ShV As Worksheet
Dim ShSpc As Worksheet

Set ShSpc = Worksheets("バージョンアップ仕様")
Set ShV = Worksheets("バージョンアップ")
'バージョンアップ仕様シートに設定
With ShSpc

'バージョンアップ仕様シートの全体行数取得
  MaxRow = ShSpc.Range("A1").SpecialCells(xlLastCell).Row

'バージョンアップシート全体の最終行を求める
  END1 = ShV.Range("D" & Rows.Count).End(xlUp).Row

'エラー行は無視して次へ
  On Error Resume Next

'29行目から最終行まで実行
  For CNT2 = 29 To MaxRow
'バージョンアップ仕様シートX列最終行までの検索値
  K = ShSpc.Range("X" & CNT2)
'検索値をバージョンアップシートD列から探す
  Set FoundCell2 = ShV.Range("D3:D" & END1).Find(K)
'T列が「テープル名」でない行はスキップ
  If Not Cells(CNT2, 20) = "テーブル名" Then
  Resume Next

  Else 'バージョンアップ仕様シートへのハイパーリンク設定
  h = FoundCell2.Address(False, False) ' 相対参照で取得
  ShSpc.Range("X" & CNT2).Hyperlinks.Add Anchor:=ShSpc.Range("X" & CNT2), Address:="", SubAddress:="'バージョンアップ'!" & h
 End If
continue:
 Next CNT2

End With
End Sub

Private Sub 管理クリア処理()
'「決算期別」以外の不要なハイパーリンク設定を解除する
Sheet1.Select
 Call Find_管理
End Sub

Private Sub Find_管理()
'バージョンアップシートで「管理」というワードが合まれるセルを検索
'セルをRange型で取得して次の関数に渡す
 Dim Findcell As Range

'オートフィルタの選択がある場合は解除
 If Worksheets("バージョンアップ").FilterMode = True Then
 Worksheets("バージョンアップ").ShowAllData
 End If
'B列から「管理」を検索
 Set Findcell = Range("B:B").Find(what:="管理", LookAt:=xlWhole)

'検索セルが無い場合は終了
 If Findcell Is Nothing Then
  Exit Sub
 Else '検索セルをRange型で取得して次の関数へ
  Call 管理リンク位置検索(Findcell)
 End If

End Sub

Private Sub 管理リンク位置検索(ByRef Findcell As Range)
'バージョンアップシートで「管理」先頭行のセルから2列隣(D列)を特定
'特定したセルと同じ文字列のセルをバージョンアップ仕様シートから検索
'同じ文字列の位置をRange型で取得しハイパーリンク解除処理へ
Sheet1.Select
Dim rng As Range
Dim SarchRow As Long, SarchColumn As Long
Dim Ent1 As Variant '検索セルの値
Dim MaxRow As Long 'バージョンアップ仕様書シート行数
Dim ShV As Worksheet
Dim ShSpc As Worksheet

Set ShV = Worksheets("バージョンアップ")
Set ShSpc = Worksheets("バージョンアップ仕様")

'バージョンアップシートを「管理」で検索したセルの2列隣を指定
SarchRow = Findcell.Row
SarchColumn = Findcell.Column + 2

  Ent1 = ShV.Cells(SarchRow, SarchColumn).Value '参照セルの値=検索値
'バージョンアップ仕様シートの全体行数取得
  MaxRow = ShSpc.Range("A1").SpecialCells(xlLastCell).Row

'参照セルの値をバージョンアップ仕様シートX列から探す
  Set rng = ShSpc.Range("X29:X" & MaxRow).Find(Ent1)
'該当セルがなければ終了
  If rng Is Nothing Then
  Exit Sub
  Else
'該当セルを選択してハイパーリンククリアに進む
  ShSpc.Activate ' バージョンアップ仕様シートをアクティブ
  rng.Select
    Call 管理ハイパーリンククリア(rng)
  End If
End Sub

Private Sub 管理ハイパーリンククリア(ByRef rng As Range)
'バージョンアップ仕様シートで特定したセルを起点として範囲指定
'設定されたハイパーリンクを一括クリアする
Sheet2.Select
Dim MaxRow As Long
Dim c As Long
Dim ShSpc As Worksheet

Set ShSpc = Worksheets("バージョンアップ仕様")

'バージョンアップ仕様シートの全体行数取得
  MaxRow = Range("A1").SpecialCells(xlLastCell).Row
'バージョンアップ仕様シートの「管理」先頭行のセル
  c = rng.Row

'検索結果のセルを起点として最終行までを範囲指定
  With ShSpc.Range(Cells(c, 24), Cells(MaxRow, 24))
   .ClearHyperlinks 'ハイパーリンク解除
   .Font.Underline = False '文字のアンダーライン解除
   .Font.ColorIndex = xlAutomatic '文字色を自動に設定

  End With
 End Sub

Private Sub 事後処理()
'オートフィルタの選択が残っていたら解除
  If Worksheets("バージョンアップ").FilterMode = True Then
  Worksheets("バージョンアップ").ShowAllData
  End If

'バージョンアップシートをアクティプにしてA1セルを選択する
  Worksheets("バージョンアップ").Activate
  Range("A1").Select

End Sub
