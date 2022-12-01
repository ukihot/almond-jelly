---
title: "CSV出力マクロのリファクタリング"
emoji: "🌭"
type: "tech" # tech: 技術記事 / idea: アイデア
topics: ["vba"]
published: false
---

# 経緯
あるマクロを別の要件に使いまわすことになった．
6年以上前に前任者が作ったようで，そのまま使いまわしても動くが，色々修正しがいのある様子なのでこの際書き直すことにした．

# やりたいこと
Excelに入力された値をもとにCSVを作成したい．

# 今回修正する対象コード
## このコードの果たしていた役割
14個のカラムをもつtsvファイルを作成していた．

## ロジックの概要
1. tsvファイルの作成
1. セルを(1行1列目,1行2列目, ... ,2行1列目, ...)の順で走査
2. 14列目で切り返し
3. 4列目が空欄だったらそこでやめ
4. 各セルの値をファイルへ書き込み
5. 最後にリモートのファイルサーバにpush

:::details 元となるソースコード全体
```vba
Const pass As String = "\\192.168.xxx.yyy\hoge\fuga\piyo\"
'Const pass As String = "C:\test\"
Const CT As Integer = 4
Const LAST As Integer = 14
Const pto As String = "出力シート"

Sub CSVout()

    Dim i, j As Long
    Dim koumoku As String
    Dim hiduke As String
    Dim csvname As String
    Dim csvFile As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    
    koumoku = Cells(2, 17)
    hiduke = Cells(2, 2)

    csvname = hiduke & "_" & koumoku & ".tsv"
    csvFile = pass & csvname
 
    Open csvFile For Output As #1
 
    i = 1
 
    Do While ws.Cells(i, CT).Value <> ""
 
      j = 1
      
        Do While j < LAST
    
            Print #1, ws.Cells(i, j).Value & vbTab;
            j = j + 1
 
        Loop
 
    Print #1, ws.Cells(i, j).Value & vbCr;
    
    i = i + 1
 
    Loop
 
    Close #1
 
    MsgBox csvname & "を書き出しました"
    
    ws.Tab.ColorIndex = 10
End Sub
```
:::

# 変更するところ

* tsv ではなく csv
* 14カラムではなく5カラム
* 5カラムのうち2カラムは入力値ではなくシステム的に決定するので入力不要
* Excelとしては3列用意してあげて，できあがるcsvは5カラムあればよい．
* シート分け仕様も追加