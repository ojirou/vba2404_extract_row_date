Attribute VB_Name = "Module1"
Sub extract_row_date()
    ' 変数の宣言
    Dim search_date As Date
    Dim count As Integer
    Dim name, num, weather As String
' 列を拡張する場合⇒　 上の行を　Dim name, num, weather, abc As String に変更
   Dim sheet As Worksheet
    Dim i, j As Long
    ' 検索日を取得
    search_date = Sheets("日別一覧").Range("C1").Value
    count = 0
    ' 各シートに対してループ処理を実行
    For Each sheet In ThisWorkbook.Sheets
        ' “田中”、”佐藤”、”鈴木”のシートの場合のみ処理を実行
        ' Sheet“〇〇”を追加する場合 Then  の前に、「Or sheet.Name = “〇〇” 」を追記
        If sheet.name = "田中" Or sheet.name = "佐藤" Or sheet.name = "鈴木" Then
' Sheet“〇〇”を追加する場合⇒ Then  の前に、「Or sheet.Name = “〇〇” 」を追記
            With sheet
                ' 各行に対してループ処理を実行
                StartRow = 3
' リスト先頭行を変更する場合⇒StartRow=5（5行目からに変更）
             For i = StartRow To .Cells(.Rows.count, 1).End(xlUp).Row
                    ' 検索日と一致する行の場合のみ処理を実行
                    If .Cells(i, 1).Value = search_date Then
                        count = count + 1
                        name = .name
                        num = .Cells(i, 2).Value
                        weather = .Cells(i, 3).Value
                   ' 日別一覧シートにデータを書き込み
                        With Sheets("日別一覧")
                            .Cells(3 + count, 2).Value = name
                            .Cells(3 + count, 3).Value = search_date
                            .Cells(3 + count, 4).Value = num
                            .Cells(3 + count, 5).Value = weather
' 列を拡張する場合⇒　.Cells(3 + count, 5).Value = abc  を追加
                            ' セルの書式設定
                            For j = 2 To 5
' 列を拡張する場合⇒　上の行を　For j = 2 To 6 に変更
                                .Cells(3 + count, j).HorizontalAlignment = xlCenter
                                .Cells(3 + count, j).Borders.LineStyle = xlContinuous
                            Next j
                        End With
                    End If
                Next i
            End With
        End If
    Next sheet
End Sub