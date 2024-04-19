Attribute VB_Name = "Module1"
Sub extract_row_date2()
    ' 変数の宣言
    Dim search_date As Date
    Dim count As Integer
    Dim sheet As Worksheet
    Dim tempSheet As Worksheet
    Dim i As Long

    ' 検索日を取得
    search_date = Sheets("日別一覧").Range("C1").Value
    count = 0

    ' temporaryシートを作成するかチェックし、存在しない場合は作成する
    On Error Resume Next
    Set tempSheet = Sheets("temporary")
    On Error GoTo 0
    If tempSheet Is Nothing Then
        Set tempSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.count))
        tempSheet.name = "temporary"
    End If

    ' 各シートに対してループ処理を実行
    For Each sheet In ThisWorkbook.Sheets
        ' “日別一覧”、”temporary”以外のシートに対して処理を実行
        If sheet.name <> "日別一覧" And sheet.name <> "temporary" Then
            With sheet
                ' 各行に対してループ処理を実行
                StartRow = 3
                ' リスト先頭行を変更する場合⇒StartRow=5（5行目からに変更）
                For i = StartRow To .Cells(.Rows.count, 1).End(xlUp).Row
                    ' 検索日と一致する行の場合のみ処理を実行
                    If .Cells(i, 1).Value = search_date Then
                        count = count + 1
                        name = .name
                   ' 日別一覧シートにデータを書き込み
                        With Sheets("日別一覧")
                            .Cells(3 + count, 2).Value = name
                            .Cells(3 + count, 2).HorizontalAlignment = xlCenter
                            .Cells(3 + count, 2).Borders.LineStyle = xlContinuous
                            .Cells(3 + count, 2).Interior.Color = RGB(255, 255, 255)
                        End With
                        .Rows(i).Copy Destination:=tempSheet.Rows(3 + count)
'                        With Sheets("temporary")
'                            .Rows(3 + count) = xlCenter
'                            .Rows(3 + count) = xlContinuous
''                            .Rows(3 + count).Interior.Color = RGB(255, 255, 255)
'                        End With
                    End If
                Next i
            End With
        End If
    Next sheet
        ' temporaryシートから値が入力されているセル範囲をコピーして、Sheet"日別一覧"のCells(4,3)に貼り付ける
    If Not tempSheet Is Nothing Then
        On Error Resume Next
        Set copyRange = tempSheet.UsedRange
        On Error GoTo 0
        If Not copyRange Is Nothing Then
            copyRange.Copy Destination:=Sheets("日別一覧").Cells(4, 3)
        End If
    End If
        ' temporaryシートを削除する
    Application.DisplayAlerts = False ' 確認ダイアログを非表示にする
    If Not tempSheet Is Nothing Then tempSheet.Delete
    Application.DisplayAlerts = True
End Sub