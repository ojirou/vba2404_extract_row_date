Attribute VB_Name = "Module1"
Sub clear_rows()
    With Sheets("���ʈꗗ")
        .Rows("4:" & .Rows.count).Delete
    End With
End Sub
