Attribute VB_Name = "Module1"
Sub extract_row_date()
    ' �ϐ��̐錾
    Dim search_date As Date
    Dim count As Integer
    Dim name, num, weather As String
' ����g������ꍇ�ˁ@ ��̍s���@Dim name, num, weather, abc As String �ɕύX
   Dim sheet As Worksheet
    Dim i, j As Long
    ' ���������擾
    search_date = Sheets("���ʈꗗ").Range("C1").Value
    count = 0
    ' �e�V�[�g�ɑ΂��ă��[�v���������s
    For Each sheet In ThisWorkbook.Sheets
        ' �g�c���h�A�h�����h�A�h��؁h�̃V�[�g�̏ꍇ�̂ݏ��������s
        ' Sheet�g�Z�Z�h��ǉ�����ꍇ Then  �̑O�ɁA�uOr sheet.Name = �g�Z�Z�h �v��ǋL
        If sheet.name = "�c��" Or sheet.name = "����" Or sheet.name = "���" Then
' Sheet�g�Z�Z�h��ǉ�����ꍇ�� Then  �̑O�ɁA�uOr sheet.Name = �g�Z�Z�h �v��ǋL
            With sheet
                ' �e�s�ɑ΂��ă��[�v���������s
                StartRow = 3
' ���X�g�擪�s��ύX����ꍇ��StartRow=5�i5�s�ڂ���ɕύX�j
             For i = StartRow To .Cells(.Rows.count, 1).End(xlUp).Row
                    ' �������ƈ�v����s�̏ꍇ�̂ݏ��������s
                    If .Cells(i, 1).Value = search_date Then
                        count = count + 1
                        name = .name
                        num = .Cells(i, 2).Value
                        weather = .Cells(i, 3).Value
                   ' ���ʈꗗ�V�[�g�Ƀf�[�^����������
                        With Sheets("���ʈꗗ")
                            .Cells(3 + count, 2).Value = name
                            .Cells(3 + count, 3).Value = search_date
                            .Cells(3 + count, 4).Value = num
                            .Cells(3 + count, 5).Value = weather
' ����g������ꍇ�ˁ@.Cells(3 + count, 5).Value = abc  ��ǉ�
                            ' �Z���̏����ݒ�
                            For j = 2 To 5
' ����g������ꍇ�ˁ@��̍s���@For j = 2 To 6 �ɕύX
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