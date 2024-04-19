Attribute VB_Name = "Module1"
Sub extract_row_date2()
    ' �ϐ��̐錾
    Dim search_date As Date
    Dim count As Integer
    Dim sheet As Worksheet
    Dim tempSheet As Worksheet
    Dim i As Long

    ' ���������擾
    search_date = Sheets("���ʈꗗ").Range("C1").Value
    count = 0

    ' temporary�V�[�g���쐬���邩�`�F�b�N���A���݂��Ȃ��ꍇ�͍쐬����
    On Error Resume Next
    Set tempSheet = Sheets("temporary")
    On Error GoTo 0
    If tempSheet Is Nothing Then
        Set tempSheet = ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.count))
        tempSheet.name = "temporary"
    End If

    ' �e�V�[�g�ɑ΂��ă��[�v���������s
    For Each sheet In ThisWorkbook.Sheets
        ' �g���ʈꗗ�h�A�htemporary�h�ȊO�̃V�[�g�ɑ΂��ď��������s
        If sheet.name <> "���ʈꗗ" And sheet.name <> "temporary" Then
            With sheet
                ' �e�s�ɑ΂��ă��[�v���������s
                StartRow = 3
                ' ���X�g�擪�s��ύX����ꍇ��StartRow=5�i5�s�ڂ���ɕύX�j
                For i = StartRow To .Cells(.Rows.count, 1).End(xlUp).Row
                    ' �������ƈ�v����s�̏ꍇ�̂ݏ��������s
                    If .Cells(i, 1).Value = search_date Then
                        count = count + 1
                        name = .name
                   ' ���ʈꗗ�V�[�g�Ƀf�[�^����������
                        With Sheets("���ʈꗗ")
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
        ' temporary�V�[�g����l�����͂���Ă���Z���͈͂��R�s�[���āASheet"���ʈꗗ"��Cells(4,3)�ɓ\��t����
    If Not tempSheet Is Nothing Then
        On Error Resume Next
        Set copyRange = tempSheet.UsedRange
        On Error GoTo 0
        If Not copyRange Is Nothing Then
            copyRange.Copy Destination:=Sheets("���ʈꗗ").Cells(4, 3)
        End If
    End If
        ' temporary�V�[�g���폜����
    Application.DisplayAlerts = False ' �m�F�_�C�A���O���\���ɂ���
    If Not tempSheet Is Nothing Then tempSheet.Delete
    Application.DisplayAlerts = True
End Sub