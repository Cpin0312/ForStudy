Attribute VB_Name = "HELP_SAMPLE_VIEW"
'**helpview() �w���v�V�[�g�\��
'**sumpleview()�@�T���v���V�[�g�\��
'**dialogshow() �_�C�A���O�\��


'==========================================================
'�y�֐����zhelpview
'�y�T�@�v�z�w���v�V�[�g��\��
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================

Sub helpview()
    Dim wnum As Integer
    '�J����Ă��郏�[�N�u�b�N�����邩���ׂ�
    '���[�N�u�b�N���J����Ă��Ȃ��ꍇ���[�N�u�b�N��ǉ����ăV�[�g�ǉ�
    If Workbooks.Count = 0 Then
        Workbooks.Add
        Workbooks(TEMPLATE).Worksheets(HELPSHEET).Copy _
        before:=Workbooks(1).Worksheets(1)
        '���b�Z�[�W��OFF�ɂ���
        Application.DisplayAlerts = False
        '�]���ȃV�[�g���폜����
        For i = 2 To Worksheets.Count
            Worksheets(2).Delete
        Next i
        Application.DisplayAlerts = True
    Else
        wnum = HaveSheet(ActiveWorkbook.name, HELPSHEET)
        If wnum > 0 Then
          ActiveWorkbook.Worksheets(wnum).Activate
        Else
            wnum = HaveSheet2(ActiveWorkbook.name, 9)
            If wnum > 0 Then
                ActiveWorkbook.Worksheets(wnum).Activate
            Else
                Workbooks(TEMPLATE).Worksheets(HELPSHEET).Copy _
                before:=Worksheets(1)
            End If
        End If
    End If
End Sub


'==========================================================
'�y�֐����zsampleview
'�y�T�@�v�z�T���v���V�[�g��\��
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================

Sub sampleview()
    Dim wnum As Integer
    '�J����Ă��郏�[�N�u�b�N�����邩���ׂ�
    '���[�N�u�b�N���J����Ă��Ȃ��ꍇ���[�N�u�b�N��ǉ����ăV�[�g�ǉ�
    If Workbooks.Count = 0 Then
        Workbooks.Add
        GV_book = ActiveWorkbook.name
        Workbooks(TEMPLATE).Worksheets(SAMPLESHEET).Copy _
        before:=Workbooks(1).Worksheets(1)
        '���b�Z�[�W��OFF�ɂ���
        Application.DisplayAlerts = False
        '�]���ȃV�[�g���폜����
        For i = 2 To Worksheets.Count
            Worksheets(2).Delete
        Next i
        Application.DisplayAlerts = True
    Else
        wnum = HaveSheet(ActiveWorkbook.name, SAMPLESHEET)
        If wnum > 0 Then
            ActiveWorkbook.Worksheets(wnum).Activate
        Else
            Workbooks(TEMPLATE).Worksheets(SAMPLESHEET).Copy _
            before:=Worksheets(1)
        End If
    End If
End Sub

Sub dialogshow()
    VOYAGERDIALOG.Show
    'ActiveWorkbook.DialogSheets("Dialog1").Show
End Sub

