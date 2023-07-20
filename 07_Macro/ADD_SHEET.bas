Attribute VB_Name = "ADD_SHEET"
'***NewBookT() �V�K�u�b�N�쐬
'***NewTableSheet() �e�[�u�����ڃV�[�g�}��
'***DeleteSheet �V�[�g��DELETE�i�s�g�p�j
'***Runstop �������f

'==========================================================
'�y�v���V�[�W�����zNewBookT
'�y�T�@�v�z�e�[�u����`���̒ǉ�
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================
Sub NewBookT()
    Dim wnum As Integer
    Dim i As Integer
    Dim ans As Integer
    Dim sname As String
    
    '�e�[�u���V�[�g���擾
    Tb_SheetNMInp = ""
    Tb_SheetNMInp = UCase(InputBox("�e�[�u��ID����͂��Ă�������"))
    
    If Len(Tb_SheetNMInp) > 31 Then
        MsgBox "�e�[�u��ID��31���ȓ��œ��͂��Ă��������B"
        Exit Sub
    End If
    
    If Tb_SheetNMInp = "" Then
        Exit Sub
    Else
        '���[�N�u�b�N���J����Ă��邩���ׂ�
        If Workbooks.Count = 0 Then
            Application.ScreenUpdating = False
            Workbooks().Add
            wnum = Worksheets.Count
            NewTableSheet (ActiveWorkbook.name)
            Application.DisplayAlerts = False   '�x�����b�Z�[�W�\��OFF
            For i = 1 To wnum
                Worksheets(2).Delete
            Next i
            Application.DisplayAlerts = True   '�x�����b�Z�[�W�\��ON
            Application.ScreenUpdating = True
        Else
            Application.ScreenUpdating = False
            If Worksheets().Count > 1 Then
                DeleteSheet (Tb_SheetNMInp)
                NewTableSheet (ActiveWorkbook.name)
            Else
                If HaveSheet(ActiveWorkbook.name, Tb_SheetNMInp) > 0 Then
                    Workbooks(ActiveWorkbook.name).Worksheets(Tb_SheetNMInp).name = Tb_SheetNMInp & "(1)"
                    NewTableSheet (ActiveWorkbook.name)
                    Workbooks(ActiveWorkbook.name).Worksheets(Tb_SheetNMInp).name = Tb_SheetNMInp & "(2)"
                    Workbooks(ActiveWorkbook.name).Worksheets(Tb_SheetNMInp & "(1)").name = Tb_SheetNMInp
                Else
                    NewTableSheet (ActiveWorkbook.name)
                End If
            Application.ScreenUpdating = True
            End If
        End If
    End If
End Sub

'==========================================================
'�y�v���V�[�W�����zNewTableSheet
'�y�T�@�v�z�e�[�u����`���̒ǉ�
'�y���@���zbook����
'�y�߂�l�z�Ȃ�
'==========================================================

Sub NewTableSheet(bname As String)
        Dim wkrange As String
        Workbooks("voyager.xla").Worksheets(Tb_SheetNm).Copy _
            before:=Workbooks(bname).Worksheets(1)
        Workbooks(bname).Worksheets(Tb_SheetNm).name = Tb_SheetNMInp
        Workbooks(bname).Worksheets(Tb_SheetNMInp).Activate
        'Cells(R_TblId, C_TblId).Value = Tb_SheetNMInp '�e�[�u��ID
        Cells(R_TblId2, C_TblId2).Value = Tb_SheetNMInp '�e�[�u������
        Cells(R_COLNAME, C_COLNAME).Select '�J�����𕨗����ɃZ�b�g
        kaktyoNV
End Sub


'==========================================================
'�y�v���V�[�W�����zDeleteSheet
'�y�T�@�v�z�V�[�g�폜
'�y���@���zsheet����
'�y�߂�l�z�Ȃ�
'==========================================================
Sub DeleteSheet(sname As String)
    Dim ret As Integer
    ret = HaveSheet(ActiveWorkbook.name, sname)
    If ret > 0 Then
        Beep
        answer = MsgBox(sname & " �͊��ɍ쐬����Ă��܂� " & Chr(13) & "�㏑�����Ă�낵���ł����H", vbQuestion + vbOKCancel)
        If answer = vbOK Then
            Application.DisplayAlerts = False   '�x�����b�Z�[�W�\��OFF
            ActiveWorkbook.Worksheets(ret).Delete
            Application.DisplayAlerts = True   '�x�����b�Z�[�W�\��ON
        Else
            RunStop
        End If
    End If
End Sub


'==========================================================
'�y�v���V�[�W�����zRunStop
'�y�T�@�v�z�������f
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================

Sub RunStop()
    Application.StatusBar = "�����𒆒f���܂���."
    Application.Cursor = xlNormal
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    End
    Stop
End Sub
