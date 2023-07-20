Attribute VB_Name = "HIDE_VIEW"
'*** kaktyoV() �g���J�����\��
'*** kaktyoNV() �g���J������\��
'*** seisho() ����
'*** dialogshow ���\��


'==========================================================
'�y�v���V�[�W�����zkaktyoV
'�y�T�@�v�z�g���J������\��
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================
Sub kaktyoV()
    Dim sname As Integer
    Dim i As Integer
        
    If Workbooks.Count = 0 Then
        MsgBox ("�u�b�N������܂���")
        Exit Sub
    End If
    
    sname = ActiveWorkbook.ActiveSheet.Cells(R_DocId, C_DocId)  '--- Mod S.Iwanaga 2010/04/08
    
    If sname <> 1 Then
        MsgBox ("�e�[�u����`�����A�N�e�B�u�ɂ��Ă�������")
    Else
        ActiveSheet.Unprotect
        Columns(C_HideSNm & ":" & C_HideENm).Select     '--- Mod S.Iwanaga 2010/04/08
        Selection.EntireColumn.Hidden = False
        
        'ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        'False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        'AllowFormattingRows:=True
    End If
        
End Sub


'==========================================================
'�y�v���V�[�W�����zkaktyoNV
'�y�T�@�v�z�g���J�������\��
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================

Sub kaktyoNV()
    Dim sname As Integer
    Dim i As Integer
    
    '--- Add Start S.Iwanaga 2010/04/08
    If Workbooks.Count = 0 Then
        MsgBox ("�u�b�N������܂���")
        Exit Sub
    End If
    '--- Add End
    
    sname = ActiveWorkbook.ActiveSheet.Cells(R_DocId, C_DocId)  '--- Mod S.Iwanaga 2010/04/08
    
    If sname <> 1 Then
        MsgBox ("�e�[�u����`�����A�N�e�B�u�ɂ��Ă�������")
    Else
        ActiveSheet.Unprotect
        Columns(C_HideSNm & ":" & C_HideENm).Select     '--- Mod S.Iwanaga 2010/04/08
        Selection.EntireColumn.Hidden = True
        'ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        'False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
        'AllowFormattingRows:=True
    End If
    
End Sub


'==========================================================
'�y�v���V�[�W�����zseisho
'�y�T�@�v�z�e�[�u�����ڏ��Ɍr������������
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================

Sub seisho()
    Dim sname As String
    Dim bname As String
    Dim MaxRow As Integer
    Dim MaxCol As Integer
    Dim DmaxRow As Integer
    Dim i As Integer
    Dim seq As Integer
    Dim wktext As String
    Dim wrange As String
    
    If Workbooks.Count = 0 Then
        Exit Sub
    End If
    bname = ActiveWorkbook.name
    sname = Cells(R_TblId, C_TblId).Value '�V�[�g��
    
    If Cells(R_COLNAME, C_COLNAME).Value = "" Then
        Exit Sub
    End If
    
    
    Application.ScreenUpdating = False '�`����~�߂�
    ActiveSheet.Unprotect
    With Range(Cells(R_COLNAME, 1), Cells(R_COLNAME, 1)).SpecialCells(xlLastCell)
        MaxRow = .Row
        MaxCol = .Column
    End With
    
    Range(Cells(R_COLNAME, 1), Cells(MaxRow, C_KeiEnd)).Select
    
    With Selection
        .Borders.LineStyle = xlNone
    End With
    
    '�f�[�^�����͂���Ă���ŏI�s���擾����
    If Cells(R_COLNAME + 1, C_COLNAME).Value = "" Then
        DmaxRow = R_COLNAME
    Else
        DmaxRow = Range(Cells(R_COLNAME, C_COLNAME), Cells(R_COLNAME, C_COLNAME)).End(xlDown).Row
    End If
    
    '�f�[�^�����͂���Ă���J�����ɐ��������Ȃ���
    Range(Cells(R_COLNAME, 1), Cells(DmaxRow, C_KeiEnd)).Select
    
    With Selection
        .Borders.LineStyle = xlContinuous
    End With
    
    With Range(Cells(R_COLNAME, 1), Cells(R_COLNAME, 1)).SpecialCells(xlLastCell)
        MaxRow = .Row
        MaxCol = .Column
    End With

    '�f�[�^���͈�O�̃f�[�^���폜����
    Range(Cells(DmaxRow + 1, 1), Cells(MaxRow, C_KeiEnd)).Select
    Selection.ClearContents
    With Selection.Validation
        .Delete
    End With

    '���Ԃ�U�蒼��
    seq = 0
    For i = R_COLNAME To DmaxRow
        seq = seq + 1
        Cells(i, 1).Value = seq
        wktext = Cells(i, C_COLNAME).Value
        Cells(i, C_COLNAME).Value = UCase(wktext)
    Next i
    'Range(Cells(1, 1), Cells(DmaxRow, 1)).Select
    'Selection.VerticalAlignment = xlCenter
    'Range(Cells(R_COLNAME, C_COLNAME - 1), Cells(DmaxRow, C_COLNAME)).Select
    'Selection.VerticalAlignment = xlCenter
    'Range(Cells(R_COLNAME, C_kata), Cells(DmaxRow, C_def)).Select
    'Selection.VerticalAlignment = xlCenter
    
    '�e�[�u�����̂ɉ����Z�b�g����Ă��Ȃ��ꍇ�e�[�u��ID���Z�b�g
    wktext = Cells(R_TblNm2, C_TblNm2).Value
    If wktext = "" Then
        Cells(R_TblNm2, C_TblNm2).Value = sname
        Cells(R_TblNm, C_TblNm).Value = sname
    Else
        Cells(R_TblNm, C_TblNm).Value = wktext
    End If
       
    '�쐬���ɉ����Z�b�g����Ă��Ȃ��ꍇ�{�����Z�b�g
    wktext = Cells(R_Create, C_Create).Value
    
    If wktext = "" Then
        Cells(R_Create, C_Create).Value = Format(Date, "yyyy/mm/dd")
    End If
    Rows(R_COLNAME & ":" & DmaxRow).Select
    Selection.RowHeight = 24
    Selection.Font.name = "�l�r �S�V�b�N"
    Selection.Font.FontStyle = "�W��"
    Selection.Font.Size = 9
    Selection.Font.Bold = False
    'With Selection.Font
        '.name = "MS UI Gothic"
        '.FontStyle = "�W��"
        '.Size = 9
        '.Strikethrough = False
        '.Superscript = False
        '.Subscript = False
        '.OutlineFont = False
        '.Shadow = False
        '.Underline = xlUnderlineStyleNone
        '.ColorIndex = xlAutomatic
        '.Bold = False
    'End With
    Columns(C_printsta & ":" & C_printend).Select
    Selection.ColumnWidth = 2
    Range(Cells(1, 1), Cells(DmaxRow, C_KeiEnd)).Select
    wrange = Selection.Address(False, False)
    ActiveSheet.PageSetup.PrintArea = wrange
    Application.ScreenUpdating = True '�`��ĊJ
    Cells(1, 1).Select
    'ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
    '    False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
    '    AllowFormattingRows:=True
End Sub
