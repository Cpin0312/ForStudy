Attribute VB_Name = "Global_PosGet"

'==========================================================
'�y�v���V�[�W�����zTb_Posget
'�y�T�@�v�z�e�[�u����`���̃J�����|�W�V�������擾
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================

Sub Tb_Posget()
    Dim bname As String
    Dim sname As String
    '--- MOD Start 2015/02/27 TFC
    bname = ThisWorkbook.name
    '--- MOD End 2015/02/27 TFC
    sname = Workbooks(bname).Worksheets("�v���p�e�B").name
    Tb_SheetNm = CStr(Workbooks(bname).Worksheets(sname).Cells(4, 5).Value) '�e�[�u�����ځ@�V�[�g��
    R_COLNAME = CInt(Workbooks(bname).Worksheets(sname).Cells(9, 8).Value) '�e�[�u���������@��     ����ID
    C_COLNAME = CInt(Workbooks(bname).Worksheets(sname).Cells(10, 8).Value) '�e�[�u���������@�J����   ����ID
    R_ITEMNAME = CInt(Workbooks(bname).Worksheets(sname).Cells(56, 8).Value) '�e�[�u�����ږ��@�s  ���ږ��J����POS
    C_ITEMNAME = CInt(Workbooks(bname).Worksheets(sname).Cells(57, 8).Value) '�e�[�u�����ږ��@�J����   ���ږ��J����POS
    R_TblId = CInt(Workbooks(bname).Worksheets(sname).Cells(3, 8).Value) '�e�[�u��ID�ʒu(�B��)
    C_TblId = CInt(Workbooks(bname).Worksheets(sname).Cells(4, 8).Value) '�e�[�u��ID�ʒu(�B��)
    R_TblNm = CInt(Workbooks(bname).Worksheets(sname).Cells(6, 8).Value) '�e�[�u�����ʒu(�B��)
    C_TblNm = CInt(Workbooks(bname).Worksheets(sname).Cells(7, 8).Value) '�e�[�u�����ʒu(�B��)
    C_KeiEnd = CInt(Workbooks(bname).Worksheets(sname).Cells(13, 8).Value)
    C_HideSta = CInt(Workbooks(bname).Worksheets(sname).Cells(16, 8).Value)
    C_HideEnd = CInt(Workbooks(bname).Worksheets(sname).Cells(16, 8).Value)
    R_TblId2 = CInt(Workbooks(bname).Worksheets(sname).Cells(47, 8).Value) '�e�[�u��ID�ʒu(���o���j
    C_TblId2 = CInt(Workbooks(bname).Worksheets(sname).Cells(48, 8).Value) '�e�[�u��ID�ʒu(���o���j
    R_Schima = CInt(Workbooks(bname).Worksheets(sname).Cells(21, 8).Value) '�X�L�[�}��
    C_Schima = CInt(Workbooks(bname).Worksheets(sname).Cells(22, 8).Value) '�X�L�[�}��
    R_TblSp = CInt(Workbooks(bname).Worksheets(sname).Cells(24, 8).Value) '�e�[�u���\�̈�
    C_TblSp = CInt(Workbooks(bname).Worksheets(sname).Cells(25, 8).Value) '�e�[�u���\�̈�
    R_DataTyp = CInt(Workbooks(bname).Worksheets(sname).Cells(27, 8).Value) '�f�[�^�^�C�v
    C_DataTyp = CInt(Workbooks(bname).Worksheets(sname).Cells(28, 8).Value) '�f�[�^�^�C�v
    R_LdOp = CInt(Workbooks(bname).Worksheets(sname).Cells(30, 8).Value) '���[�h�I�v�V����
    C_LdOp = CInt(Workbooks(bname).Worksheets(sname).Cells(31, 8).Value) '���[�h�I�v�V����
    R_IdxSp = CInt(Workbooks(bname).Worksheets(sname).Cells(33, 8).Value) 'INDEX�\�̈�
    C_IdxSp = CInt(Workbooks(bname).Worksheets(sname).Cells(34, 8).Value) 'INDEX�\�̈�
    R_Create = CInt(Workbooks(bname).Worksheets(sname).Cells(53, 8).Value) '�쐬��
    C_Create = CInt(Workbooks(bname).Worksheets(sname).Cells(54, 8).Value) '�쐬��
    R_TblNm2 = CInt(Workbooks(bname).Worksheets(sname).Cells(50, 8).Value) '�e�[�u�����ʒu(���o���j
    C_TblNm2 = CInt(Workbooks(bname).Worksheets(sname).Cells(51, 8).Value) '�e�[�u�����ʒu(���o���j
    C_printsta = Trim(Workbooks(bname).Worksheets(sname).Cells(63, 8).Value) '����͈͗�
    C_printend = Trim(Workbooks(bname).Worksheets(sname).Cells(64, 8).Value) '����͈͗�
    '--- Add Start OU 2010/07/29
    R_DirectPath = CInt(Workbooks(bname).Worksheets(sname).Cells(66, 8).Value) '�_�C���N�g�p�X
    C_DirectPath = CInt(Workbooks(bname).Worksheets(sname).Cells(67, 8).Value) '�_�C���N�g�p�X
    C_IndexStart = CInt(Workbooks(bname).Worksheets(sname).Cells(69, 8).Value) 'Index�L�[�J�n��
    C_IndexEnd = CInt(Workbooks(bname).Worksheets(sname).Cells(71, 8).Value) 'Index�L�[�ŏI��
    '--- Add End
    '--- ADD Start 2019/07/19 SPC
    R_PartitionKind = CInt(Workbooks(bname).Worksheets(sname).Cells(73, 8).Value) ' �p�[�e�B�V������ލs
    C_PartitionKind = CInt(Workbooks(bname).Worksheets(sname).Cells(74, 8).Value) ' �p�[�e�B�V������ޗ�
    R_PartitionKoumoku = CInt(Workbooks(bname).Worksheets(sname).Cells(76, 8).Value) ' �p�[�e�B�V�����Ώۍ��ڍs
    C_PartitionKoumoku = CInt(Workbooks(bname).Worksheets(sname).Cells(77, 8).Value) ' �p�[�e�B�V�����Ώۍ��ڗ�
    '--- ADD End 2019/07/19 SPC

    '--- Add Start S.Iwanaga 2010/04/08
    '�h�L�������gID�ʒu���擾
    R_DocId = CInt(Workbooks(bname).Worksheets(sname).Cells(2, 11).Value)
    C_DocId = CInt(Workbooks(bname).Worksheets(sname).Cells(3, 11).Value)
    '�V�[�gID�ʒu���擾
    R_SheetId = CInt(Workbooks(bname).Worksheets(sname).Cells(5, 11).Value)
    C_SheetId = CInt(Workbooks(bname).Worksheets(sname).Cells(6, 11).Value)
    '��\���J�����J�n�I���񖼏��擾
    C_HideSNm = Trim(Workbooks(bname).Worksheets(sname).Cells(18, 8).Value)
    C_HideENm = Trim(Workbooks(bname).Worksheets(sname).Cells(19, 8).Value)
    '--- Add End
    '�������ϊ��e�[�u���t�@�C���p�X
    ConvFilePath = Trim(Workbooks(bname).Worksheets(sname).Cells(1, 14).Value)  '--- Add S.Iwanaga 2010/04/13

    '--- Add Start S.Iwanaga 2010/04/16
    '�t���b�g�t�@�C���ʒu
    C_FFilePosition = Trim(Workbooks(bname).Worksheets(sname).Cells(59, 8).Value)
    '�t���b�g�t�@�C����
    C_FFileLength = Trim(Workbooks(bname).Worksheets(sname).Cells(61, 8).Value)
    '--- Add End

    '�e�[�u����`���e���ڌ��o��������e���Z�b�g����ׂ̃J�����ʒu���擾����
    C_kata = colposget("�^")
    C_keta = colposget("����")
    C_shou = colposget("����")
    C_primary = colposget("��L�[")
    C_uniq = colposget("���")
    C_nnul = colposget("�K�{")
    C_check = colposget("�`�F�b�N����")
    C_def = colposget("�f�t�H���g�l")
    '�C���f�b�N�X�\�̈悾���͍s���Ⴄ�̂Ŋ֐����g��Ȃ�
    For i = C_COLNAME To C_KeiEnd
        If Workbooks(bname).Worksheets("�e�[�u������").Cells(R_COLNAME - 1, i).Value = "�\�̈�" Then
            C_IdxSp2 = i
            Exit For
        End If
    Next i
End Sub


