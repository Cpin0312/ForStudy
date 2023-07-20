Attribute VB_Name = "DDL_CTL_MAKE"
'***PutDdl()  DDL�o��
'***colname_seach(bname As String, sname As String) As Integer ��������T���ăJ�������Z�b�g����
'***spadd_r(i_st As String, i_len As Integer) As String �E���Ɏw�肳�ꂽ���̃X�y�[�X���Z�b�g����

Sub PutDdlSheet()
    Dim abook As Workbook
    Dim sname As String
    
    'On Error GoTo err1
    Call Tb_Posget
    
    If Workbooks.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
        Exit Sub
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
        Exit Sub
    End If
        
    Set abook = ActiveWorkbook
    If abook.name <> "" Then
        If Cells(R_SheetId, C_SheetId).Value <> 2 Then
            Exit Sub
        End If
        Call MakeDdlSheet
    End If
End Sub

Sub PutDdlCpy()
    Dim abook As Workbook
    Dim sname As String
    
    On Error GoTo err1
    Set abook = ActiveWorkbook
    If abook.name <> "" Then
        If Cells(R_SheetId, C_SheetId).Value <> 2 Then
            Exit Sub
        End If
        Call MakeDdlCpy
    End If
    Exit Sub
err1:
    MsgBox ("�V�[�g��ǉ����Ă�������")
End Sub

' [DDL�o��]->[�t�@�C��]�I����
Sub PutDdlFil()
    Dim abook As Workbook
    Dim sname As String
    Dim tableListRow As Integer
    Dim tableListCol As Integer
    Dim tableName As String
    Dim dialog As FileDialog
    Dim rootFolderPath As String
    Dim retVal As Integer
    
    On Error GoTo err1
    
    Call Tb_Posget
    
    Set abook = ActiveWorkbook
    If abook.name <> "" Then
        
        '-----------------------------
        ' �o�͐惋�[�g�t�H���_�I��
        '-----------------------------
        Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
        retVal = dialog.Show
        
        If retVal = -1 Then
            rootFolderPath = dialog.SelectedItems(1)
            ' ������΍��
            If Dir(rootFolderPath, vbDirectory) = "" Then
                MkDir (rootFolderPath)
            End If
            Set dialog = Nothing
        Else
            Set dialog = Nothing
            Exit Sub
        End If

        
        '-----------------------------
        ' �I���V�[�g�ɂ��Ώۂ�ύX
        '-----------------------------
        sname = abook.ActiveSheet.name
        
        If sname = "�c�a�ꗗ" Then
            
            tableListRow = 5
            tableListCol = 2
            
            ' �J�����Q�ƃ��[�v
            While abook.Worksheets(sname).Cells(tableListRow, tableListCol) <> ""
                ' �e�[�u�����擾
                tableName = abook.Worksheets(sname).Cells(tableListRow, tableListCol)
                
                ' �����̃V�[�g��I��
                abook.Worksheets(tableName).Select
                
'--- MOD Start 2019/06/20 SPC
'                If Cells(R_SheetId, C_SheetId).Value = 2 Then
                ' �u�e�[�u��ID�v���r���[�i"Z_"�n�܂�j�̏ꍇ�I��
                If Cells(R_TblId2, C_TblId2).Value = "" _
                    Or (Len(Cells(R_TblId2, C_TblId2).Value) >= 2 And Left(Cells(R_TblId2, C_TblId2).Value, 2) = "Z_") Then
                Else
                    ' �e�[�u����`�V�[�g�ł����DDL�쐬���s
                    Call MakeDdlFil(rootFolderPath)
                End If
'--- MOD End 2019/06/20 SPC
                
                tableListRow = tableListRow + 1
            Wend
        Else
'--- MOD Start 2019/06/20 SPC
'            If Cells(R_SheetId, C_SheetId).Value <> 2 Then
            ' �u�e�[�u��ID�v���r���[�i"Z_"�n�܂�j�̏ꍇ�I��
            If Cells(R_TblId2, C_TblId2).Value = "" _
                    Or (Len(Cells(R_TblId2, C_TblId2).Value) >= 2 And Left(Cells(R_TblId2, C_TblId2).Value, 2) = "Z_") Then
'--- MOD End 2019/06/20 SPC
                Exit Sub
            End If
            Call MakeDdlFil(rootFolderPath)
        End If
    End If
    
    MsgBox "�����I"
    
    Exit Sub
err1:
    MsgBox ("�V�[�g��ǉ����Ă��������B�F" & tableName)
End Sub

Sub MakeDdlSheet()
    Dim bname As String
    Dim sname As String
    Dim ddl_sname As String
    Dim Dao As DataObject
    Dim wktext As String
    Dim MaxRow As Integer
    
    
    If ActiveWorkbook.ActiveSheet.Cells(R_DocId, C_DocId) <> 1 Then
        MsgBox ("�e�[�u����`�����A�N�e�B�u�ɂ��Ă�������")
        Exit Sub
    End If

    MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
    If MaxRow = 0 Then
        MsgBox ("����ID�ɋ󗓂�����܂�")
        Exit Sub
    End If
    
    
    If Checkspace(R_COLNAME, C_kata, MaxRow) = 0 Then
        MsgBox ("�^�ɋ󗓂�����܂�")
        Exit Sub
    End If
        
    '---MOD START 2010/07/27 OU
    'If Checkspace(R_COLNAME, C_keta, MaxRow) = 0 Then
    If Checkspaceketa(R_COLNAME, C_keta, MaxRow) = 0 Then
    '---MOD END
        MsgBox ("���ɋ󗓂�����܂�")
        Exit Sub
    End If
    
    Call seisho
    bname = ActiveWorkbook.name
    sname = ActiveWorkbook.ActiveSheet.name
'--- Mod Start 2012/03/05 TFC
    wktext = CreateDdl(bname, sname, DDL_KIND_ALL)
'--- Mod End 2012/03/05 TFC
    ddl_sname = "Create_" & sname
    DeleteSheet (ddl_sname)
    Set Dao = New DataObject
    Worksheets().Add
    Workbooks(ActiveWorkbook.name).ActiveSheet.name = ddl_sname
    ActiveWindow.DisplayGridlines = False
    Cells.Select
    With Selection.Font
        .name = "�l�r �S�V�b�N"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
    Dao.SetText wktext
    Dao.PutInClipboard
    Cells(1.1).PasteSpecial
    Cells(1, 1).Select
End Sub

Sub MakeDdlCpy()
    Dim bname As String
    Dim sname As String
    Dim Dao As DataObject
    Dim wktext As String
    Dim MaxRow As Integer
    
    Call Tb_Posget
    
    If ActiveWorkbook.ActiveSheet.Cells(R_DocId, C_DocId) <> 1 Then
        MsgBox ("�e�[�u����`�����A�N�e�B�u�ɂ��Ă�������")
        Exit Sub
    End If

    MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
    If MaxRow = 0 Then
        MsgBox ("����ID�ɋ󗓂�����܂�")
        Exit Sub
    End If
    
    
    If Checkspace(R_COLNAME, C_kata, MaxRow) = 0 Then
        MsgBox ("�^�ɋ󗓂�����܂�")
        Exit Sub
    End If
    
    '---MOD START 2010/07/27 OU
    'If Checkspace(R_COLNAME, C_keta, MaxRow) = 0 Then
    If Checkspaceketa(R_COLNAME, C_keta, MaxRow) = 0 Then
    '---MOD END
        MsgBox ("���ɋ󗓂�����܂�")
        Exit Sub
    End If
    
    Call seisho
    bname = ActiveWorkbook.name
    sname = ActiveWorkbook.ActiveSheet.name
'--- Mod Start 2012/03/05 TFC
    wktext = CreateDdl(bname, sname, DDL_KIND_ALL)
'--- Mod End 2012/03/05 TFC
    Set Dao = New DataObject
    Dao.SetText wktext
    Dao.PutInClipboard
End Sub

Sub MakeDdlFil(rootFolderPath As String)
    Dim bname As String
    Dim sname As String
    Dim Dao As DataObject
    Dim wktext As String
    Dim strFilePath As String
    Dim intFileNo As Integer
    Dim MaxRow As Integer
    Dim folderPath As String
    Dim strTableId As String
    
    Call Tb_Posget
    
'--- DEL Start 2019/06/20 SPC
'    If ActiveWorkbook.ActiveSheet.Cells(R_DocId, C_DocId) <> 1 Then
'        MsgBox ("�e�[�u����`�����A�N�e�B�u�ɂ��Ă�������")
'        Exit Sub
'    End If
'--- DEL Start 2019/06/20 SPC

    MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
    If MaxRow = 0 Then
        MsgBox ("����ID�ɋ󗓂�����܂�")
        Exit Sub
    End If
    
    
    If Checkspace(R_COLNAME, C_kata, MaxRow) = 0 Then
        MsgBox ("�^�ɋ󗓂�����܂�")
        Exit Sub
    End If
        
    '---MOD START 2010/07/27 OU
    'If Checkspace(R_COLNAME, C_keta, MaxRow) = 0 Then
    If Checkspaceketa(R_COLNAME, C_keta, MaxRow) = 0 Then
    '---MOD END
        MsgBox ("���ɋ󗓂�����܂�")
        Exit Sub
    End If
    
'--- DEL Start 2019/06/20 SPC
'    Call seisho
'--- DEL Start 2019/06/20 SPC
    
    ' �e�[�u��ID�ێ�
    strTableId = Cells(R_TblId2, C_TblId2).Value
    
    bname = ActiveWorkbook.name
    sname = ActiveWorkbook.ActiveSheet.name
    
    ' ----------------------------------
    ' �e�[�u����`�o��
    ' ----------------------------------
    wktext = CreateDdl(bname, sname, DDL_KIND_TABLE)
    
    '�ۑ���̎擾
    folderPath = rootFolderPath & "\table"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir (folderPath)
    End If
    
    strFilePath = folderPath & "\" & strTableId & ".sql"
    
    '�t�@�C���o��
'--- MOD Start 2019/07/08 SPC
'    intFileNo = FreeFile
'    Open strFilePath For Output As #intFileNo
'    Print #intFileNo, wktext
'    Close #intFileNo
    Call outputUtf8File(strFilePath, wktext)
'--- MOD End 2019/07/08 SPC
    
    ' ----------------------------------
    ' PK�A�C���f�b�N�X��`�o��
    ' ----------------------------------
    wktext = CreateDdl(bname, sname, DDL_KIND_INDEX)
    
    If wktext <> "" Then
        '�ۑ���̎擾
        folderPath = rootFolderPath & "\index"
        
        If Dir(folderPath, vbDirectory) = "" Then
            MkDir (folderPath)
        End If
        
        strFilePath = folderPath & "\" & strTableId & ".sql"
        
        '�t�@�C���o��
'--- MOD Start 2019/07/08 SPC
'        intFileNo = FreeFile
'        Open strFilePath For Output As #intFileNo
'        Print #intFileNo, wktext
'        Close #intFileNo
        Call outputUtf8File(strFilePath, wktext)
'--- MOD Start 2019/07/08 SPC
        
    End If
    
End Sub

'==========================================================
'�y�v���V�[�W�����zPutDdl
'�y�T�@�v�zDDL���o��
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================

Sub PutDdl()
    Dim wkmsg As String
    Dim wknum As Integer
    Dim ans As Integer
    Dim tbl_sname As String
    Dim sname As String
    Dim ddl_sname As String
    Dim tableName As String
    Dim wk_len As Integer
    Dim x As Integer
    Dim y As Integer
    Dim curr_l As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim TABLENAMEJ As String
    Dim pkey(32768) As Integer
    Dim pkey_seq(32768) As Integer
    Dim pkey_pos(32768) As Integer
    Dim wktablesp As String
    Dim tblsheet As Worksheet
    Dim wktype As String
    
    Tb_Posget
    seisho
    

    If Cells(R_COLNAME, C_COLNAME) = "" Then
        MsgBox ("����ID�����͂���Ă��܂���")
        Exit Sub
    End If

    tbl_sname = Workbooks(ActiveWorkbook.name).ActiveSheet.Cells(R_TblId, C_TblId)
    ddl_sname = "CREATE��_" & tbl_sname
    DeleteSheet (ddl_sname)
    Worksheets().Add
    Workbooks(ActiveWorkbook.name).ActiveSheet.name = ddl_sname
    sname = Workbooks(ActiveWorkbook.name).ActiveSheet.name
    Workbooks(ActiveWorkbook.name).Worksheets(sname).name = ddl_sname
    ActiveWindow.DisplayGridlines = False

    
    Application.ScreenUpdating = False
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Select
    Cells.Select
    With Selection.Font
        .name = "�l�r �S�V�b�N"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
        
    '�R�����g�s1�s��
    x = 1
    y = 1
    Cells(y, x).Value = "/**********************************************************/"
    
    '�R�����g�s�@TABLENAME
    wkmsg = DdlCom_TbId(tbl_sname)
    
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    y = y + 1
    Cells(y, x).Value = wkmsg
    
    '�R�����g�s�@TABLENAME
    wkmsg = DdlCom_TbNm(tbl_sname)
    
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    y = y + 1
    Cells(y, x).Value = wkmsg
    
    wkmsg = "/*     " & "�쐬��:" & Format(Date, "yyyy/mm/dd")
    
    wk_len = LenB(StrConv(wkmsg, vbFromUnicode))
    wknum = 57 - wk_len
    wkmsg = spadd_r(wkmsg, wknum)
    wkmsg = wkmsg & " */"
    
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    y = y + 1
    Cells(y, x).Value = wkmsg
    
'--- DEL Start 2015/02/27 TFC
'    y = y + 1
'    Cells(y, x).Value = "/**********************************************************/"
'    y = y + 1
'    Cells(y, x).Value = "/* �G���[�n���h�����O */"
'    y = y + 1
'    Cells(y, x).Value = "WHENEVER OSERROR  EXIT OSCODE      ROLLBACK"
'    y = y + 1
'    Cells(y, x).Value = "WHENEVER SQLERROR EXIT SQL.SQLCODE ROLLBACK"
'--- DEL End

    y = y + 1
    Cells(y, x).Value = "/* CREATE �� */"
    
    tableName = Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Cells(R_TblId2, C_TblId2).Value
    wkschima = Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Cells(R_Schima, C_Schima).Value
    If wkschima = "" Then
        wkmsg = "CREATE TABLE " & tableName & "("
    Else
        wkmsg = "CREATE TABLE " & wkschima & "." & tableName & "("
    End If
        
    y = y + 1
    Cells(y, x).Value = wkmsg
    curr_l = R_COLNAME '������1�s��
    
    '���ږ��𕨗����J��������擾����DDL�����쐬����
    While Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Cells(curr_l, C_COLNAME) <> ""
        Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Activate
        wkmsg = "       " & Cells(curr_l, C_COLNAME)
        wkmsg = wkmsg & " " & Cells(curr_l, C_kata)
        wkmsg = wkmsg & "(" & Cells(curr_l, C_keta)
        If Cells(curr_l, C_shou) <> "" Then
           If Cells(curr_l, C_kata) = "NUMBER" Then
                wkmsg = wkmsg & "," & Cells(curr_l, C_shou)
            End If
        End If
        wkmsg = wkmsg & ")"
        If Cells(curr_l, C_nnul).Value = "��" Then
            wkmsg = wkmsg & " NOT NULL"
        End If
        If Cells(curr_l, C_uniq).Value = "��" Then
            wkmsg = wkmsg & " UNIQUE"
        End If
        
        wktype = Cells(curr_l, C_kata).Value
        If Cells(curr_l, C_def).Value <> "" Then
            wkmsg = wkmsg & " DEFAULT "
            If (wktype = "CHAR") Or (wktype = "VARCHAR2") Then
                wkmsg = wkmsg & " '" & Cells(curr_l, C_def) & "'"
            Else
                wkmsg = wkmsg & Cells(curr_l, C_def)
            End If
        End If
        
        If Cells(curr_l + 1, C_COLNAME) <> "" Then
            wkmsg = wkmsg & ","
        End If
        
        
        Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
        y = y + 1
        Cells(y, x).Value = wkmsg
        curr_l = curr_l + 1
    Wend
    
    '�Ō�́h�j�h��ǉ�����
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    wkmsg = "       )"
    wktablesp = Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Cells(R_TblSp, C_TblSp).Value
    If wktablesp <> "" Then
        wkmsg = wkmsg & " TABLESPACE "
        wkmsg = wkmsg & wktablesp
        wkmsg = wkmsg & ";"
    Else
        wkmsg = wkmsg & ";"
    End If
    y = y + 1
    Cells(y, x).Value = wkmsg

        
    '��L�[�����ǉ�����
    Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Activate
    j = 0
    For i = R_COLNAME To curr_l - 1
        If Cells(i, C_primary) <> "" Then
            pkey(j) = CInt(Cells(i, C_primary))
            pkey_pos(j) = i
            pkey_seq(j) = j + 1
            j = j + 1
        End If
    Next i
    
    wkpkey = ""
    If j > 0 Then
        For i = 0 To j - 1
            For k = 0 To j - 1
                If pkey_seq(i) = pkey(k) Then
                    If i < j - 1 And j > 0 Then
                        wkpkey = wkpkey & Cells(pkey_pos(k), C_COLNAME) & ","
                    Else
                        wkpkey = wkpkey & Cells(pkey_pos(k), C_COLNAME)
                    End If
                End If
            Next k
        Next i
    End If
    
    If wkpkey <> "" Then
        Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
        y = y + 1
        Cells(y, x).Value = "/* PRIMARY KEY */"
        y = y + 1
        Cells(y, x).Value = "ALTER TABLE " & tableName
        wkmsg = " ADD CONSTRAINT " & "PK_" & tableName & " PRIMARY KEY(" & wkpkey & ")"
        
        'TABLE�\�̈悪�w�肳��Ă���Βǉ�����
        Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Activate
        If Cells(pkey_pos(0), C_IdxSp2).Value <> "" Then
        
            Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
            y = y + 1
            Cells(y, x).Value = wkmsg
            
            Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Activate
            wktablesp = " USING INDEX TABLESPACE " & Cells(pkey_pos(0), C_IdxSp2)
            
            Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
            y = y + 1
            Cells(y, x).Value = wktablesp & ";"
        Else
            Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
            wkmsg = wkmsg & ";"
            y = y + 1
            Cells(y, x).Value = wkmsg
        End If
            
    End If
    
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    Application.ScreenUpdating = True
    Range("A1").Select
End Sub

'==========================================================
'�y�֐����zDdlCom_TbId
'�y�T�@�v�z�e�[�u������ҏW
'�y���@���z�V�[�g��
'�y�߂�l�z�R�����g��
'==========================================================

Function DdlCom_TbId(sname As String) As String
    Dim tableName As String
    Dim wkmsg As String
    Dim wknum As Integer
    
    tableName = Workbooks(ActiveWorkbook.name).Worksheets(sname).Cells(R_TblId2, C_TblId2).Value
    wkmsg = "/*     TABLE NAME: " & tableName
    wk_len = Len(wkmsg)
    wknum = 57 - wk_len
    wkmsg = spadd_r(wkmsg, wknum)
    wkmsg = wkmsg & " */"
    
    DdlCom_TbId = wkmsg
    
End Function

'==========================================================
'�y�֐����zDdlCom_TbNm
'�y�T�@�v�z�e�[�u������ҏW
'�y���@���z�V�[�g��
'�y�߂�l�z�R�����g��
'==========================================================
Function DdlCom_TbNm(sname As String) As String
    Dim wkmsg As String
    Dim TABLENAMEJ As String
    Dim wk_len As Integer
    Dim wknum As Integer
    
    TABLENAMEJ = Workbooks(ActiveWorkbook.name).Worksheets(sname).Cells(R_TblNm, C_TblNm).Value
    wkmsg = "/*     " & TABLENAMEJ
    wk_len = LenB(StrConv(wkmsg, vbFromUnicode))
    
    wknum = 57 - wk_len
    wkmsg = spadd_r(wkmsg, wknum)
    wkmsg = wkmsg & " */"
    
    DdlCom_TbNm = wkmsg
    
End Function




'==========================================================
'�y�֐����zspadd_r
'�y�T�@�v�zString�̉E���ɔ��pSpace���l�߂�
'�y���@���z�l�߂�O��String,�l�߂镶����
'�y�߂�l�z�l�߂����String
'==========================================================

Function spadd_r(i_st As String, i_len As Integer) As String
    Dim i As Integer
    Dim wk_st As String
    wk_st = i_st
    For i = 1 To i_len
        wk_st = wk_st & " "
    Next i
    spadd_r = wk_st
End Function


'--- Add Start S.Iwanaga 2010/04/08
'==========================================================
'�y�֐����zmenuExAreaView
'�y�T�@�v�z�g���J�����\��/��\������
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================
Public Sub menuExAreaView()

    Dim strSheet    As String
    
    If Workbooks.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name
        
        '�V�[�g�^�C�v���e�[�u�����ڃV�[�g�̏ꍇ�̂ݏ���
        If isTblDefSheet(strSheet) Then
            '�g���J�����̕\����Ԏ擾
            If chkExAreaState(strSheet) = 0 Then
                '��\��
                Call kaktyoV
            Else
                '�\��
                Call kaktyoNV
            End If
        End If
    End If
        
End Sub

'==========================================================
'�y�֐����zmenuConvLogicName
'�y�T�@�v�z���j���[�������ϊ�����
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================
Public Sub menuConvLogicName()

    Dim strSheet    As String
    
    If Workbooks.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name
    
        '---Mod Start OU 2010/07/27
        'If LtoP(strSheet, 8, 20) = -1 Then
        If LtoP(strSheet, R_COLNAME, 70) = -1 Then
        '---Mod End
            '�ϊ������G���[
        End If
    End If
    
End Sub
'--- Add End

'--- Add Start S.Iwanaga 2010/04/16
'==========================================================
'�y�֐����zmenuCtlToFile
'�y�T�@�v�zCTL���t�@�C���ɏo��
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================
Public Sub menuCtlToFile()

    Dim intRtn      As Integer
    Dim intFileNo   As Integer
    Dim strSheet    As String
    Dim strFilePath As String
    Dim strCtlData  As String
    Dim MaxRow As Integer
    
    strCtlData = ""
    strFilePath = ""
    
    If Workbooks.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name

        '�e�[�u�����ڃV�[�g���`�F�b�N
        If Not isTblDefSheet(strSheet) Then
            MsgBox "�e�[�u�����ڃV�[�g���A�N�e�B�u�ɂ��ĉ������B", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        
        MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
        If MaxRow = 0 Then
            MsgBox ("�����͂̍���ID�Z�������݂��܂�")
            Exit Sub
        End If
        
        'Add Start 2010/07/29 OU
        Dim strDataType As String
        strDataType = Trim(Cells(R_DataTyp, C_DataTyp).Value)
        If strDataType <> "CSV" Then
        'Add End
            If Checkspace(R_COLNAME, C_FFilePosition, MaxRow) = 0 Then
                MsgBox ("�����͂̃t���b�g�t�@�C���ʒu�Z�������݂��܂�")
                kaktyoV
                Exit Sub
            End If
            
            If Checkspace(R_COLNAME, C_FFileLength, MaxRow) = 0 Then
                MsgBox ("�����͂̃t���b�g�t�@�C�����Z�������݂��܂�")
                kaktyoV
                Exit Sub
            End If
        End If
        
        '�����������̓`�F�b�N
        intRtn = chkBlankPName(strSheet)
        If intRtn = 1 Then
            MsgBox "�����͂̍���ID�Z�������݂��܂��B", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        
        '�ۑ���̎擾
        strFilePath = Application.GetSaveAsFilename(strSheet & ".ctl", "����t�@�C��, *.ctl")
        
        '����t�@�C���쐬
        strCtlData = createCtl(strSheet)

        If Len(strCtlData) > 0 Then
            '�t�@�C���o��
            intFileNo = FreeFile
            Open strFilePath For Output As #intFileNo
            Print #intFileNo, strCtlData
            Close #intFileNo
        End If
    End If

End Sub

'==========================================================
'�y�֐����zmenuCtlToSheet
'�y�T�@�v�zCTL���V�[�g�ɏo��
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================
Public Sub menuCtlToSheet()

    Dim strSheet        As String
    Dim strCtlData      As String
    Dim strTableName    As String
    Dim strCtlShtName   As String
    Dim dd              As New DataObject
    Dim ws              As Worksheet
    Dim MaxRow As Integer
    
    strCtlData = ""
    
    If Workbooks.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("�e�[�u����`�����J���Ă�������")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name

        '�e�[�u�����ڃV�[�g���`�F�b�N
        If Not isTblDefSheet(strSheet) Then
            MsgBox "�e�[�u�����ڃV�[�g��I�����ĉ������B", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        
        MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
        If MaxRow = 0 Then
            MsgBox ("�����͂̍���ID�Z�������݂��܂�")
            Exit Sub
        End If
        
        'Add Start 2010/07/29 OU
        Dim strDataType As String
        strDataType = Trim(Cells(R_DataTyp, C_DataTyp).Value)
        If strDataType <> "CSV" Then
        'Add End
            If Checkspace(R_COLNAME, C_FFilePosition, MaxRow) = 0 Then
                MsgBox ("�����͂̃t���b�g�t�@�C���ʒu�Z�������݂��܂�")
                kaktyoV
                Exit Sub
            End If
            
            If Checkspace(R_COLNAME, C_FFileLength, MaxRow) = 0 Then
                MsgBox ("�����͂̃t���b�g�t�@�C�����Z�������݂��܂�")
                kaktyoV
                Exit Sub
            End If
        End If
        
       '�����������̓`�F�b�N
        intRtn = chkBlankPName(strSheet)
        If intRtn = 1 Then
            MsgBox "�����͂̍���ID�Z�������݂��܂��B", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
                
        '����t�@�C���쐬
        strCtlData = createCtl(strSheet)

        dd.SetText strCtlData
        dd.PutInClipboard

        strTableName = Trim(ActiveWorkbook.Sheets(strSheet).Cells(R_TblId, C_TblId).Value)
        strCtlShtName = "CTL_" & strTableName
        Call DeleteSheet(strCtlShtName)
        Set ws = ActiveWorkbook.Worksheets.Add()
        ws.name = strCtlShtName
        ws.Cells(1, 1).PasteSpecial
        'Add Start 2010/07/30 OU
        ActiveWindow.DisplayGridlines = False
        Cells(1, 1).Select
        'Add End
    End If

End Sub

'==========================================================
'�y�֐����zmenuCtlToClipboard
'�y�T�@�v�zCTL���N���b�v�{�[�h�ɃR�s�[
'�y���@���z�Ȃ�
'�y�߂�l�z�Ȃ�
'==========================================================
Public Sub menuCtlToClipboard()

    Dim strSheet    As String
    Dim strCtlData  As String
    Dim dd          As New DataObject
    Dim MaxRow As Integer
    
    strCtlData = ""
    
    If Workbooks.Count < 1 Then
        MsgBox ("�L���ȃe�[�u����`�����J���Ă�������")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("�L���ȃe�[�u����`�����J���Ă�������")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name

        '�e�[�u�����ڃV�[�g���`�F�b�N
        If Not isTblDefSheet(strSheet) Then
            MsgBox "�e�[�u�����ڃV�[�g��I�����ĉ������B", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        
        
        MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
        If MaxRow = 0 Then
            MsgBox ("�����͂̍���ID�Z�������݂��܂�")
            Exit Sub
        End If
        
        'Add Start 2010/07/29 OU
        Dim strDataType As String
        strDataType = Trim(Cells(R_DataTyp, C_DataTyp).Value)
        If strDataType <> "CSV" Then
        'Add End
            If Checkspace(R_COLNAME, C_FFilePosition, MaxRow) = 0 Then
                MsgBox ("�����͂̃t���b�g�t�@�C���ʒu�Z�������݂��܂�")
                kaktyoV
                Exit Sub
            End If
            
            If Checkspace(R_COLNAME, C_FFileLength, MaxRow) = 0 Then
                MsgBox ("�����͂̃t���b�g�t�@�C�����Z�������݂��܂�")
                kaktyoV
                Exit Sub
            End If
        
        End If
        
        
        '�����������̓`�F�b�N
        intRtn = chkBlankPName(strSheet)
        If intRtn = 1 Then
            MsgBox "�����͂̍���ID�Z�������݂��܂��B", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
                
        '����t�@�C���쐬
        strCtlData = createCtl(strSheet)

        dd.SetText strCtlData
        dd.PutInClipboard

    End If

End Sub
'--- Add End

'--- ADD Start 2019/07/08 SPC
Sub outputUtf8File(filePath As String, wtext As String)

    Dim stream As ADODB.stream
    Set stream = New ADODB.stream
    Dim byteData() As Byte
    
    stream.Type = adTypeText
    stream.Charset = "UTF-8"
    stream.LineSeparator = adCRLF
    stream.Open

    stream.WriteText wtext, adWriteLine
    
    '----------------
    ' BOM�폜
    '----------------
    ' BOM�R�[�h���΂��ăf�[�^���擾����B
    stream.Position = 0
    stream.Type = adTypeBinary
    stream.Position = 3
    byteData = stream.Read
    
    ' �擾�����f�[�^��擪���珑���o������
    stream.Position = 0
    stream.Write byteData
    stream.SetEOS
    
    ' �t�@�C���ۑ�
    stream.SaveToFile filePath, adSaveCreateOverWrite
    stream.Close
    
End Sub
'--- ADD End 2019/07/08 SPC


