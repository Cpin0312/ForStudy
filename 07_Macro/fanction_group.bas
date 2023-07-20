Attribute VB_Name = "fanction_group"
'Function
'Checkspace(i_R As Integer, i_C As Integer, chk As Integer) As Integer
'HaveSheet(bname As String, sname As String) As Integer
'HaveSheet2(bname As String, sname As Integer) As Integer
'colposget(name As String) As Integer
'CreateDdl(bname As String, sname As String) As String
'Checkphy(bname As String, sname As String) As Integer
'HaveSheet(bname As String, sname As String) As Integer

'==========================================================
'�y�֐����zCheckSpace
'�y�T�@�v�z�J�����̋󗓂��`�F�b�N
'�y���@���zi_R     :Row
'�y���@���zi_C     :Column
'�y�߂�l�z0:�󗓂��� n:�ŏI�s
'==========================================================

Function Checkspace(i_R As Integer, i_C As Integer, chk As Integer) As Integer
    Dim MaxRow As Integer
    Dim DmaxRow As Integer
    If Cells(i_R, i_C) = "" Then
        Checkspace = 0
        Exit Function
    End If
    If chk = 0 Then
        MaxRow = Cells(Rows.Count, i_C).End(xlUp).Row
    Else
        MaxRow = chk
    End If
    
    If Cells(i_R + 1, i_C).Value = "" Then
        DmaxRow = i_R
    Else
        DmaxRow = Range(Cells(i_R, i_C), Cells(i_R, i_C)).End(xlDown).Row
    End If
        
    If MaxRow <> DmaxRow Then
        Checkspace = 0
    Else
        Checkspace = MaxRow
    End If
    
End Function

'---ADD START 2010/07/27 OU
'==========================================================
'�y�֐����zCheckspaceketa
'�y�T�@�v�z�����J�����̋󗓂��`�F�b�N
'�y���@���zi_R     :Row
'�y���@���zi_C     :Column
'�y�߂�l�z0:�󗓂��� n:�ŏI�s
'==========================================================
Function Checkspaceketa(i_R As Integer, i_C As Integer, chk As Integer) As Integer
    Dim MaxRow As Integer

    If chk = 0 Then
        MaxRow = Cells(Rows.Count, i_C).End(xlUp).Row
    Else
        MaxRow = chk
    End If
    
    Dim tmp As String
    Dim i_rl As Integer
    For i_rl = i_R To MaxRow
        If Cells(i_rl, C_kata).Value = "DATE" Or Cells(i_rl, C_kata).Value = "TIMESTAMP" Or Cells(i_rl, C_kata).Value = "BLOB" _
            Or Cells(i_rl, C_kata).Value = "INTEGER" Or Cells(i_rl, C_kata).Value = "BYTEA" Then
            Cells(i_rl, i_C).Value = ""
            Cells(i_rl, C_shou).Value = ""
        ElseIf Cells(i_rl, i_C).Value = "" Then
            Checkspaceketa = 0
            Exit Function
        End If
    Next
    
    Checkspaceketa = MaxRow
End Function
'--- ADD END

'==========================================================
'�y�֐����zHaveSheet
'�y�T�@�v�z�����̃V�[�g��book��ɂ��邩���ׂ�
'�y���@���zbname     :book��
'�y���@���zsname     :sheet��
'�y�߂�l�zTrue=sheet Index�l
'==========================================================
Function HaveSheet(bname As String, sname As String) As Integer
    Dim wnum As Integer        '�o�^���[�N�V�[�g��
    Dim i As Integer           '�J�E���^
    Dim ret As Integer
    Dim bookname As String
    
    ret = 0
    wnum = Workbooks(bname).Worksheets.Count     '�o�^���[�N�V�[�g���𓾂�
    For i = 1 To wnum
        If UCase(Workbooks(bname).Worksheets(i).name) = UCase(sname) Then
            ret = i
            Exit For
        End If
    Next i
    HaveSheet = ret
End Function

'==========================================================
'�y�֐����zHaveSheet2
'�y�T�@�v�z�����V�[�gid��book��ɂ��邩���ׂ�
'�y���@���zbname     :book��
'�y���@���zsname     :sheetid
'�y�߂�l�zTrue=sheet Index�l
'==========================================================
Function HaveSheet2(bname As String, sname As Integer) As Integer
    Dim wnum As Integer        '�o�^���[�N�V�[�g��
    Dim i As Integer           '�J�E���^
    Dim ret As Integer
    Dim bookname As String
    
    ret = 0
    wnum = Workbooks(bname).Worksheets.Count     '�o�^���[�N�V�[�g���𓾂�
    For i = 1 To wnum
        If CInt(Workbooks(bname).Worksheets(i).Cells(R_SheetId, C_SheetId)) = sname Then
            ret = i
            Exit For
        End If
    Next i
    HaveSheet2 = ret
End Function


'==========================================================
'�y�֐����zcolposget
'�y�T�@�v�z�e�[�u�����ڃV�[�g�̃J�����ʒu���擾
'�y���@���z���ږ�
'�y�߂�l�z�J�����ʒu
'==========================================================

Function colposget(name As String) As Integer
    Dim i As Integer
    For i = C_COLNAME To C_KeiEnd
        If ThisWorkbook.Worksheets("�e�[�u������").Cells(R_COLNAME - 2, i).Value = name Then
            colposget = i
            Exit For
        End If
    Next i
    
End Function


'==========================================================
'�y�֐����zCreateDdl
'�y�T�@�v�z�N���G�C�g�����쐬����
'�y���@���z�u�b�N���́A�V�[�g���́A
'          ��ʁiDDL_KIND_ALL�F�e�[�u�����C���f�b�N�X
'                DDL_KIND_TABLE�F�e�[�u����`�̂�
'                DDL_KIND_INDEX�F�C���f�b�N�X��`�̂݁j
'�y�߂�l�z�N���G�C�g��
'==========================================================
Function CreateDdl(bname As String, sname As String, ddlKind As String) As String
    Dim ddl_sname As String
    Dim MaxRow As Integer
    Dim Dao As DataObject
    Dim wktext As String
    Dim wkintext As String
    Dim wkset As String
    Dim curr_l
    Dim j As Integer
    Dim pkey(32768) As Integer
    Dim pkey_seq(32768) As Integer
    Dim pkey_pos(32768) As Integer
    Dim wkpkey As String
    '--- Add Start 2012/02/27 TFC
    Dim idx() As Integer
    Dim idx_seq() As Integer
    Dim idx_pos() As Integer
    '--- Add End 2012/02/27 TFC
    Dim TableId As String
    Dim wktype As String
    '--- Add Start 2010/07/30 OU
    Dim strComment As String
    Dim strCommentTable As String
    strComment = "/* COMMENT */" & vbCrLf
    Cells(R_TblSp, C_TblSp) = UCase(Cells(R_TblSp, C_TblSp))
    Cells(R_IdxSp, C_IdxSp) = UCase(Cells(R_IdxSp, C_IdxSp))
    '--- Add End
    '--- Add End 2012/02/27 TFC
    Dim pkeyName As String
    Dim strindex As String
    Dim intIndexC As Integer
    Dim intIndexR As Integer
    Dim strIndexSpace As String
    Dim indexNamePrefix As String
    Dim indexName As String
    Dim fileHeader As String
    '--- Add End 2012/02/27 TFC
    '--- ADD Start 2019/07/19 SPC
    Dim wkPartitionKind As String
    Dim wkPartitionKoumoku As String
    '--- ADD End 2019/07/19 SPC
    
    fileHeader = "/**********************************************************/" + vbCrLf
    '--- Mod Start 2010/09/21
    'TableId = ActiveWorkbook.Worksheets(sname).Cells(R_TblId, C_TblId).Value
    TableId = Cells(R_TblId2, C_TblId2).Value
    '--- Mod End
    wkset = "/*     TABLE NAME: " & TableId
    wkset = spadd_r(wkset, 57 - Len(wkset))
    fileHeader = fileHeader & wkset & " */" + vbCrLf
    
    wkintext = Workbooks(bname).Worksheets(sname).Cells(R_TblNm, C_TblNm).Value
    wkset = "/*     �e�[�u�����F" & wkintext
        
    '--- Add Start OU 2019/10/15
    strComment = strComment & "COMMENT ON TABLE " & Cells(R_TblId2, C_TblId2).Value & " IS '" & wkintext & "';" & vbCrLf
    
    wkset = spadd_r(wkset, 57 - LenB(StrConv(wkset, vbFromUnicode)))
    fileHeader = fileHeader & wkset & " */" + vbCrLf
    
'--- DEL Start 2019/06/20 SPC
'    wkset = "/*     " & "�쐬��:" & Format(Date, "yyyy/mm/dd")
'
'    wkset = spadd_r(wkset, 57 - LenB(StrConv(wkset, vbFromUnicode)))
'    fileHeader = fileHeader & wkset & " */" + vbCrLf
'--- DEL End 2019/06/20 SPC
    
    fileHeader = fileHeader & "/**********************************************************/" + vbCrLf
'--- DEL Start 2015/02/27 TFC
'    wktext = wktext & "/* �G���[�n���h�����O */" + vbCrLf
'    wktext = wktext & "WHENEVER OSERROR  EXIT OSCODE      ROLLBACK" + vbCrLf
'    wktext = wktext & "WHENEVER SQLERROR EXIT SQL.SQLCODE ROLLBACK" + vbCrLf
'--- DEL End 2015/02/27 TFC
    
    ' �e�[�u��ID�擾
    wkintext = Cells(R_TblId2, C_TblId2).Value
    If Cells(R_Schima, C_Schima).Value <> "" Then
        wkintext = Cells(R_Schima, C_Schima).Value + "." + wkintext
    End If
    strCommentTable = wkintext
    
    
    If ddlKind = DDL_KIND_ALL Or ddlKind = DDL_KIND_TABLE Then
        
        wktext = wktext & "/* CREATE �� */" + vbCrLf
        
        wktext = wktext + "CREATE TABLE " + wkintext + "(" + vbCrLf
        
        curr_l = R_COLNAME '������1�s��
        
        While Cells(curr_l, C_COLNAME) <> ""
            wktext = wktext + "       " & Cells(curr_l, C_COLNAME)
            '--- Add Start 2010/07/30
            Cells(curr_l, C_kata) = UCase(Cells(curr_l, C_kata))
            wktype = Cells(curr_l, C_kata)
            If wktype = "INTEGER" Then
                wktype = "INT"
            End If
            '--- Add End
            wktext = wktext & " " & wktype
            '--- Add START 2010/07/27 OU
            If wktype <> "DATE" And wktype <> "TIMESTAMP" And wktype <> "BLOB" And wktype <> "INT" And wktype <> "BYTEA" Then
                wktext = wktext & "(" & Cells(curr_l, C_keta)
            End If
            '--- Add END
            If (Cells(curr_l, C_kata) = "NUMBER" Or Cells(curr_l, C_kata) = "NUMERIC") And Cells(curr_l, C_shou) <> "" Then
                wktext = wktext & "," & Cells(curr_l, C_shou)
            End If
            '---ADD START 2010/07/27 OU
            If wktype <> "DATE" And wktype <> "TIMESTAMP" And wktype <> "BLOB" And wktype <> "INT" And wktype <> "BYTEA" Then
                wktext = wktext & ")"
            End If
            '---ADD END
            
            '--- Add Start OU 2010/07/26
            '�f�t�H���g�l�̐ݒ�
            If Cells(curr_l, C_def) <> "" Then
                wktext = wktext & " DEFAULT "
                If (wktype = "CHAR") Or (wktype = "VARCHAR2") Or (wktype = "VARCHAR") Then
                    wktext = wktext & "'" & Cells(curr_l, C_def) & "'"
                Else
                    wktext = wktext & Cells(curr_l, C_def)
                End If
            End If
            
            
            strComment = strComment & "COMMENT ON COLUMN " & strCommentTable & "." & Cells(curr_l, C_COLNAME) & " IS '" & Cells(curr_l, C_ITEMNAME) & "';" & vbCrLf
            '---Add End
            If Cells(curr_l, C_uniq) = "��" Then
                wktext = wktext + " UNIQUE"
            End If
            If Cells(curr_l, C_nnul) = "��" Then
                wktext = wktext + " NOT NULL"
            End If
            If Cells(curr_l + 1, C_COLNAME) <> "" Then
                wktext = wktext & ","
            End If
            wktext = wktext + vbCrLf
            curr_l = curr_l + 1
        Wend
        wktext = wktext + "       )"
        
        '--- ADD Start 2019/07/19 SPC
        wkPartitionKind = Cells(R_PartitionKind, C_PartitionKind).Value
        wkPartitionKoumoku = Cells(R_PartitionKoumoku, C_PartitionKoumoku).Value
        ' �p�[�e�B�V������`�����݂���ꍇ
        If wkPartitionKind <> "" And wkPartitionKoumoku <> "" Then
            wktext = wktext & vbCrLf
            wktext = wktext & "       "
            wktext = wktext & "PARTITION BY " & wkPartitionKind & " (" & wkPartitionKoumoku & ")" & vbCrLf
            wktext = wktext & "       "
        End If
        '--- ADD End 2019/07/19 SPC
        
        wkintext = Cells(R_TblSp, C_TblSp).Value
        If wkintext <> "" Then
            wktext = wktext & "TABLESPACE "
            wktext = wktext & wkintext
            wktext = wktext & ";" + vbCrLf
        Else
            wktext = wktext & ";" + vbCrLf
        End If
        
        '--- Add Start 2010/07/30 OU
        wktext = wktext & vbCrLf & strComment
        '--- Add End
    End If
    
    
    If ddlKind = DDL_KIND_ALL Or ddlKind = DDL_KIND_INDEX Then
            
        '=====================================
        ' �v���C�}���L�[�쐬
        '=====================================
        j = 0
        intIndexR = R_COLNAME '������1�s��
        
        ' �J�����Q�ƃ��[�v
        While Cells(intIndexR, C_COLNAME) <> ""
            ' ��L�[��ɐݒ肪����ꍇ
            If Cells(intIndexR, C_primary) <> "" Then
                pkey(j) = CInt(Cells(intIndexR, C_primary)) ' ��L�[�ݒ�l�i���l�ϊ���j�i�[
                pkey_pos(j) = intIndexR                     ' �s�ԍ��i�[
                pkey_seq(j) = j + 1                 ' �V�[�P���X�ԍ��i�[
                j = j + 1
            End If
            
            intIndexR = intIndexR + 1
        Wend
        
        wkpkey = ""
        ' ��L�[�ݒ�f�[�^��ێ������ꍇ
        If j > 0 Then
            ' �V�[�P���X�ԍ��Q�ƃ��[�v
            For i = 0 To j - 1
                ' ��L�[�ݒ�l�Q�ƃ��[�v
                For k = 0 To j - 1
                    If pkey_seq(i) = pkey(k) Then
                        ' ����������ꍇ
                        If i < j - 1 And j > 0 Then
                            wkpkey = wkpkey & Cells(pkey_pos(k), C_COLNAME) & ","
                        Else
                            wkpkey = wkpkey & Cells(pkey_pos(k), C_COLNAME)
                        End If
                    End If
                Next k
            Next i
        End If
        
        ' CREATE���쐬
        If wkpkey <> "" Then
            wktext = wktext & "/* PRIMARY KEY */" + vbCrLf
            '--- Mod Start 2010/09/21
            'wktext = wktext & "ALTER TABLE " & TableId + vbCrLf
            wktext = wktext & "ALTER TABLE " & strCommentTable + vbCrLf
            '--- Mod End
            
            '--- Mod Start 2012/02/27 TFC
            ' ��L�[��`���쐬
            pkeyName = "PK_" & TableId
            pkeyName = Left(pkeyName, 30)
            
            wktext = wktext & " ADD CONSTRAINT " & pkeyName & " PRIMARY KEY(" & wkpkey & ")"
            '--- Mod End 2012/02/27 TFC
        End If
        
        If pkey_pos(0) = 0 Then
            ElseIf Cells(pkey_pos(0), C_IdxSp2).Value <> "" Then
                '--- Add Start 2010/07/30
                Cells(pkey_pos(0), C_IdxSp2).Value = UCase(Cells(pkey_pos(0), C_IdxSp2).Value)
                '--- Add End
                wktext = wktext & " USING INDEX TABLESPACE " & Cells(pkey_pos(0), C_IdxSp2)
                wktext = wktext & ";" + vbCrLf
            ElseIf Cells(R_IdxSp, C_IdxSp).Value <> "" Then
                wktext = wktext & " USING INDEX TABLESPACE " & Cells(R_IdxSp, C_IdxSp).Value
                wktext = wktext & ";" + vbCrLf
            Else
                wktext = wktext & ";" + vbCrLf
        End If
        
        If wkpkey <> "" Then
            wktext = wktext & vbCrLf
        End If
        '--- Add Start 2010/07/29 OU
        '=====================================
        '�C���f�b�N�X�쐬
        '=====================================
        ' �C���f�b�N�X��`��Q�ƃ��[�v
        For intIndexC = C_IndexStart To C_IndexEnd Step 2
        
            strindex = ""
            strIndexSpace = ""
            intIndexR = R_COLNAME '������1�s��
            
            '--- ADD Start 2019/06/21 SPC
            If (Cells(R_COLNAME - 1, intIndexC) <> "") Then
            '--- ADD End 2019/06/21 SPC
                '--- Mod Start 2012/02/27 TFC
                If (Cells(R_COLNAME - 1, intIndexC) = "FNC") Then
                
                    '=======================================
                    ' ���t�@���N�V�����C���f�b�N�X�̏ꍇ
                    ' �C���f�b�N�X��`�̈�ɒ��ڋL�����ꂽ�������
                    ' �����A������B
                    '=======================================
                
                    ' �J�����Q�ƃ��[�v
                    While Cells(intIndexR, C_COLNAME) <> ""
                        
                        '�C���f�b�N�X���ݒ肳��Ă���ꍇ
                        If Cells(intIndexR, intIndexC) <> "" Then
                        
                            If strindex <> "" Then
                                strindex = strindex & ","
                            End If
                            ' ���L�����ꂽ�֐�����������̂܂ܐݒ肷��B
                            strindex = strindex & Cells(intIndexR, intIndexC)
                        End If
                        
                        ' �\�̈於���擾�A����
                        ' �C���f�b�N�X��`�g�ɕ\�̈惊�X�g����`����Ă���ꍇ
                        If strIndexSpace = "" And Cells(intIndexR, C_IdxSp2) <> "" Then
                            ' �ŏ��ɏo�������\�̈於���g�p����
                            strIndexSpace = Cells(intIndexR, C_IdxSp2)
                        End If
                        
                        intIndexR = intIndexR + 1
                    Wend
                
                Else
                    
                    '=======================================
                    ' ���t�@���N�V�����C���f�b�N�X�ȊO�̏ꍇ
                    ' �C���f�b�N�X��`�̈�ɐݒ肳�ꂽ���l���ɁA
                    ' �J����ID�𕶎��A������B
                    '=======================================
                
                    ReDim idx(32768) As Integer
                    ReDim idx_seq(32768) As Integer
                    ReDim idx_pos(32768) As Integer
                    j = 0
                    
                    '-----------------
                    ' �C���f�b�N�X��`�ʒu�̕ێ��ƁA�\�̈於�擾
                    '-----------------
                    ' �J�����Q�ƃ��[�v
                    While Cells(intIndexR, C_COLNAME) <> ""
                        
                        '�C���f�b�N�X���ݒ肳��Ă���ꍇ
                        If Cells(intIndexR, intIndexC) <> "" Then
                            
                            idx(j) = CInt(Cells(intIndexR, intIndexC))  ' �C���f�b�N�X�ݒ�l�i���l�ϊ���j�i�[
                            idx_pos(j) = intIndexR                      ' �s�ԍ��i�[
                            idx_seq(j) = j + 1                          ' �V�[�P���X�ԍ��i�[
                            j = j + 1
                        End If
                        
                        ' �\�̈於���擾�A����
                        ' �C���f�b�N�X��`�g�ɕ\�̈惊�X�g����`����Ă���ꍇ
                        If strIndexSpace = "" And Cells(intIndexR, C_IdxSp2) <> "" Then
                            ' �ŏ��ɏo�������\�̈於���g�p����
                            strIndexSpace = Cells(intIndexR, C_IdxSp2)
                        End If
                        
                        intIndexR = intIndexR + 1
                    Wend
                    
                    '-----------------
                    ' �C���f�b�N�X�ݒ�̍���ID��������쐬
                    '-----------------
                    strindex = ""
                    
                    ' �C���f�b�N�X�ݒ�f�[�^��ێ������ꍇ
                    If j > 0 Then
                        ' �V�[�P���X�ԍ��Q�ƃ��[�v
                        For i = 0 To j - 1
                            ' �C���f�b�N�X�ݒ�l�Q�ƃ��[�v
                            For k = 0 To j - 1
                                If idx_seq(i) = idx(k) Then
                                    ' ����������ꍇ
                                    If i < j - 1 And j > 0 Then
                                        strindex = strindex & Cells(idx_pos(k), C_COLNAME) & ","
                                    Else
                                        strindex = strindex & Cells(idx_pos(k), C_COLNAME)
                                    End If
                                End If
                            Next k
                        Next i
                    End If
                End If
                '--- Mod End 2012/02/27 TFC
            '--- ADD Start 2019/06/21 SPC
            End If
            '--- ADD End 2019/06/21 SPC
            
            ' CREATE���쐬
            If strindex <> "" Then
                If intIndexC = C_IndexStart Then
                    wktext = wktext & "/* INDEX */" + vbCrLf
                End If
                
                '--- Mod Start 2012/02/27 TFC
                wktext = wktext & "CREATE"
                '-----------------
                ' �C���f�b�N�X�̎�ނɉ����č\����ύX
                '-----------------
                ' ���j�[�N�C���f�b�N�X
                If (Cells(R_COLNAME - 1, intIndexC) = "UNQ") Then
                    wktext = wktext & " UNIQUE INDEX"
                    
                ' �r�b�g�}�b�v�C���f�b�N�X
                ElseIf (Cells(R_COLNAME - 1, intIndexC) = "BMP") Then
                    wktext = wktext & " BITMAP INDEX"
                
                ' �m�[�}���C���f�b�N�X�A�t�@���N�V�����C���f�b�N�X
                Else
                    wktext = wktext & "        INDEX"
                    
                End If
                
                '-----------------
                ' �C���f�b�N�X��`���쐬
                '-----------------
                ' ���j�[�N�C���f�b�N�X
                If (Cells(R_COLNAME - 1, intIndexC) = "UNQ") Then
                    indexNamePrefix = "UDX"
                    
                ' �r�b�g�}�b�v�C���f�b�N�X
                ElseIf (Cells(R_COLNAME - 1, intIndexC) = "BMP") Then
                    indexNamePrefix = "BDX"
                
                ' �t�@���N�V�����C���f�b�N�X
                ElseIf (Cells(R_COLNAME - 1, intIndexC) = "FNC") Then
                    indexNamePrefix = "FDX"
                    
                ' �m�[�}���C���f�b�N�X
                Else
                    indexNamePrefix = "IDX"
                
                End If
                
                indexName = indexNamePrefix & CStr((intIndexC - C_IndexStart) / 2 + 1) & "_" & TableId
                ' 30�o�C�g�ȓ��ɒ���
                indexName = Left(indexName, 30)
                wktext = wktext & " " & indexName
                
                wktext = wktext & " ON " & strCommentTable
                wktext = wktext & "(" & strindex & ")"
                
                '-----------------
                ' ���[�J���C���f�b�N�X�̏ꍇ��LOCAL���ǉ�
                '-----------------
                If (Cells(R_COLNAME - 2, C_IndexStart) = "LOCAL INDEX") Then
                    wktext = wktext & " LOCAL"
                End If
                '--- Mod End 2012/02/27 TFC
                
                If strIndexSpace <> "" Then
                    ' �C���f�b�N�X��`�g�̕\�̈��`���g�p����B
                    wktext = wktext & " TABLESPACE " & strIndexSpace
                ElseIf Cells(R_IdxSp, C_IdxSp).Value <> "" Then
                    ' �w�b�_�̃C���f�b�N�X�\�̈��`���g�p����B
                    wktext = wktext & " TABLESPACE " & Cells(R_IdxSp, C_IdxSp).Value
                End If
                wktext = wktext & ";" + vbCrLf
            End If
            
        Next
        '--- Add End
    End If
    
    If ddlKind = DDL_KIND_INDEX And wktext = "" Then
        CreateDdl = ""
    Else
        CreateDdl = fileHeader & wktext
    End If
End Function

'--- Mod Start S.Iwanaga 2010/04/13
'==========================================================
'�y�֐����zLtoP
'�y�T�@�v�z�_�����𕨗����ɕϊ�
'�y���@���zstrSheet     :�����ΏۃV�[�g��
'�@�@�@�@�@lngStartRow  :�����J�n�s
'�@�@�@�@�@lngRepeatCnt :�������J��Ԃ���
'�y�߂�l�z0=����I���@-1=�ُ�I��
'==========================================================
Public Function LtoP(ByVal strSheet As String, _
                    ByVal lngStartRow As Long, _
                    ByVal lngRepeatCnt As Long) As Integer
On Error GoTo Err_Handler
    Dim intDiv      As Integer
    Dim intNoEntry  As Integer
    Dim lngFrom     As Long
    Dim strSrc      As String
    Dim strTmp      As String
    Dim strResult   As String
    Dim strConvFile As String
    Dim valRtn      As Variant
    Dim wsCurrent   As Worksheet
    Dim wbConv      As Workbook
    Dim wsConv      As Worksheet
    
    '�������͉�ʂ̍X�V���~�߂�
    Application.ScreenUpdating = False
    
    lngFrom = lngStartRow
    Set wsCurrent = ActiveWorkbook.Sheets(strSheet)
    
    '�g�p����ϊ��e�[�u���t�@�C���̎w��
    If Len(ConvFilePath) > 0 Then
        strConvFile = ConvFilePath
    Else
        strConvFile = Workbooks(TEMPLATE).Path & "\" & CONVERT_LIST_FILE
    End If
    
    '�ϊ��e�[�u���̃t�@�C�����J��
    If Len(Dir(strConvFile)) = 0 Then
        '�ϊ��e�[�u���t�@�C�������݂��Ȃ�
        MsgBox " �ϊ��p��`�t�@�C����������܂���B" & vbCrLf & strConvFile, vbExclamation + vbOKOnly, "Error"
        GoTo Exit_Handler
    End If
    Set wbConv = Workbooks.Open(strConvFile, 0, True)
    Set wsConv = wbConv.Sheets(CONVERT_LIST_SHEET)

    '�����J�n�s���珈���I���s�܂ł̘_�����𕨗����ɕϊ�
    With wsCurrent
        Do While lngStartRow + lngRepeatCnt >= lngFrom
            
            '�����������͂���Ă��Ȃ����ڂ̂ݏ�������
            If .Cells(lngFrom, C_COLNAME).Value = "" Then
                strSrc = .Cells(lngFrom, C_ITEMNAME).Value
                intDiv = 0
                intNoEntry = 0
                strTmp = ""
                strResult = ""
                valRtn = Null
                While 0 < Len(strSrc)
                    strTmp = strSrc
                    Do While 0 < Len(strTmp)
                        valRtn = Application.VLookup(strTmp, wsConv.Range(DEFINITION_NAME), 2, False)
                        If (Not IsError(valRtn)) Then
                            If (0 < Len(strResult)) Then
                                strResult = strResult + "_"
                            End If
                            strResult = strResult + CStr(valRtn)
                            strSrc = Right(strSrc, Len(strSrc) - Len(strTmp))
                            intDiv = 1
                            Exit Do
                        End If
                        strTmp = Left(strTmp, Len(strTmp) - 1)
                    Loop
                    If (0 = Len(strTmp)) Then
                        If (1 = intDiv) Then
                            strResult = strResult + "_"
                        End If
                        strResult = strResult + Left(strSrc, 1)
                        strSrc = Right(strSrc, Len(strSrc) - 1)
                        intDiv = 0
                        intNoEntry = 1  '�o�^����Ă��Ȃ����ږ������݂���
                    End If
                Wend
                .Cells(lngFrom, C_COLNAME).Value = UCase(strResult)
                
                '�������o�^�`�F�b�N
                If intNoEntry = 0 Then
                    '���͂����S�Ă̘_�������o�^�ς�
                    .Cells(lngFrom, C_COLNAME).Interior.ColorIndex = xlColorIndexNone
                Else
                    '���o�^�̘_�������܂܂��
                    .Cells(lngFrom, C_COLNAME).Interior.ColorIndex = 46
                End If
                
                '������byte���`�F�b�N
                If LenB(StrConv(strResult, vbFromUnicode)) > 30 Then
                    .Cells(lngFrom, C_COLNAME).Interior.ColorIndex = 3
                End If

            End If
            
            lngFrom = lngFrom + 1
        Loop
    End With
    
    LtoP = 0
    GoTo Exit_Handler
    
Err_Handler:
    LtoP = -1
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
    
Exit_Handler:
    If wsConv Is Nothing = False Then Set wsConv = Nothing
    If wbConv Is Nothing = False Then wbConv.Close: Set wbConv = Nothing
    If wsCurrent Is Nothing = False Then Set wsCurrent = Nothing

    '��ʂ̍X�V��~����
    Application.ScreenUpdating = True
    
End Function
'--- Mod End

'--- Add Start S.Iwanaga 2010/04/08
'==========================================================
'�y�֐����zisTblDefSheet
'�y�T�@�v�z�ΏۃV�[�g���e�[�u�����ڃV�[�g���`�F�b�N����
'�y���@���zstrSheet     :�����ΏۃV�[�g��
'�y�߂�l�zTrue=�e�[�u�����ڃV�[�g�@False=�e�[�u�����ڃV�[�g�ȊO
'==========================================================
Public Function isTblDefSheet(ByVal strSheet As String) As Integer

    If ActiveWorkbook.Sheets(strSheet).Cells(R_SheetId, C_SheetId) = 2 Then
        isTblDefSheet = True
    Else
        isTblDefSheet = False
    End If

End Function

'==========================================================
'�y�֐����zchkExAreaState
'�y�T�@�v�z�g���J�����̏�Ԃ��擾
'�y���@���zstrSheet     :�����ΏۃV�[�g��
'�y�߂�l�z0=��\���@1=�\��
'==========================================================
Public Function chkExAreaState(ByVal strSheet As String) As Integer

    Dim ws  As Worksheet
    
    Set ws = ActiveWorkbook.Sheets(strSheet)
    
    If ws.Columns(C_HideSNm & ":" & C_HideENm).EntireColumn.Hidden = True Then
        chkExAreaState = 0
    Else
        chkExAreaState = 1
    End If
    
End Function
'--- Add End

'--- Add Start S.Iwanaga 2010/04/16
'==========================================================
'�y�֐����zchkBlankPName
'�y�T�@�v�z�������ɖ����͂̃Z�����Ȃ����`�F�b�N
'�y���@���zstrSheet     :�����ΏۃV�[�g��
'�y�߂�l�z0=�����̓Z���Ȃ� 1=�����̓Z������
'==========================================================
Public Function chkBlankPName(ByVal strSheet As String) As Integer

    'TODO: �����̓`�F�b�N�������쐬����

    chkBlankPName = 0
    
End Function

'==========================================================
'�y�֐����zsetFFileData
'�y�T�@�v�z�^�ƌ�������Ƀt���b�g�t�@�C���̈ʒu�ƌ������Z�b�g����
'�y���@���zstrSheet     :�����ΏۃV�[�g��
'�y�߂�l�z�Ȃ�
'==========================================================
Public Function setFFileData(ByVal strSheet As String)

    'TODO: �t���b�g�t�@�C�����Z�b�g�������쐬����

    
End Function



'==========================================================
'�y�֐����zcreateCtl
'�y�T�@�v�zSQL*Loader����t�@�C���f�[�^�쐬
'�y���@���zstrSheet     :�����ΏۃV�[�g��
'�y�߂�l�z�쐬��������t�@�C��������
'==========================================================
Public Function createCtl(ByVal strSheet As String) As String

    Dim intI            As Integer
    Dim intIndex        As Integer
    Dim strRtn          As String
    Dim strTableName    As String   '�e�[�u����
    Dim strLoadType     As String   '���[�h�I�v�V����
    Dim strPName        As String   '������
    Dim strPosStart     As String   '�t���b�g�t�@�C���J�n�ʒu
    Dim strPosEnd       As String   '�t���b�g�t�@�C���I���ʒu
    Dim strFFileData()  As String
    '--- Add Start 2010/07/29 OU
    Dim strDataType     As String   '�f�[�^�^�C�v
    Dim strDataTypeCSV As String
    strDataTypeCSV = "CSV"
    Dim strDataHead As String


    '--- Add End
    
    createCtl = ""
    
    '�e�[�u�����ڏ��Ɍr������������
    Call seisho
    
    '�t���b�g�t�@�C�����̃Z�b�g
    Call setFFileData(strSheet)
    
    intIndex = 0
    intI = 0
    With ActiveWorkbook.Sheets(strSheet)
        strTableName = Trim(.Cells(R_TblId, C_TblId).Value)
        strLoadType = Trim(.Cells(R_LdOp, C_LdOp).Value)
        'Add Start 2010/07/29 OU
        strDataType = Trim(.Cells(R_DataTyp, C_DataTyp).Value)
        'Add End
        
        ReDim strFFileData(intIndex)
        strFFileData(intIndex) = getCTLHead & " " & strLoadType & " INTO TABLE " & strTableName & vbCrLf
        'Add Start 2010/07/29 OU
        If strDataType = strDataTypeCSV Then
            strFFileData(intIndex) = strFFileData(intIndex) & " FIELDS TERMINATED BY "",""" & vbCrLf
        End If
        'Add End
        strFFileData(intIndex) = strFFileData(intIndex) & "  (" & vbCrLf
        
        '�������Z���̒l����ɂȂ�܂ŌJ��Ԃ�
        Do While .Cells(R_COLNAME + intI, C_COLNAME).Value <> ""
            
            strPName = Trim(.Cells(R_COLNAME + intI, C_COLNAME).Value)
            intIndex = UBound(strFFileData) + 1
            ReDim Preserve strFFileData(intIndex)
            
            strFFileData(intIndex) = "    " & strPName

            If strDataType <> strDataTypeCSV Then
                strPosStart = Trim(.Cells(R_COLNAME + intI, C_FFilePosition).Value)
                strPosEnd = CStr(CInt(Trim(.Cells(R_COLNAME + intI, C_FFilePosition).Value)) + (CInt(Trim(.Cells(R_COLNAME + intI, C_FFileLength).Value) - 1)))
                strFFileData(intIndex) = strFFileData(intIndex) & "  POSITION(" & strPosStart & ":" & strPosEnd & ")"
            End If
            
            
            intI = intI + 1
            If .Cells(R_COLNAME + intI, 1).Value = "" Then
                strFFileData(intIndex) = strFFileData(intIndex) & vbCrLf
            Else
                strFFileData(intIndex) = strFFileData(intIndex) & "," & vbCrLf
            End If
        Loop
        
        intIndex = UBound(strFFileData) + 1
        ReDim Preserve strFFileData(intIndex)
        
        strFFileData(intIndex) = "  )"

    End With
    
    strRtn = ""
    For intI = 0 To UBound(strFFileData)
        strRtn = strRtn & strFFileData(intI)
    Next intI
    
    createCtl = strRtn

End Function
'--- Add End

'--- Add Start 2010/07/29 OU
Function getCTLHead() As String
    Dim strDirectPath As String
    strDirectPath = Trim(Cells(R_DirectPath, C_DirectPath).Value)

    strDataHead = strDataHead & " -- *****************************************************" & vbCrLf
    strDataHead = strDataHead & " -- SQL*LOADER ����t�@�C��" & vbCrLf
    strDataHead = strDataHead & " -- �e�[�u��ID :" & Cells(R_TblId2, C_TblId2).Value & vbCrLf
    strDataHead = strDataHead & " -- �e�[�u������ :" & Cells(R_TblNm, C_TblNm).Value & vbCrLf
    strDataHead = strDataHead & " -- �쐬�� : " & Format(Date, "yyyy/mm/dd") & " Ver.1.0" & vbCrLf
    strDataHead = strDataHead & " -- �X�V���� : " & vbCrLf
    strDataHead = strDataHead & " -- *****************************************************" & vbCrLf
    If UCase(strDirectPath) = "TRUE" Then
        strDataHead = strDataHead & " OPTIONS (ERRORS=0,DIRECT=TRUE)" & vbCrLf
        strDataHead = strDataHead & " UNRECOVERABLE" & vbCrLf
    Else
        strDataHead = strDataHead & " OPTIONS (ERRORS=0,DIRECT=FALSE)" & vbCrLf
    End If
    
    strDataHead = strDataHead & " LOAD DATA" & vbCrLf
    strDataHead = strDataHead & " INFILE "".DAT""" & vbCrLf
    
    getCTLHead = strDataHead
    
End Function
'--- Add End
