Attribute VB_Name = "����_�W���u��`�쐬"
Option Explicit

Public Const CELL_JIKOU_FILE_NAME = "���s�t�@�C����"                   ' �y���s�t�@�C�����z��`
Public Const CELL_JOB As String = "I5"                                 ' �y�W���u�z��`
Public Const NAME_JOB_ID As String = "�W���uID"                        ' �y�W���uID�z��`
Public Const NAME_DB_CONDITIONAL As String = "DB�폜����"              ' �y�폜�����z��`
Public Const NAME_JIKO_SYORI As String = "���s����"                    ' �y���s�����z��`
Public Const NAME_KINOU_ID As String = "�@�\ID"                        ' �y�@�\ID�z��`
Public Const NAME_KINOU_RENBAN As String = "�@�\���A��"                ' �y�@�\���A�ԁz��`
Public Const NAME_KOUBAN As String = "����"                            ' �y���ԁz��`
Public Const NAME_OUTPUT_PATH As String = "�o�̓p�X(��΃p�X)"         ' �y�o�̓p�X(��΃p�X)�z��`
Public Const NAME_SYORI_SYUBETSU As String = "�������"                ' �y������ʁz��`
Public Const NAME_HULFT_SYUBETSU As String = "HULFT���"               ' �yHULFT��ʁz��`
Public Const NAME_ACMS As String = "ACMS"                              ' �yACMS�z��`
Public Const OUTPUT_ADD_NAME As String = "Config"                      ' �y�o�̓t�@�C���t�����z��`
Public Const OUTPUT_FILE_TYPE As String = ".sh"                        ' �y�o�̓t�@�C����ށz��`
Public Const STR_BIN_BASH As String = "#!/bin/bash"                    ' �ySH�t�@�C���̃w�b�_�z��`
Public Const STR_ACMS_USER As String = "ACMS_USER_ID"                  ' �yACMS���[�U [����旪��]�z��`
Public Const STR_ACMS_FILE As String = "ACMS_FILE_ID"                  ' �yACMS�t�@�C�� [�t�@�C������]�z��`
Public Const WS_CELL_SETTING_BUILD_FLG As String = "�쐬�t���O"      ' �V�[�g�y���ڐݒ�z�́y�쐬�t���O�z�Z����`
Public Const WS_CELL_SETTING_CONSTANT As String = "�Œ薼"             ' �V�[�g�y���ڐݒ�z�́y�Œ薼�z�Z����`
Public Const WS_CELL_SETTING_PRC_PATH As String = "���s�p�X"           ' �V�[�g�y���ڐݒ�z�́y���s�p�X�z�Z����`
Public Const WS_CELL_SETTING_SETSUZOKU_SAKI As String = "SFTP�ڑ���"   ' �V�[�g�y���ڐݒ�z�́ySFTP�ڑ���z�Z����`
Public Const WS_CELL_SETTING_SYORI_SYUBETSU As String = "�����敪"     ' �V�[�g�y���ڐݒ�z�́y�����敪�z�Z����`
Public Const WS_NAME_SHEET_KOMOKU_SETTING As String = "���ڐݒ�"       ' �V�[�g�y���ڐݒ�z�� ��`


' ���C������
' �߂�l   : �Ȃ�
Public Function CreateShFile_Seibu()

    Dim DictShFile As Object: Set DictShFile = createDictionary                       ' �y�V�F�����e�z���X�g���`(Dictionary)
    Dim backFlg As Boolean                                                            ' ��ʍX�V�@�\�t���O���ꎞ��~
    Dim cnt As Integer: cnt = 0                                                       ' ���[�v��
    Dim curKey As Variant                                                             ' Dic�̃��[�v���L�[���擾
    Dim folderPath As String: folderPath = getFolderPath(NAME_OUTPUT_PATH)            ' �t�H���_�p�X�̒�`
    Dim key As String                                                                 ' Dic�̃��[�v���L�[���擾�iString�j

    Application.ScreenUpdating = False                                                ' ��ʍX�V�@�\�t���O���ꎞ��~
    backFlg = Application.ScreenUpdating                                              ' ��ʍX�V�@�\�t���O���o�b�N�A�b�v
    Set DictShFile = getShText                                                        ' �o�̓��X�g���擾
    Application.ScreenUpdating = backFlg                                              ' ��ʍX�V�@�\�t���O�����ɖ߂�

    For Each curKey In DictShFile                                                     ' �ݒ肵��Dic���e�ɂāA���[�v����
        key = curKey                                                                  ' Dic��key��Variant�^�C�v�ł��BString�ɕϊ�
        key = getFileName(key, OUTPUT_FILE_TYPE, OUTPUT_ADD_NAME)
        cnt = cnt + CreateFileWithoutBom(folderPath, key, DictShFile.item(curKey))
    Next

    ' �o�͊������b�Z�[�W
    MsgBox "�V�F���t�@�C���̏o�͂��������܂����B" & vbCrLf & "�o�͐� : " & DictShFile.count

     ' �쐬������̂����݂���ꍇ�A�t�H���_���J��
    If cnt > 0 Then
        Dim filepath As Range
        Set filepath = Range(searchCell(NAME_OUTPUT_PATH))
        Shell "C:\Windows\Explorer.exe " & Cells(filepath.Row + 1, filepath.Column), vbNormalFocus
    End If

End Function

' �p�����^ : �Ȃ�
' �߂�l   : �o�͓��e
Private Function getShText() As Object

    Dim DETAIL_JIKKO_SYORI As String:                                                                                                                      ' ���s�����Z���̓��e���擾
    Dim DictShFile As Object: Set DictShFile = createDictionary                                                                                            ' �y�V�F�����e�z���X�g���`(Dictionary)
    Dim ListHulftSyubetsu As Object: Set ListHulftSyubetsu = getGroupList(NAME_HULFT_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING, True)                         ' Hulft��ʃ��X�g���擾
    Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getGroupList(WS_CELL_SETTING_SYORI_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING, False)             ' ������ʃ��X�g���擾
    Dim cellJikkoSyori As Range: Set cellJikkoSyori = Range(searchCell(NAME_JIKO_SYORI))                                                                   ' ���s�����̃Z�����擾
    Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(NAME_SYORI_SYUBETSU))                                                         ' �y������ʁz�Z��
    Dim getTotalCase As Integer: getTotalCase = getCountCase(NAME_KOUBAN, 2)                                                                               ' �P�[�X���̒�`
    Dim jobList As Object: Set jobList = getVertivalListbyCnt(searchCell(NAME_JOB_ID), getTotalCase, 2)                                                                   ' �W���u���X�g
    Dim curCell As Range                                                                                                                                   ' ���[�v���Z�����擾�iString�j
    Dim curKey As Variant                                                                                                                                       ' Dic�̃��[�v���L�[���擾�iString�j
    Dim shText As String                                                                                                                                   ' �y�V�F�����e�z���`

    ' �W���u���X�g�Ń��[�v����
    For Each curKey In jobList.keys
        Set curCell = Range(curKey)
        DETAIL_JIKKO_SYORI = Cells(curCell.Row, cellJikkoSyori.Column)
        ' �w�b�_�̐ݒ�
        shText = STR_BIN_BASH + vbLf
        shText = shText + vbLf + setKomokuComment(removeSpecCode(DETAIL_JIKKO_SYORI))
        ' ���e�̐ݒ�
        shText = setShText(curCell.Row, shText, ListSyoriSyubetsu, ListHulftSyubetsu)

        ' Dic�ɑ��
        If Len(shText) > 0 Then
            If DictShFile.Exists(jobList.item(curKey)) Then
                DictShFile.item(jobList.item(curKey)) = shText                                                    ' ���łɑ��݂���ꍇ�A���e���X�V����
            Else
                DictShFile.Add jobList.item(curKey), shText                                                       ' ���݂��Ȃ��ꍇ�A�ǉ�����
            End If

            Dim cellJikkoFileName As Range: Set cellJikkoFileName = Range(searchCell(CELL_JIKOU_FILE_NAME))       ' ���s�t�@�C�����̂̃Z�����擾
            Dim jikoKbn As String: jikoKbn = Cells(curCell.Row, cellSyoriSyubetsu.Column)                         ' �Ώ�Row�̎��s��ʂ��擾
            Dim prcPath As String: prcPath = ListSyoriSyubetsu.item(jikoKbn)(2)                                   ' �Ώێ��s��ʂ̒�`�p�X���擾
            Cells(curCell.Row, cellJikkoFileName.Column) = prcPath                                                ' ���s�t�@�C�����̂�ݒ�
            Cells(curCell.Row, cellJikkoFileName.Column + 1) = curCell.value                                      ' ���s�p�����[�^���̂�ݒ�
        End If
    Next

    Set getShText = DictShFile

End Function

' �p�����^ : ���݃��[�A���ݏo�͓��e�A������ʃ��X�g�AHULFT��ʃ��X�g
' �߂�l   : �o�͓��e
Private Function setShText(ByVal curRow As Integer, shText As String, ListSyoriSyubetsu As Object, ListHulftSyubetsu As Object) As String

    Dim cellNextCtgl As Range                                                                                                        ' ���W���u�J�^���O
    Dim count As Integer                                                                                                             ' ���ڃ��[�v��
    Dim countCtgy As Integer                                                                                                         ' �J�^���O���[�v��
    Dim startCol As Integer                                                                                                          ' �J�n�J����
    Dim getJobCtgyContent() As String: getJobCtgyContent() = getTitleList(NAME_SYORI_SYUBETSU, getCountKomoku(NAME_SYORI_SYUBETSU))  ' �W���u�J�^���O�̒�`
    Dim sizeJobCatagoryList As Integer: sizeJobCatagoryList = getArrayLength(getJobCtgyContent())                                    ' �W���u�J�^���O���X�g�̓��e�𒷂�

    ' �W���u�J�^���O���X�g�̓��e�̒����Ń��[�v����
    For countCtgy = 0 To sizeJobCatagoryList - 1
        ' ���ڃ��[�v�񐔂�������
        count = 0
        ' ���W���u�J�^���O�̎擾
        Set cellNextCtgl = Nothing
        ' ���̃J�^���O�����݂���ꍇ�A�擾����
        If countCtgy < sizeJobCatagoryList - 1 Then
            Set cellNextCtgl = Range(searchCell(getJobCtgyContent(countCtgy + 1)))
        End If

        ' ���W���u�J�^���O���擾
        Dim curJobCatagory As Range: Set curJobCatagory = Range(searchCell(getJobCtgyContent(countCtgy)))

        ' �J�n�Z���̐ݒ�
        Dim startCell As Range: Set startCell = Cells(curRow, curJobCatagory.Column)

        ' �y������ʁz�̃J�����́A�ΏۊO�̏�����ʂ����͂��ꂽ�ꍇ�A�쐬���Ȃ�
        If curJobCatagory.value = NAME_SYORI_SYUBETSU Then
            If ListSyoriSyubetsu.Exists(startCell.value) = False Then
                shText = ""
                Exit For
            ElseIf ListSyoriSyubetsu.item(startCell.value)(1) <> "�Z" Then
                shText = ""
                Exit For
            End If
        End If

        ' �J�n�J����
        startCol = startCell.Column
        shText = shText + vbLf
        ' ���W���u�J�^���O�����݂���ꍇ
        If Not (cellNextCtgl Is Nothing) Then
            ' ���W���u�J�^���O�̃J�����Ɠ����܂ŁA���[�v����
            Do While startCol + count <> cellNextCtgl.Column
                shText = shText + setShText02(startCol, count, curJobCatagory, startCell, ListSyoriSyubetsu, ListHulftSyubetsu)
                count = count + 1
            Loop
        Else
            ' ���̃J���������݂��Ȃ��܂ŁA���[�v����
            Do While Len(Cells(curJobCatagory.Row + 1, startCol + count).value) > 0
                shText = shText + setShText02(startCol, count, curJobCatagory, startCell, ListSyoriSyubetsu, ListHulftSyubetsu)
                count = count + 1
            Loop
        End If

        ' ACMS���e���蓮�ō쐬
        shText = addACMSExtendDetail(shText, curRow, curJobCatagory.value)

    Next

    If Len(shText) > 0 Then
        ' �萔���e�̒ǉ�
        shText = addConstantText(shText, curRow)
    End If

    setShText = shText

End Function


' �p�����^ : ���ݏo�͓��e�A���݃��[
' �߂�l   : �o�͓��e�i�Œ�l�j
Private Function addConstantText(ByVal shText As String, curRow As Integer) As String

    ' �A�ԃZ��
    Dim rnBCell As Range: Set rnBCell = Range(searchCell(NAME_KINOU_RENBAN)): Set rnBCell = Cells(curRow, rnBCell.Column)
    ' �@�\�Z��
    Dim kinouId As Range: Set kinouId = Range(searchCell(NAME_KINOU_ID)): Set kinouId = Cells(curRow, kinouId.Column)
    ' �l
    Dim value As String: value = ""

    If Len(kinouId.Text) > 0 Then
        value = kinouId.Text
    End If

    ' �萔�̓��e(����)����
    shText = shText + vbLf + setKomokuComment("�萔���e")
    If Len(value) > 0 Then
        shText = shText + setDetailByOneSet("�v���Z�XID", "PROC_ID", value + padLeftString(rnBCell.value, "0", 3))
        shText = shText + setDetailByOneSet("�W���uID", "JOB_ID", value + padLeftString(rnBCell.value, "0", 4))
    Else
        shText = shText + setDetailByOneSet("�v���Z�XID", "PROC_ID", "")
        shText = shText + setDetailByOneSet("�W���uID", "JOB_ID", "")
    End If

    ' �Œ薼���X�g���擾
    Dim listConstant As Object: Set listConstant = getGroupList(WS_CELL_SETTING_CONSTANT, WS_NAME_SHEET_KOMOKU_SETTING, True)
    Dim Constkeys As Variant
    ' Dic�̃��[�v���L�[���擾�iString�j
    Dim Constkey As String
    Dim value2 As String
    For Each Constkeys In listConstant
        Constkey = Constkeys
        value = listConstant.item(Constkey)(0)
        value2 = listConstant.item(Constkey)(1)
        shText = shText + setDetailByOneSet(Constkeys, value, value2)
    Next

    addConstantText = shText

End Function

' �p�����^ : ���ݏo�͓��e�A�Z�����e
' �߂�l   : �o�͓��e�iSFTP�̒ǉ����e�j
Private Function setSFTPExtraDetail(ByVal shText As String, cellValue As String) As String

    Dim value As String: value = ""

    Dim getContent As Boolean: getContent = False

    'SFTP�ǉ����e
    Dim sftpObject As Object
    If Len(cellValue) > 0 Then
        'Set sftpObject = getGroupListbySelectedValue(WS_CELL_SETTING_SETSUZOKU_SAKI, WS_NAME_SHEET_KOMOKU_SETTING, True, cellValue)
        Set sftpObject = getGroupListbySelectedValue(WS_CELL_SETTING_SETSUZOKU_SAKI, WS_NAME_SHEET_KOMOKU_SETTING, True, False, 0, cellValue)
        Dim key As Variant
        ' �ꌏ�����Ȃ��\��
        For Each key In sftpObject

            value = sftpObject.item(key)(1)
            shText = shText + setDetailByOneSet("SFTP�z�X�g", "SFTP_HOST", value)
            value = sftpObject.item(key)(2)
            shText = shText + setDetailByOneSet("SFTP���[�U�[", "SFTP_USER", value)
            value = sftpObject.item(key)(3)
            shText = shText + setDetailByOneSet("SFTP�閧���p�X", "SFTP_KEY_PATH", value)
            ' �ݒ�σt���O
            getContent = True
        Next
    End If

    If getContent = False Then
        shText = shText + setDetailByOneSet("SFTP�z�X�g", "SFTP_HOST", "")
        shText = shText + setDetailByOneSet("SFTP���[�U�[", "SFTP_USER", "")
        shText = shText + setDetailByOneSet("SFTP�閧���p�X", "SFTP_KEY_PATH", "")
    End If
    setSFTPExtraDetail = shText

End Function

' �p�����^ : �R�����g���e
' �߂�l   : ���ڃR�����g�̍쐬
Private Function setKomokuComment(ByVal comment As String) As String

    setKomokuComment = padRightString("# *----" + comment, "-", 60) + vbLf

End Function

' �p�����^ : �^�C�g�����e
' �߂�l   : ���ڃ^�C�g���̍쐬
Private Function setKomokuTitle(ByVal title As String) As String

    setKomokuTitle = "# " + title + vbLf

End Function

' �p�����^ : �R���e���c�^�C�g���A���e
' �߂�l   : �R���e���c���e�̍쐬
Private Function setKomokuDetail(ByVal title As String, value As String) As String

    setKomokuDetail = title + "=" + """" + value + """" + vbLf

End Function

' �p�����^ : ���ځi�����j�A���ځi�p���j�A���e
' �߂�l   : �R���e���c���e�i�Z�b�g�j�̍쐬
Private Function setDetailByOneSet(ByVal komokuKanji As String, komoku As String, value As String) As String

    setDetailByOneSet = ""
    setDetailByOneSet = setDetailByOneSet + setKomokuTitle(komokuKanji)
    setDetailByOneSet = setDetailByOneSet + setKomokuDetail(komoku, value)

End Function

' �p�����^ : �J�n���[�A���݃��[�v�񐔁A���݃J�e�S���A�J�n�Z���A������ʃ��X�g�AHULFT��ʃ��X�g
' �߂�l   : �o�͓��e
Private Function setShText02(ByVal startCol As Integer, _
                            count As Integer, _
                            curJobCatagory As Range, _
                            startCell As Range, _
                            ListSyoriSyubetsu As Object, _
                            ListHulftSyubetsu As Object _
                            ) As String
    setShText02 = ""
    Dim value As String
    ' ���݃J���� =���W���u�J�^���O�̃J�����̏ꍇ
    If startCol + count = curJobCatagory.Column Then
        ' ���W���u�J�^���O�̓��e(����)����
        setShText02 = setShText02 + setKomokuComment(Replace(curJobCatagory.value, vbCrLf, ""))
    End If

    ' �����ڂ̐����i�����j����
    setShText02 = setShText02 + setKomokuTitle(Replace(Cells(curJobCatagory.Row + 2, startCol + count).value, vbLf, ""))

    ' �����ڂ���
    Dim curTitle As String: curTitle = Cells(curJobCatagory.Row + 1, startCol + count)
    ' �����ڂ̓��e����
    value = Replace(Cells(startCell.Row, startCol + count).value, vbCrLf, "")
    ' �y������ʁz�̃J�����̏ꍇ
    If curJobCatagory.value = NAME_SYORI_SYUBETSU Then
        value = ListSyoriSyubetsu.item(value)(0)
    ElseIf curTitle = "HULFT_TYPE" And Len(value) > 0 Then
        value = ListHulftSyubetsu.item(value)(0)
    ElseIf curTitle = STR_ACMS_USER And Len(value) > 0 Then
        ' Acms���[�UID���X�g���擾
        value = getValueByKeyFromDictionary(STR_ACMS_USER, value, WS_NAME_SHEET_KOMOKU_SETTING)
    ElseIf curTitle = STR_ACMS_FILE And Len(value) > 0 Then
        ' Acms�t�@�C��ID���X�g���擾
        value = getValueByKeyFromDictionary(STR_ACMS_FILE, value, WS_NAME_SHEET_KOMOKU_SETTING)
    End If

    setShText02 = setShText02 + setKomokuDetail(Replace(Cells(curJobCatagory.Row + 1, startCol + count).value, vbLf, ""), value)

    If curTitle = "SFTP_DEST" Then
        'SFTP�����敪�̒ǉ�
        'setShText02 = addSFTPSyoriKbn(setShText02, startCell)
        'SFTP�ǉ����̒ǉ�
        setShText02 = setSFTPExtraDetail(setShText02, value)
    End If

End Function


' �p�����^ : ���ݓ��e�A�J�n�Z��
' �߂�l   : �o�͓��e
' ���g�p 20191007
Private Function addSFTPSyoriKbn(ByVal setShText02 As String, startCell As Range)

        ' �����ڂ̐����i�����j����
        setShText02 = setShText02 + setKomokuTitle("SFTP�����敪")
        Dim SFTP_KBNObject As Object
        Set SFTP_KBNObject = getGroupList("SFTP�����敪", WS_NAME_SHEET_KOMOKU_SETTING, True)

        Dim colKbn As Range
        Set colKbn = Range(searchCell(NAME_SYORI_SYUBETSU))

        Dim selectedKbn As Range
        Set selectedKbn = Range(Cells(startCell.Row, colKbn.Column).Address)

        Dim valueSFTP As String
        If Len(selectedKbn.value) > 0 And SFTP_KBNObject.Exists(selectedKbn.value) = True Then
            valueSFTP = SFTP_KBNObject.item(selectedKbn.value)(0)
        Else
            valueSFTP = ""
        End If
        addSFTPSyoriKbn = setShText02 + setKomokuDetail("SFTP_KBN", valueSFTP)

End Function

' �p�����^ : ���ݓ��e�A���݃��[�A���݃W���u�J�e�S��
' �߂�l   : �o�͓��e�iACMS ���e�j
Private Function addACMSExtendDetail(ByVal shText As String, curRow As Integer, jobTitle As String) As String

    If jobTitle = NAME_ACMS Then

        Dim acmsUserRange As Range: Set acmsUserRange = Range(searchCell(STR_ACMS_USER))
        Dim acmsFileRange As Range: Set acmsFileRange = Range(searchCell(STR_ACMS_FILE))

        Dim strCombine As String
        If Cells(curRow, acmsUserRange.Column) <> "" Then
            strCombine = ""
            strCombine = strCombine + getValueByKeyFromDictionary(STR_ACMS_USER, Cells(curRow, acmsUserRange.Column).value, WS_NAME_SHEET_KOMOKU_SETTING) ' Acms���[�UID���X�g���擾
            strCombine = strCombine + "_"
            strCombine = strCombine + getValueByKeyFromDictionary(STR_ACMS_FILE, Cells(curRow, acmsFileRange.Column).value, WS_NAME_SHEET_KOMOKU_SETTING) ' Acms���[�UID���X�g���擾
            shText = shText + setDetailByOneSet("ACMS�A�v���P�[�V���� [����旪��]_[�t�@�C������]", "ACMS_APL_ID", strCombine)
        Else
            shText = shText + setDetailByOneSet("ACMS�A�v���P�[�V���� [����旪��]_[�t�@�C������]", "ACMS_APL_ID", "")
        End If

    End If

    addACMSExtendDetail = shText
End Function





