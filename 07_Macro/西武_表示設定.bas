Attribute VB_Name = "����_�\���ݒ�"
' �����@   : ������ʂ̃��X�g�֐����쐬
Public Function setListForSyoriKbn()
    Const WS_CELL_SETTING_SYORI_SYUBETSU As String = "�����敪"     ' �V�[�g�y���ڐݒ�z�́y�����敪�z�Z����`
    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "���ڐݒ�"       ' �V�[�g�y���ڐݒ�z�� ��`
    Const CELL_SYORI_SYUBETSU As String = "�������"                ' �y������ʁz��`
    Const KOUBAN As String = "����"                                 ' �y���ԁz��`

    If isCellExist(CELL_SYORI_SYUBETSU) = True Then

        ' ������ʃ��X�g���擾
        Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getListContent(WS_CELL_SETTING_SYORI_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING)
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(CELL_SYORI_SYUBETSU))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellSyoriSyubetsu.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellSyoriSyubetsu.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim item As Variant

        For Each item In ListSyoriSyubetsu.Items
            If Len(str) = 0 Then
                str = item
            Else
                str = str + "," + item
            End If
        Next

        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
    End If
End Function

' �����@   : SFTP�̃��X�g�֐����쐬
Public Function setListForSFTPKbn()
    Const WS_CELL_SETTING_SFTP_SYUBETSU As String = "SFTP�����敪"      ' �V�[�g�y���ڐݒ�z�́ySFTP�����敪�z�Z����`
    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "���ڐݒ�"           ' �V�[�g�y���ڐݒ�z�� ��`
    Const CELL_SFTP_SYUBETSU As String = "SFTP�����敪"                 ' �ySFTP�����敪�ʁz��`
    Const KOUBAN As String = "����"                                     ' �y���ԁz��`

    If isCellExist(CELL_SFTP_SYUBETSU) = True Then
        ' ������ʃ��X�g���擾
        Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getListDictionary(WS_CELL_SETTING_SFTP_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING)
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(CELL_SFTP_SYUBETSU))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellSyoriSyubetsu.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellSyoriSyubetsu.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim key As Variant

        For Each key In ListSyoriSyubetsu.keys
            If Len(str) = 0 Then
                str = key
            Else
                str = str + "," + key
            End If
        Next

        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
    End If
End Function

' �����@   : SFTP�ڑ���̃h���b�v���X�g���쐬
Public Function isSFTPDestCell(ByVal Target As Range) As Integer
    ' �߂�l
    isSFTPDestCell = 9

    Const CELL_SFTP_SETSUZOKU_SAKI As String = "SFTP�ڑ���"                 ' �ySFTP�ڑ���z��`
    Const CELL_SFTP_SYUBETSU As String = "�����敪"                         ' �ySFTP�����敪�ʁz��`
    Const WS_CELL_SETTING_SFTP_SYUBETSU As String = "SFTP�����敪"          ' �V�[�g�y���ڐݒ�z�́ySFTP�����敪�z�Z����`
    Const WS_CELL_SFTP_KEY As String = "SFTP�L�["                           ' �V�[�g�y���ڐݒ�z�́ySFTP�L�[�z�Z����`
    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "���ڐݒ�"               ' �V�[�g�y���ڐݒ�z�� ��`


    If isCellExist(CELL_SFTP_SETSUZOKU_SAKI) = True Then
        ' �ڑ���̃Z�����擾
        Dim cellSetzuZokuSaki As Range: Set cellSetzuZokuSaki = Range(searchCell(CELL_SFTP_SETSUZOKU_SAKI))
    
        ' �I���J���� = �ڑ���̃Z���̃J�����̏ꍇ
        If Target.Column = cellSetzuZokuSaki.Column Then
            ' SFTP�敪�Z���̎擾
            Dim cellSFTPSyubetsu As Range: Set cellSFTPSyubetsu = Range(searchCell(CELL_SFTP_SYUBETSU))
            ' SFTP�敪�Z����Dictionary���擾
            Dim getKey As Object: Set getKey = getListDictionary(WS_CELL_SETTING_SFTP_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING)
            ' �L�[�̒�`
            Dim key As String: key = Cells(Target.Row, cellSFTPSyubetsu.Column)
            ' �h���b�v���X�g�̎擾
            Dim getDropList As Object: Set getDropList = getGroupListbySelectedValue(WS_CELL_SFTP_KEY, WS_NAME_SHEET_KOMOKU_SETTING, True, False)
            ' �h���b�v���X�g�̎擾�\�̏ꍇ
    
            Dim backFlg As Boolean: backFlg = Application.ScreenUpdating
            Application.ScreenUpdating = False
            If getDropList.count > 0 Then
    
                Dim str As String
                Dim keys As Variant
    
                For Each keys In getDropList.keys
                    key = keys
                    If Len(str) = 0 Then
                        str = getDropList.item(key)(1)
                    Else
                        str = str + "," + getDropList.item(key)(1)
                    End If
                Next
                Target.Validation.Delete
                Target.Validation.Add Type:=xlValidateList, Formula1:=str
            Else
                Target.Validation.Delete
    
            End If
    
            Application.ScreenUpdating = backFlg
        Else
    
        End If
    End If

    isSFTPDestCell = 0
End Function

' �����@   : ������ʂ̃��X�g�֐����쐬
Public Function setListForEmptyFileFlg()

    Const CELL_KARA_FILE_SAKUSEI As String = "��t�@�C���쐬"                ' �y��t�@�C���쐬�z��`
    Const KOUBAN As String = "����"                                          ' �y���ԁz��`

    If isCellExist(CELL_KARA_FILE_SAKUSEI) = True Then
        
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellKaraFileSakusei As Range: Set cellKaraFileSakusei = Range(searchCell(CELL_KARA_FILE_SAKUSEI))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellKaraFileSakusei.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellKaraFileSakusei.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim item As Variant
        str = "YES,NO"
        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
    End If
End Function

' �����@   : HULFT��ʃ��X�g�֐����쐬
Public Function setListForHulftType()

    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "���ڐݒ�"           ' �V�[�g�y���ڐݒ�z�� ��`
    Const CELL_HULFT_TYPE As String = "HULFT���"                ' �yHULFT��ʁz��`
    Const KOUBAN As String = "����"                                     ' �y���ԁz��`

    If isCellExist(CELL_HULFT_TYPE) = True Then
        
        ' ������ʃ��X�g���擾
        Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getListDictionary(CELL_HULFT_TYPE, WS_NAME_SHEET_KOMOKU_SETTING)
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(CELL_HULFT_TYPE))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellSyoriSyubetsu.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellSyoriSyubetsu.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim key As Variant

        For Each key In ListSyoriSyubetsu.keys
            If Len(str) = 0 Then
                str = key
            Else
                str = str + "," + key
            End If
        Next

        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
    End If
End Function

' �����@   : HULFT��ʃ��X�g�֐����쐬
Public Function setListForAcmsType(ByVal strInput As String)

    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "���ڐݒ�"           ' �V�[�g�y���ڐݒ�z�� ��`
    Const KOUBAN As String = "����"                                     ' �y���ԁz��`

    If isCellExist(strInput) = True Then
        
        ' ������ʃ��X�g���擾
        Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getListDictionary(strInput, WS_NAME_SHEET_KOMOKU_SETTING)
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(strInput))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellSyoriSyubetsu.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellSyoriSyubetsu.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim key As Variant

        For Each key In ListSyoriSyubetsu.keys
            If Len(str) = 0 Then
                str = key
            Else
                str = str + "," + key
            End If
        Next

        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
        
    End If
End Function




