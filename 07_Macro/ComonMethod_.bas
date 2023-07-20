Attribute VB_Name = "ComonMethod"

Public Sub CallMacro()
    ' ��`�쐬����
    CreateShFile_Seibu

End Sub

' ���s�E�X�y�[�X�R�[�h�̍폜
' �p�����^ : �C���O������
' �߂�l   : �C���㕶����
Public Function removeSpecCode(ByVal str As String) As String

    Dim retStr As String
    retStr = str
    retStr = Replace(retStr, " ", "")
    retStr = Replace(retStr, "�@", "")
    retStr = Replace(retStr, vbCrLf, "")
    retStr = Replace(retStr, vbCr, "")
    retStr = Replace(retStr, vbLf, "")

    removeSpecCode = retStr

End Function

' �L�[��Dictionary�̃C���f�b�N�X���擾
' ���g�p20190910
' �p�����^ : �ΏۃL�[�A�Ώ�Dict�A(Optional)���Zindex
' �߂�l   : �C���f�b�N�X
Public Function getIndexFromDicByKey(ByVal key As Variant, dic As Object, Optional plusAfterIndex As Integer = 0) As Integer

    Dim keys As Variant
    Dim index As Integer
    index = 0
    For Each keys In dic.keys
        If keys = key Then
            Exit For
        End If
        index = index + 1
    Next
    getIndexFromDicByKey = index + plusAfterIndex
End Function

' ���[�N�V�[�g�̎擾
' �p�����^ : �V�[�g����
' �߂�l   : ���[�N�V�[�g
Public Function getWorSheet(ByVal selectSheet As String) As Worksheet
        ' ���[�N�V�[�g�̒�`
    Dim s_workSheet As Worksheet
    ' ���[�N�V�[�g�����͂���Ȃ��ꍇ�A���݂̃��[�N�V�[�g�Ƃ���
    If selectSheet <> "" Then
        Set s_workSheet = Worksheets(selectSheet)
    Else
        Set s_workSheet = ActiveSheet
    End If
    Set getWorSheet = s_workSheet
End Function

' ����̐����擾(�w���̎w��Z������AX���̋󔒃Z���܂Ōv�Z)
' �p�����^ : �J�n�Z���A�h���X�A(Optional)�ΏۃV�[�g��
' �߂�l   : �J�E���g��
Public Function getCountHorizon(ByVal cellAddressStart As String, Optional nameWorkSheet As String = "") As Integer

    ' ���[�N�V�[�g�̒�`
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(nameWorkSheet)
    ' �J�n�Z���̎擾
    Dim startCell As Range: Set startCell = s_workSheet.Range(cellAddressStart)
    ' ���[�v�J�E���g
    Dim cntLoop As Integer: cntLoop = 0
    Do While s_workSheet.Cells(startCell.Row, startCell.Column + cntLoop) <> ""
        cntLoop = cntLoop + 1
    Loop

    getCountHorizon = cntLoop

End Function

' Y�����X�g�̎擾
' �J�n�Z�����܂߂��Ȃ�
' ���g�p20190910
' �p�����^ : �J�n�Z���A�h���X�A(Optional)�ΏۃV�[�g��
' �߂�l   : Dictionary(�L�[ : �Z���A�h���X�AItem : �Z�����e)
Public Function getListContent(ByVal cellContent As String, Optional nameWorkSheet As String = "", Optional offset As Integer = 0) As Object
    ' Dictionary�̒�`
    Dim retOut   As Object: Set retOut = createDictionary
    ' ���[�N�V�[�g�̎擾
    Dim sworkSheet As Worksheet: Set sworkSheet = getWorSheet(nameWorkSheet)
    ' �J�n�Z���̎擾
    Dim cellStart As Range: Set cellStart = sworkSheet.Range(searchCell(cellContent, nameWorkSheet))
    ' ' �J�n�Z���܂߂Ȃ����߁A�I�t�Z�b�g�y1�z����J�n����
    Dim cntLoop As Integer: cntLoop = offset
    ' ���[�v�Z���̎擾
    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).value <> ""
        ' ���[�v�Z���̐ݒ�
        Set Cellcur = sworkSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        ' ��L�[�d���`�F�b�N�̐ݒ�
        If retOut.Exists(Cellcur.value) = False Then
            retOut.Add Cellcur.Address, Cellcur.value
        End If
        ' ���[�v�J�E���g�𑫂�
        cntLoop = cntLoop + 1
    Loop
    ' �߂�l
    Set getListContent = retOut

End Function

' �z��̒������Z�o
' �p�����^ : �Ώ۔z��
' �߂�l   : �J�E���g��
Public Function getArrayLength(ByRef arry() As String) As Integer
    ' �Ō�̃C���f�b�N - �ŏ��̃C���f�b�N + 1
    getArrayLength = UBound(arry()) - LBound(arry()) + 1

End Function

' �Z���̌���
' �p�����^ : �������e�A�Ώۃ��[�N�V�[�g��
' �߂�l   : �Z���A�h���X
Public Function searchCell(ByVal cellContent As String, Optional worksheetName As String = "") As String
    ' ���[�N�V�[�g�̒�`
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(worksheetName)
    ' �����Ώۂ��擾
    Dim retRange As Range: Set retRange = s_workSheet.Cells.Find(cellContent, LookAt:=xlWhole)
    ' �����Ώۂ����݂��Ȃ��ꍇ�A�����I��
    If (retRange Is Nothing) Then
        MsgBox cellContent & "��������܂���B" & vbCrLf & "�V�F���t�@�C�����쐬�ł��܂���B" _
                , vbYes + vbExclamation, "�ُ�"
        End
    End If
    ' �����Ώۂ̃Z����Ԃ�
    searchCell = retRange.Address

End Function

' �Z���̑��݊m�F
' �p�����^ : �������e�A�Ώۃ��[�N�V�[�g��
' �߂�l   : �Z���A�h���X
Public Function isCellExist(ByVal cellContent As String, Optional worksheetName As String = "") As Boolean
    isCellExist = False
    ' ���[�N�V�[�g�̒�`
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(worksheetName)
    ' �����Ώۂ��擾
    Dim retRange As Range: Set retRange = s_workSheet.Cells.Find(cellContent, LookAt:=xlWhole)
    ' �����Ώۂ����݂��Ȃ��ꍇ�A�����I��
    If (retRange Is Nothing = False) Then
        isCellExist = True
    End If
End Function

' �t�@�C���쐬
' Linux�Ή��̂��߁ABOM�Ȃ��̃t�H�}�b�g
' �p�����^ : �t�H���_�p�X�A�t�@�C����Ζ��A�t�@�C�����e
' �߂�l   : �����l(1)�̂�
Public Function CreateFileWithoutBom(ByVal folderPath As String, file As String, fileContent As String) As Integer

    ' �t�@�C���p�X�̐錾
    Dim strFilePath As String
    ' �t�@�C���p�X + �t�@�C����
    strFilePath = ""
    strFilePath = strFilePath + folderPath
    strFilePath = strFilePath + "\"
    strFilePath = strFilePath + file
    ' Bom���폜
    Dim myStream As Object
    Set myStream = CreateObject("ADODB.Stream")
    myStream.Type = 2
    myStream.Charset = "UTF-8"
    myStream.Open
    myStream.WriteText fileContent
    Dim byteData() As Byte
    myStream.Position = 0
    myStream.Type = 1
    myStream.Position = 3
    byteData = myStream.Read
    myStream.Close
    myStream.Open
    myStream.Write byteData
    myStream.SaveToFile strFilePath, 2
    CreateFileWithoutBom = 1
End Function

' �t�@�C���p�X�̎擾
' ���݂��Ȃ��ꍇ�A�����I��
' �p�����^ : �m�F�Ώۃp�X
' �߂�l   : ���͓��e
Public Function getFolderPath(ByVal pathName As String) As String

    ' �o�̓p�X(��΃p�X)�̃Z��������
    Dim cellPath As Range: Set cellPath = Range(searchCell(pathName))
    ' �t�H���_�p�X
    Dim strFilePath As String: strFilePath = Cells(cellPath.Row + 1, cellPath.Column).value

    ' �o�̓t�H���_�����݂��Ȃ��ꍇ
    If Dir(strFilePath, vbDirectory) = "" Then
            MkDir strFilePath
    End If
    
    getFolderPath = strFilePath

End Function

' �P�[�X���̎擾
' Y���̃Z������AY���̋󔒃Z���܂Ōv�Z(���͑ΏۃZ���͑ΏۊO)
' �p�����^ : �Z����
' �߂�l   : �J�E���g��
Public Function getCountCase(ByVal cellName As String, Optional plsuOffset As Integer = 0) As Integer

    Dim getSeiseiCell As String: getSeiseiCell = searchCell(cellName)
    ' �J�n�J����
    Dim startCol As Integer: startCol = Range(getSeiseiCell).Column
    ' �J�nROW
    Dim startRow As Integer: startRow = Range(getSeiseiCell).Row + plsuOffset
    ' �P�[�X��
    Dim countCase As Integer: countCase = 0

    ' ���̒l�����݂��Ȃ��܂ŁA�擾
    Do While Len(Cells(startRow, startCol).value) > 0
        countCase = countCase + 1
        startRow = startRow + 1
    Loop
    ' �P�[�X����߂�
    getCountCase = countCase

End Function

' �O���[�v���e���擾
' �p�����^ : �J�n�Z���A���[�N�V�[�g���A�J�n�Z���܂߃t���O
' �߂�l   : Dictionary(�L�[ : �J�n�Z����̓��e, �l : �J�n�Z����ȍ~�̓��e)
Public Function getGroupList(ByVal cellContentStart As String, Optional worksheetName As String = "", Optional parameterIncludeSelf As Boolean = False, Optional offset As Integer = 1) As Object
    ' �Ώۃ��[�N�V�[�g
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(worksheetName)
    ' �߂�l
    Dim retOut   As Object: Set retOut = createDictionary
    ' �J�n�Z�����擾
    Dim cellStart As Range: Set cellStart = s_workSheet.Range(searchCell(cellContentStart, worksheetName))
    ' �O���[�v�̉E���܂ł̒������擾
    Dim lenArray As Integer: lenArray = getCountHorizon(cellStart.Address, worksheetName)
    ' ����p�Ȕz��
    Dim arryStr() As String
    ' �J�n�Z�����܂߂�ꍇ�A�C���f�b�N�X���𑫂�1
    If parameterIncludeSelf = True Then
        ReDim arryStr(lenArray - 1)
    Else
        ReDim arryStr(lenArray - 2)
    End If
    ' ���[�v�p�J�E���g 1 ����v�Z
    Dim cntLoop As Integer: cntLoop = offset

    ' ���݃Z��
    Dim cellLoop As Range
    ' Y���J�n�Z������Y���󔒃Z���܂Ń��[�v����
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).value <> ""
        ' ���݃Z���̎擾
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        ' �J�n�Z����X������E���܂ł̐��܂Ń��[�v���A�z��ɑ������
        For lenArray = 0 To getArrayLength(arryStr()) - 1
            arryStr(lenArray) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + 1 + lenArray)
        Next
        ' ��L�[�d���`�F�b�N
        If retOut.Exists(cellLoop.value) = False Then
            retOut.Add cellLoop.value, arryStr()
        End If
        ' ���[�v�J�E���g����1
        cntLoop = cntLoop + 1
    Loop
    ' �߂�l�̐ݒ�
    Set getGroupList = retOut

End Function

' �t�@�C�����쐬
' �p�����^ : �t�@�C�����A�t�@�C����ʁA(Optional)�ǉ����e
' �߂�l   : �t�@�C���̐�Ζ���
Public Function getFileName(ByVal fileName As String, fileType As String, Optional fileOptNm As String = "") As String
    ' �t�@�C���p�X�̐錾
    Dim strFilePath As String
    strFilePath = ""
    strFilePath = strFilePath + "/"
    strFilePath = strFilePath + fileName
    ' �ǉ����e�����݂���ꍇ
    If fileOptNm <> "" Then
        strFilePath = strFilePath + fileOptNm
    End If
    strFilePath = strFilePath + fileType
    getFileName = strFilePath
End Function

' ���[�v���ɂ��AY�����X�g�擾�A�ΏۊO�̏ꍇ�X�L�b�v����
' �ΏۊO : �� / �y-�z
' �p�����^ : �J�n�Z���A�h���X�A���[�v���A(Optional) �J�n�I�t�Z�b�g�ǉ�
' �߂�l   : Dictionary(�L�[ : �Z���A�h���X, �l : �Z�����e)
Public Function getVertivalListbyCnt(ByVal cellAdressStart As String, cntTotalLoop As Integer, Optional rowPlus As Integer) As Object
    ' �߂�l
    Dim retDic As Object: Set retDic = createDictionary
    ' �W���uID�̃J����
    Dim colJobId As Range: Set colJobId = Range(cellAdressStart)
    ' ���[�v�Ώۂ̃Z��
    Dim targetloop As Range
    '�ǉ�Row��������
    If rowPlus < 1 Then
        rowPlus = 0
    End If
    ' ���[�v�J�E���g
    Dim cntLoop As Integer
    For cntLoop = 0 To cntTotalLoop - 1
        ' ���[�v�Ώۂ̃Z�����Z�b�g
        Set targetloop = Range(Cells(colJobId.Row + rowPlus + cntLoop, colJobId.Column).Address)
        ' ���[�v�Ώۂ̃Z�����l�L ���� �y-�z�ł͂Ȃ��ꍇ
        If Not (IsEmpty(targetloop.value)) And targetloop.value <> "-" Then
            retDic.Add targetloop.Address, targetloop.value
        End If
    Next

    Set getVertivalListbyCnt = retDic

End Function

' �W���u�J�^���O���X�g�̎擾(���[�v���ɂ��AX���̃��X�g���擾�A�ΏۊO�̏ꍇ�X�L�b�v����)
' �ΏۊO : ��
' �p�����^ : �J�n�Z�����e�A���[�v��
' �߂�l   : ������z��
Public Function getTitleList(ByVal cellContentStart As String, cntLoop As Integer) As String()

    ' �J�n�Z��������
    Dim cellTarget As Range: Set cellTarget = Range(searchCell(cellContentStart))
    ' ���[�v����������
    Dim countLoop As Integer: countLoop = 0
    ' �o�͔z��̒�`
    Dim retList() As String: ReDim retList(0)
    ' �o�͔z��T�C�Y�̒�`
    Dim sizeList As Integer
    ' ���[�v�őΏۂ��擾
    For countLoop = 0 To cntLoop
        ' �ΏۃZ�����󔒂ł͂Ȃ��ꍇ
        If Len(Cells(cellTarget.Row, cellTarget.Column + countLoop).value) > 0 Then
            ' �o�͔z��̒������擾
            sizeList = getArrayLength(retList())
            ' �o�͔z��ɍŌ�̒l���󔒂ł͂Ȃ��ꍇ
            If Len(retList(sizeList - 1)) > 0 Then
                ' �o�͔z����Ē�`����i�ȑO�̒l�͎c��j
                ReDim Preserve retList(sizeList)
                sizeList = getArrayLength(retList())
            End If
                ' �Ώۂ�ǉ�
                retList(sizeList - 1) = Cells(cellTarget.Row, cellTarget.Column + countLoop).value
        End If
    Next
    ' �߂�l�̐ݒ�
    getTitleList = retList()
End Function

' X�����X�g�����J�E���g����i�J�n�Z���͊܂߂Ȃ��j
' �d���`�F�b�N����
' �p�����^ : �J�n�Z�����e
' �߂�l   : �J�E���g��
Public Function getCountKomoku(ByVal cellContentStart As String) As Integer
    ' �����ΏۃZ��
    Dim cellTarget As Range: Set cellTarget = Range(searchCell(cellContentStart))
    ' ���ڐ�
    Dim countTotalContent As Integer
    ' ���[�v��
    Dim countLoop As Integer: countTotalContent = 0
    '  ���ږ���Dic
    Dim dicKomoku As Object: Set dicKomoku = createDictionary
    ' �ΏۃZ��
    Dim komokuCell As Range
    ' ���[�v����������
    countLoop = 0
    ' ���[�v�ō��ڐ����擾
    Do While Len(Cells(cellTarget.Row + 1, cellTarget.Column + countLoop).value) > 0
        Set komokuCell = Cells(cellTarget.Row + 1, cellTarget.Column + countLoop)
        If dicKomoku.Exists(komokuCell.value) Then
            MsgBox "����ID���d�����Ă��܂��B" _
                    , vbYes + vbExclamation, "�ُ�"
            End
        End If
        countTotalContent = countTotalContent + 1
        countLoop = countLoop + 1
        dicKomoku.Add komokuCell.value, komokuCell.Address
    Loop
    ' ���ڐ���߂�
    getCountKomoku = countTotalContent

End Function

' �������PadLeft
' �p�����^ : �C���O������
' �߂�l   : �C���㕶����
Public Function padLeftString(ByVal str As String, ByVal char As String, ByVal digit As Long) As String
  Dim tmp As String: tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Right(String(digit, char) & str, digit)
  End If
  padLeftString = tmp
End Function

' �������PadRight
' �p�����^ : �C���O������
' �߂�l   : �C���㕶����
Public Function padRightString(ByVal str As String, ByVal char As String, ByVal digit As Long) As String
  Dim tmp As String: tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Left(str & String(digit, char), digit)
  End If
  padRightString = tmp
End Function

' �O���[�v���e���擾
' �p�����^ : �J�n�Z���A���[�N�V�[�g���A�J�n�Z���܂߃t���O
' �߂�l   : Dictionary(�L�[ : �J�n�Z���A�h���X, �l : �J�n�Z����ȍ~�̓��e)
Public Function getGroupListbySelectedValue(ByVal cellContentStart As String, _
                                            Optional worksheetName As String = "", _
                                            Optional parameterIncludeSelf As Boolean = False, _
                                            Optional includeOnly As String = "", _
                                            Optional limitCount As Integer = 0, _
                                            Optional offset As Integer = 1) As Object
    ' ���[�N�V�[�g�̎擾
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(worksheetName)
    Dim retOut   As Object: Set retOut = createDictionary
    ' �J�n�Z�����擾
    Dim cellStart As Range: Set cellStart = s_workSheet.Range(searchCell(cellContentStart, worksheetName))
    ' �o�͔z��̐ݒ�
    Dim arryStr() As String
    ' �o�͔z�񒷂��̐ݒ�
    Dim lenOutput As Integer
    ' �z�񒷂��̎w�肪���݂���ꍇ
    If limitCount > 0 Then
        lenOutput = limitCount
        If parameterIncludeSelf = True Then
            ReDim arryStr(lenOutput - 1)
        Else
            ReDim arryStr(lenOutput - 2)
        End If
    Else
        lenOutput = getCountHorizon(cellStart.Address, worksheetName)
        If parameterIncludeSelf = True Then
            ReDim arryStr(lenOutput - 1)
        Else
            ReDim arryStr(lenOutput - 2)
        End If
    End If

    Dim cntLoop As Integer: cntLoop = offset
    Dim cellLoop As Range
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).value <> ""
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        If cellLoop.value = includeOnly Then
            For lenOutput = 0 To getArrayLength(arryStr()) - 1
                arryStr(lenOutput) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + lenOutput)
            Next
            retOut.Add cellLoop.Address, arryStr()
        End If
            cntLoop = cntLoop + 1
    Loop

    Set getGroupListbySelectedValue = retOut

End Function

' �p�����^ : �J�n�Z���A���[�N�V�[�g���A�J�n�Z���܂߃t���O
' �߂�l   : Dictionary(�L�[ : �J�n�Z���A�h���X, �l : �J�n�Z����ȍ~�̓��e)
Public Function getListDictionary(ByVal setworksheet As String, listTitle As String, Optional offset As Integer = 1) As Object

    Dim retOut   As Object: Set retOut = createDictionary
    Dim sworkSheet As Worksheet: Set sworkSheet = getWorSheet(setworksheet)
    Dim cnt As Integer: cnt = offset
    ' ������ʃ��X�g���擾
    Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = sworkSheet.Range(searchCell(listTitle, setworksheet))

    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column).value <> ""
        Set Cellcur = sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column)
        If retOut.Exists(Cellcur.value) = False Then
            retOut.Add Cellcur.value, sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column + 1)
        End If
        cnt = cnt + 1
    Loop

    Set getListDictionary = retOut

End Function

' �p�����^ : �Ȃ�
' �߂�l   : Dictionary
Public Function createDictionary() As Object
    Set createDictionary = CreateObject("Scripting.Dictionary")
End Function

' �p�����^ : Grp�L�[�A�L�[�A���[�N�V�[�g
' �߂�l   : �ΏۃL�[��Item
Public Function getValueByKeyFromDictionary(ByVal DictGrpKey As String, key As String, Optional worksheetName As String = "") As String

        Dim dict As Object: Set dict = getGroupList(DictGrpKey, worksheetName, True)             ' Acms�t�@�C��ID���X�g���擾
        getValueByKeyFromDictionary = dict.item(key)(0)

End Function


' �t�H���_���e�̍폜
Public Function deleteAllFileFromFolder(ByVal folderPath As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fl As Object
    
    Set fl = fso.GetFolder(folderPath) ' �t�H���_���擾
    
    Dim f As Object
    For Each f In fl.Files ' �t�H���_���̃t�@�C�����擾
        f.Delete (True)         ' �t�@�C�����폜
    Next
    
    ' ��n��
    Set fso = Nothing

End Function


