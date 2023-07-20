Attribute VB_Name = "K_CMN_METHOD"
Option Explicit

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
Public Function getWorkSheet(ByVal selectSheet As String) As Worksheet
        ' ���[�N�V�[�g�̒�`
    Dim s_workSheet As Worksheet
    ' ���[�N�V�[�g�����͂���Ȃ��ꍇ�A���݂̃��[�N�V�[�g�Ƃ���
    If selectSheet <> "" Then
        Set s_workSheet = Worksheets(selectSheet)
    Else
        Set s_workSheet = ActiveSheet
    End If
    Set getWorkSheet = s_workSheet
End Function

' ����̐����擾(�w���̎w��Z������AX���̋󔒃Z���܂Ōv�Z)
' �p�����^ : �J�n�Z���A�h���X�A(Optional)�ΏۃV�[�g��
' �߂�l   : �J�E���g��
Public Function getCountHorizon(ByVal cellAddressStart As String, Optional nameWorkSheet As String = "") As Integer

    ' ���[�N�V�[�g�̒�`
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorkSheet(nameWorkSheet)
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
    Dim sworkSheet As Worksheet: Set sworkSheet = getWorkSheet(nameWorkSheet)
    ' �J�n�Z���̎擾
    Dim cellStart As Range: Set cellStart = sworkSheet.Range(searchCell(cellContent, nameWorkSheet))
    ' ' �J�n�Z���܂߂Ȃ����߁A�I�t�Z�b�g�y1�z����J�n����
    Dim cntLoop As Integer: cntLoop = offset
    ' ���[�v�Z���̎擾
    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).Value <> ""
        ' ���[�v�Z���̐ݒ�
        Set Cellcur = sworkSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        ' ��L�[�d���`�F�b�N�̐ݒ�
        If retOut.Exists(Cellcur.Value) = False Then
            retOut.Add Cellcur.Address, Cellcur.Value
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
Public Function searchCell(ByVal cellContent As String, Optional workSheetName As String = "") As String
    ' ���[�N�V�[�g�̒�`
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorkSheet(workSheetName)
    ' �����Ώۂ��擾
    Dim retRange As Range: Set retRange = s_workSheet.Cells.Find(cellContent, LookAt:=xlWhole)
    ' �����Ώۂ����݂��Ȃ��ꍇ�A�����I��
    If (retRange Is Nothing) Then
        showMsg cellContent & "��������܂���B" & vbCrLf & "�V�F���t�@�C�����쐬�ł��܂���B" _
                , vbYes + vbExclamation, "�ُ�"
        End
    End If
    ' �����Ώۂ̃Z����Ԃ�
    searchCell = retRange.Address

End Function

' �Z���̑��݊m�F
' �p�����^ : �������e�A�Ώۃ��[�N�V�[�g��
' �߂�l   : �Z���A�h���X
Public Function isCellExist(ByVal cellContent As String, Optional workSheetName As String = "") As Boolean
    isCellExist = False
    ' ���[�N�V�[�g�̒�`
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorkSheet(workSheetName)
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
    Dim strFilePath As String: strFilePath = Cells(cellPath.Row + 1, cellPath.Column).Value

    ' �o�̓t�H���_�����݂��Ȃ��ꍇ
    If Dir(strFilePath, vbDirectory) = "" Then
        showMsg strFilePath & vbCrLf & "�����݂��Ă��܂���B" _
                , vbYes + vbExclamation, "�ُ�"
        End
    End If
    getFolderPath = strFilePath

End Function

' �P�[�X���̎擾
' Y���̃Z������AY���̋󔒃Z���܂Ōv�Z(���͑ΏۃZ���͑ΏۊO)
' �p�����^ : �Z����
' �߂�l   : �J�E���g��
Public Function getCountCase(ByVal cellName As String, Optional workSheetName As String = "", Optional plsuOffset As Integer = 0) As Integer
    ' �J�n�J����
    Dim startCol As Integer
    ' �J�nROW
    Dim startRow As Integer
    ' �P�[�X��
    Dim countCase As Integer

    Dim Worksheet As Worksheet
    Set Worksheet = getWorkSheet(workSheetName)

    getSeiseiCell = searchCell(cellName, workSheetName)
    ' �J�n�J������ݒ�
    startCol = Worksheet.Range(getSeiseiCell).Column
    ' �J�nROW��ݒ�
    startRow = Worksheet.Range(getSeiseiCell).Row + plsuOffset
    ' �P�[�X����������
    countCase = 0
    ' ���̒l�����݂��Ȃ��܂ŁA�擾
    Do While Len(Worksheet.Cells(startRow, startCol).Value) > 0
        countCase = countCase + 1
        startRow = startRow + 1
    Loop
    ' �P�[�X����߂�
    getCountCase = countCase

End Function

' �O���[�v���e���擾
' �p�����^ : �J�n�Z���A���[�N�V�[�g���A�J�n�Z���܂߃t���O
' �߂�l   : Dictionary(�L�[ : �J�n�Z����̓��e, �l : �J�n�Z����ȍ~�̓��e)
Public Function getGroupList(ByVal cellContentStart As String, Optional workSheetName As String = "", Optional parameterIncludeSelf As Boolean = False, Optional offset As Integer = 1) As Object
    ' �Ώۃ��[�N�V�[�g
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorkSheet(workSheetName)
    ' �߂�l
    Dim retOut   As Object: Set retOut = createDictionary
    ' �J�n�Z�����擾
    Dim cellStart As Range: Set cellStart = s_workSheet.Range(searchCell(cellContentStart, workSheetName))
    ' �O���[�v�̉E���܂ł̒������擾
    Dim lenArray As Integer: lenArray = getCountHorizon(cellStart.Address, workSheetName)
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
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).Value <> ""
        ' ���݃Z���̎擾
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        ' �J�n�Z����X������E���܂ł̐��܂Ń��[�v���A�z��ɑ������
        For lenArray = 0 To getArrayLength(arryStr()) - 1
            arryStr(lenArray) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + 1 + lenArray)
        Next
        ' ��L�[�d���`�F�b�N
        If retOut.Exists(cellLoop.Value) = False Then
            retOut.Add cellLoop.Value, arryStr()
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

' �p�����^ : �J�n�Z���A���[�N�V�[�g���A�J�n�Z���܂߃t���O
' �߂�l   : Dictionary(�L�[ : �J�n�Z���A�h���X, �l : �J�n�Z����ȍ~�̓��e)
Public Function getListDictionary(ByVal listTitle As String, Optional setworksheet As String = "", Optional offset As Integer = 1) As Object

    Dim retOut   As Object: Set retOut = createDictionary
    Dim sworkSheet As Worksheet: Set sworkSheet = getWorkSheet(setworksheet)
    Dim cnt As Integer: cnt = offset
    ' ������ʃ��X�g���擾
    Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = sworkSheet.Range(searchCell(listTitle, setworksheet))

    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column).Value <> ""
        Set Cellcur = sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column)
        If retOut.Exists(Cellcur.Value) = False Then
            retOut.Add Cellcur.Value, sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column + 1)
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
Public Function getValueByKeyFromDictionary(ByVal DictGrpKey As String, key As String, Optional workSheetName As String = "") As String

        Dim dict As Object: Set dict = getGroupList(DictGrpKey, workSheetName, True)             ' Acms�t�@�C��ID���X�g���擾
        getValueByKeyFromDictionary = dict.Item(key)(0)

End Function

' �O���[�v���e���擾
' �p�����^ : �J�n�Z���A���[�N�V�[�g���A�J�n�Z���܂߃t���O
' �߂�l   : Dictionary(�L�[ : �J�n�Z���A�h���X, �l : �J�n�Z����ȍ~�̓��e)
Public Function getGroupListbySelectedValue(ByVal cellContentStart As String, _
                                                    Optional workSheetName As String = "", _
                                                    Optional includeStartCol As Boolean = True, _
                                                    Optional includeStartRow As Boolean = True, _
                                                    Optional offsetRow As Integer = 0, _
                                                    Optional nmFilter As String = "") As Object
    ' ���[�N�V�[�g�̎擾
    Dim s_workSheet As Worksheet
    Set s_workSheet = getWorkSheet(workSheetName)
    Dim retOut   As Object
    Set retOut = createDictionary
    ' �J�n�Z�����擾
    Dim cellStart As Range
    Set cellStart = s_workSheet.Range(searchCell(cellContentStart, workSheetName))
    ' �o�͔z��̐ݒ�
    Dim arryStr() As String
    ' �o�͔z�񒷂��̐ݒ�
    Dim lenOutput As Integer

    lenOutput = getCountHorizon(cellStart.Address, workSheetName)
    If includeStartCol = True Then
        ReDim arryStr(lenOutput - 1)
    Else
        ReDim arryStr(lenOutput - 2)
    End If

    Dim cntLoop As Integer
    cntLoop = 0

    If (includeStartRow = False) Then
        cntLoop = 1
    End If

    If (offsetRow > 0) Then
        cntLoop = cntLoop + offsetRow
    End If

    Dim cellLoop As Range
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).Value <> ""
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
            For lenOutput = 0 To getArrayLength(arryStr()) - 1
                arryStr(lenOutput) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + lenOutput)
            Next
            If (nmFilter <> "") Then
                If (arryStr(0) = nmFilter) Then
                    retOut.Add cellLoop.Address, arryStr()
                End If
            Else
                retOut.Add cellLoop.Address, arryStr()
            End If
            cntLoop = cntLoop + 1
    Loop

    Set getGroupListbySelectedValue = retOut

End Function

Public Function checkExistArray(ByRef ary() As Variant, checkStr As Variant) As Boolean

    Dim retBool As Boolean: retBool = False
    Dim cnt As Integer
    Dim src As String
    Dim target As String
    If (checkStr <> "") Then
        target = convert2Unicode(checkStr)
    Else
        target = convert2Unicode(checkVar)
    End If
    
    For cnt = LBound(ary()) To UBound(ary())
    
        src = convert2Unicode(ary(cnt))
        If src = target Then
            retBool = True
            Exit For
        End If

    Next

    checkExistArray = retBool

End Function

Public Function checkExistArrayStr(ByRef ary() As String, checkStr As String) As Boolean

    Dim retBool As Boolean: retBool = False
    Dim cnt As Integer
    Dim src As String
    Dim target As String
    If (checkStr <> "") Then
        target = convert2Unicode(checkStr)
    Else
        target = convert2Unicode(checkVar)
    End If
    
    For cnt = LBound(ary()) To UBound(ary())
    
        src = convert2Unicode(ary(cnt))
        If src = target Then
            retBool = True
            Exit For
        End If

    Next

    checkExistArray = retBool

End Function


Public Function checkStringEqual(ByVal input1 As String, input2 As String) As Boolean

    Dim retBool As Boolean
    retBool = False

    src = convert2Unicode(input1)
    target = convert2Unicode(input2)

    If src = target Then
        retBool = True
    End If

    checkStringEqual = retBool

End Function

Public Function convert2Unicode(ByVal inputStr As String) As String

    convert2Unicode = StrConv(inputStr, vbFromUnicode)

End Function

' �p�����^ : �J�n�Z���A���[�N�V�[�g���A�J�n�Z���܂߃t���O
' �߂�l   : Dictionary(�L�[ : �J�n�Z���A�h���X, �l : �J�n�Z����ȍ~�̓��e)
Public Function getListDictionaryAsAddress(ByVal listTitle As String, Optional setworksheet As String) As Object

    Dim retOut   As Object
    Set retOut = createDictionary
    Dim sworkSheet As Worksheet
    Set sworkSheet = getWorkSheet(setworksheet)
    Dim cnt As Integer
    ' ������ʃ��X�g���擾
    Dim cellSyoriSyubetsu As Range
    Set cellSyoriSyubetsu = sworkSheet.Range(searchCell(listTitle, setworksheet))

    cnt = 1
    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column).Value <> ""
        Set Cellcur = sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column)
        If retOut.Exists(Cellcur.Value) = False Then
            retOut.Add Cellcur.Address, sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column)
        End If
        cnt = cnt + 1
    Loop

    Set getListDictionaryAsAddress = retOut

End Function

' �z��̒������Z�o
' �p�����^ : �Ώ۔z��
' �߂�l   : �J�E���g��
Public Function getArrayLengthVariant(ByRef arry() As Variant) As Integer
    ' �Ō�̃C���f�b�N - �ŏ��̃C���f�b�N + 1
    getArrayLengthVariant = UBound(arry()) - LBound(arry()) + 1

End Function

Public Function getLeftString(ByVal str As String, count As Integer) As String

    getLeftString = Left(str, count)

End Function

Public Function getRightString(ByVal str As String, count As Integer) As String

    getRightString = Right(str, count)

End Function

Public Function isQuery(ByVal str As String, pathAry() As String) As String
    Dim retBol As Boolean: retBol = False

    Dim skipBol As Boolean: skipBol = False
    Dim cnt As Integer: cnt = 0

    For cnt = LBound(pathAry()) To UBound(pathAry())
        Dim xxx As Integer
        xxx = InStr(str, pathAry(cnt))

        If (InStr(str, pathAry(cnt)) > 0) Then
            skipBol = True
            Exit For
        End If

    Next

    If (skipBol = False) Then

        Dim strStart As String: strStart = getLeftString(str, 1)
        Dim strEnd As String: strEnd = getRightString(str, 1)

        If (strStart = "@" And strEnd <> "@") Then
            retBol = True
        End If

    End If

    isQuery = retBol

End Function

Public Function removeLeftStr(ByVal str As String, cnt As Integer) As String

    removeLeftStr = Mid(str, 1 + cnt)

End Function

' �O���[�v���e���擾
' �p�����^ : �J�n�Z���A���[�N�V�[�g���A�J�n�Z���܂߃t���O
' �߂�l   : Dictionary(�L�[ : �J�n�Z���A�h���X, �l : �J�n�Z����ȍ~�̓��e)
Public Function getGroupListbyCellAddress(ByVal cellAddress As String, _
                                                    Optional workSheetName As String = "", _
                                                    Optional includeStartCol As Boolean = True, _
                                                    Optional includeStartRow As Boolean = True, _
                                                    Optional offsetRow As Integer = 0) As Object
    ' ���[�N�V�[�g�̎擾
    Dim s_workSheet As Worksheet
    Set s_workSheet = getWorkSheet(workSheetName)
    Dim retOut   As Object
    Set retOut = createDictionary
    ' �J�n�Z�����擾
    Dim cellStart As Range
    Set cellStart = s_workSheet.Range(cellAddress)
    ' �o�͔z��̐ݒ�
    Dim arryStr() As String
    ' �o�͔z�񒷂��̐ݒ�
    Dim lenOutput As Integer

    lenOutput = getCountHorizon(cellStart.Address, workSheetName)
    If includeStartCol = True Then
        ReDim arryStr(lenOutput - 1)
    Else
        ReDim arryStr(lenOutput - 2)
    End If

    Dim cntLoop As Integer
    cntLoop = 0

    If (includeStartRow = False) Then
        cntLoop = 1
    End If

    If (offsetRow > 0) Then
        cntLoop = cntLoop + offsetRow
    End If

    Dim cellLoop As Range
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).Value <> ""
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
            For lenOutput = 0 To getArrayLength(arryStr()) - 1
                arryStr(lenOutput) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + lenOutput)
            Next
            retOut.Add cellLoop.Address, arryStr()
            cntLoop = cntLoop + 1
    Loop

    Set getGroupListbyCellAddress = retOut

End Function

Public Function addItemToArray(ByRef targetArry() As Variant, addItem As Variant) As Variant()

    Dim outAry() As Variant: outAry() = targetArry()

    Dim newCnt As Integer: newCnt = getArrayLengthVariant(targetArry())

    If newCnt = 1 And outAry(0) = Empty Then

        outAry(0) = addItem

    Else

        ReDim Preserve outAry(newCnt)

        outAry(newCnt) = addItem

    End If

    addItemToArray = outAry()

End Function

Public Function removeItemFromArray(ByRef targetArray() As Variant, deleteItem As String) As Variant()

    Dim outAry() As Variant: ReDim outAry(0)

    Dim loopStr As Variant

    For Each loopStr In targetArray

        If loopStr <> deleteItem Then

            outAry() = addItemToArray(outAry(), loopStr)

        End If

    Next

    removeItemFromArray = outAry()

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
        If Not (IsEmpty(targetloop.Value)) And targetloop.Value <> "-" Then
            retDic.Add targetloop.Address, targetloop.Value
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
        If Len(Cells(cellTarget.Row, cellTarget.Column + countLoop).Value) > 0 Then
            ' �o�͔z��̒������擾
            sizeList = getArrayLength(retList())
            ' �o�͔z��ɍŌ�̒l���󔒂ł͂Ȃ��ꍇ
            If Len(retList(sizeList - 1)) > 0 Then
                ' �o�͔z����Ē�`����i�ȑO�̒l�͎c��j
                ReDim Preserve retList(sizeList)
                sizeList = getArrayLength(retList())
            End If
                ' �Ώۂ�ǉ�
                retList(sizeList - 1) = Cells(cellTarget.Row, cellTarget.Column + countLoop).Value
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
    Do While Len(Cells(cellTarget.Row + 1, cellTarget.Column + countLoop).Value) > 0
        Set komokuCell = Cells(cellTarget.Row + 1, cellTarget.Column + countLoop)
        If dicKomoku.Exists(komokuCell.Value) Then
            showMsg "����ID���d�����Ă��܂��B" _
                    , vbYes + vbExclamation, "�ُ�"
            End
        End If
        countTotalContent = countTotalContent + 1
        countLoop = countLoop + 1
        dicKomoku.Add komokuCell.Value, komokuCell.Address
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

Public Function saveWorkBookMacro() As String

    saveWorkBook = ""
    '�_�C�A���O�ŕۑ���E�t�@�C�������w��
    Dim strFilePath As String
    strFilePath = Application.GetSaveAsFilename( _
           title:="�ۑ����I�����Ă��������I" _
         , InitialFileName:="initialTBL" _
         , FileFilter:="Excel�}�N���L���u�b�N,*.xlsm")
        
        '�w�肵���p�X�Ƀt�@�C�����쐬�ςłȂ������m�F�B
    If strFilePath <> "False" And Dir(strFilePath) = "" Then
        '�V�����t�@�C�����쐬
        Set newBook = Workbooks.Add
        '�V�����t�@�C����VBA�����s�����t�@�C���Ɠ����t�H���_�ۑ�
        newBook.SaveAs strFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        saveWorkBook = 0
        
    Else
        If strFilePath = "False" Then
            showMsg "�t�@�C���������͂���Ă��܂���B"
        ElseIf Dir(strFilePath) <> "" Then
            
            saveWorkBook = strFilePath
        
        Else
            '���ɓ����̃t�@�C�������݂���ꍇ�̓��b�Z�[�W��\��
            showMsg "����" & newBookName & "�Ƃ����t�@�C���͑��݂��܂��B"
            
        End If
    End If
    
End Function

Public Function saveWorkBook() As String

    saveWorkBook = ""
    '�_�C�A���O�ŕۑ���E�t�@�C�������w��
    Dim strFilePath As String
    strFilePath = Application.GetSaveAsFilename( _
           title:="�ۑ����I�����Ă��������I" _
         , InitialFileName:="initialTBL" _
         , FileFilter:="Excel�u�b�N,*.xlsx")
        
        '�w�肵���p�X�Ƀt�@�C�����쐬�ςłȂ������m�F�B
    If strFilePath <> "False" And Dir(strFilePath) = "" Then
        '�V�����t�@�C�����쐬
        Set newBook = Workbooks.Add
        '�V�����t�@�C����VBA�����s�����t�@�C���Ɠ����t�H���_�ۑ�
        newBook.SaveAs strFilePath, FileFormat:=xlOpenXMLWorkbook
        saveWorkBook = "SUCCESS"
        
    Else
        If strFilePath = "False" Then
            showMsg "�t�@�C���������͂���Ă��܂���B"
        ElseIf Dir(strFilePath) <> "" Then
            
            saveWorkBook = strFilePath
        
        Else
            '���ɓ����̃t�@�C�������݂���ꍇ�̓��b�Z�[�W��\��
            showMsg "����" & newBookName & "�Ƃ����t�@�C���͑��݂��܂��B"
            
        End If
    End If
    
End Function

Public Function delSheet(ByVal sheetName As String)
    Dim loopNm As Worksheet
    For Each loopNm In ActiveWorkbook.Worksheets
        
        If loopNm.name = sheetName Then
            Application.DisplayAlerts = False
            ActiveWorkbook.Worksheets(sheetName).Delete
            Application.DisplayAlerts = True
        End If
    Next
    
End Function

Public Function checkBooksExist(ByVal path As String) As Boolean

    checkBooksExist = False

    If Dir(strFilePath) <> "" Then
        checkBooksExist = True
    End If

End Function


Public Function selectBook() As String
    selectBook = ""
    Dim strFilePath As String
    strFilePath = Application.GetOpenFilename( _
           title:="�ۑ����I�����Ă��������I" _
         , FileFilter:="Excel�}�N���L���u�b�N,*.xlsm")
         
    selectBook = strFilePath

End Function


Public Function showMsg(ByVal msg As String, Optional btn As VbMsgBoxStyle = vbOK, Optional title As String = "") As Integer
    'vbOK 1 [OK]�{�^���������ꂽ
    'vbCancel 2 [�L�����Z��]�{�^���������ꂽ
    'vbAbort 3 [���~]�{�^���������ꂽ
    'vbRetry 4 [�Ď��s]�{�^���������ꂽ
    'vbIgnore 5 [����]�{�^���������ꂽ
    'vbYes 6 [�͂�]�{�^���������ꂽ
    'vbNo 7 [������]�{�^���������ꂽ

    showMsg = MsgBox(msg, btn, title)

End Function



