Attribute VB_Name = "SetPage"

Option Explicit

Public Const C8 As String = "C8"
Public Const TITLE_WORKSHEET_PATH_SETTING As String = "�����ق̂���ݒ�ɂ���"

Public Sub SetPageMethod()

    ' �e�[�u�����̃Z�����擾
    Dim tableNameRange As Range: Set tableNameRange = Range(searchCell("�e�[�u����"))
    ' ���ږ����X�g���擾
    Dim listItemsName As Object: Set listItemsName = getGroupListbySelectedValue(Cells(tableNameRange.Row + 5, tableNameRange.Column).Value)

    ' ���[�v�p�̃L�[��錾
    Dim loopKey As Variant
    ' ���ڐ��̎擾
    Dim CntCol As Integer
    ' �ꌏ�߂̂ݎ擾
    For Each loopKey In listItemsName.keys
        Dim ary() As String
        ary() = listItemsName.Item(loopKey)
        CntCol = getArrayLength(ary())
        Exit For
    Next

    ' ��`�p�p�X�̐ݒ���擾
    Dim listSetpath As Object: Set listSetpath = getListDictionaryAsAddress(getWorSheet(TITLE_WORKSHEET_PATH_SETTING).Range(C8), TITLE_WORKSHEET_PATH_SETTING)
    Dim setArray() As String: ReDim setArray(listSetpath.count - 1)
    Dim cntLoop As Integer: cntLoop = 0
    For Each loopKey In listSetpath.keys
        setArray(cntLoop) = listSetpath.Item(loopKey)
        cntLoop = cntLoop + 1
    Next

    ' �v���_�E���̓��e��ǉ�
    Dim listSet() As Variant: listSet() = Array("user", "current_timestamp", "�� NULL ��")

    ' ��`�p�p�X�̐ݒ� �� �ǉ��v���_�E���̓��e����������
    setArray() = Split(Join(setArray(), ",") + "," + Join(listSet(), ","), ",")

    Dim cntCase As Integer
    cntCase = getCountCase("����ID", "", 1)

    ' Y���C���̃J�E���g
    Dim cntY As Integer: cntY = 0
    ' X���C���̃J�E���g
    Dim cntX As Integer: cntX = 0
    ' �ݒ���e
    Dim str As String: str = Join(setArray(), ",")

    Dim backFlg As Boolean: backFlg = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim rangeNow As Range

    Dim rangeStart As String
    rangeStart = Cells(tableNameRange.Row + 6, tableNameRange.Column).Address
    Dim rangeEnd As String
    rangeEnd = Cells(tableNameRange.Row + 6 + cntCase - 1, tableNameRange.Column + CntCol - 1).Address

    Dim rangeUse As String: rangeUse = rangeStart + " : " + rangeEnd

    Set rangeNow = Range(rangeUse)
    rangeNow.Validation.Delete
    rangeNow.Validation.Add Type:=xlValidateList, Formula1:=str
    rangeNow.Validation.ShowError = False

    Application.ScreenUpdating = backFlg

End Sub

Public Function setSqlPage()

    Dim cellFlg As Range: Set cellFlg = Range(searchCell("�쐬�t���O"))
    Dim cntLoop As Integer: cntLoop = getCountCase("�V�[�g���i�K�{�j", "", 0)

    Dim backFlg As Boolean: backFlg = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim cnt As Integer

    For cnt = 0 To cntLoop - 1

        Dim rangeNow As Range: Set rangeNow = Range(Cells(cellFlg.Row + 1 + cnt, cellFlg.Column).Address)
        rangeNow.Validation.Delete
        rangeNow.Validation.Add Type:=xlValidateList, Formula1:="��"

    Next
    Application.ScreenUpdating = backFlg

    setSqlPage = 0

End Function


