Attribute VB_Name = "AddOnFunction"
' �����f�[�^�쐬����
Public Function CreateInitialData()

    If ActiveSheet.name = "SQL�쐬" Then

        CallMacro

    Else

        MsgBox "�V�[�g�ySQL�쐬�z�Ɉړ����āA���s���Ă�������"

    End If

End Function

' �o�[�W�����\��
Public Function ShowVersion()

    Dim msg As String

    msg = msg + "Version 0.1 : �V�K�쐬" + vbCrLf
    msg = msg + "Version 0.2 : �����y�[�W�쐬�@�\��ǉ�" + vbCrLf

    MsgBox msg

End Function

' �����y�[�W�쐬����
Public Function CreateInitialPage()

    Dim ws As Worksheet
    Dim flag As Boolean
    Dim createList() As Variant
    createList() = Array("�ύX����", "SQL�쐬", "�g�p���@�̐���", "�����ق̂���ݒ�ɂ���")

    For Each ws In ActiveWorkbook.Worksheets
        flag = False

        If ws.name = "�ύX����" Or ws.name = "SQL�쐬" Or ws.name = "�g�p���@�̐���" Or ws.name = "�����ق̂���ݒ�ɂ���" Then
            createList() = removeItemFromArray(createList(), ws.name)
        End If

    Next ws

    If createList(0) <> Empty Then
        Dim nameSheet As Variant
        For Each nameSheet In createList
            ThisWorkbook.Worksheets(nameSheet).Copy After:=ActiveWorkbook.Worksheets(Worksheets.count)

            If nameSheet = "SQL�쐬" Then

                 setSqlPage

            End If

        Next
        NewInitialPage
        MsgBox "�쐬�������܂����B"
    Else

        MsgBox "�쐬�\�V�[�g������܂���B"

    End If

End Function

' �y�[�W�ǉ�����
Public Function NewInitialPage()

    Dim ws As Worksheet
    Dim flag As Boolean
    Dim newSheetName As String: newSheetName = "�T���v���e�[�u��"

    For Each ws In ActiveWorkbook.Worksheets
        flag = False

        If ws.name = newSheetName Then
            MsgBox newSheetName & "�����łɑ��݂��Ă��܂��A�쐬�ł��܂���B"
            flag = True
            Exit For
        End If

    Next ws

    If flag = False Then

        ThisWorkbook.Worksheets(newSheetName).Copy After:=ActiveWorkbook.Worksheets(Worksheets.count)
        SetPageMethod

    End If

End Function


