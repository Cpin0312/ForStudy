Attribute VB_Name = "MainFunction"

Option Explicit

Public Const TITLE_SHEETNAME As String = "�V�[�g���i�K�{�j"                                 ' �y�V�[�g���i�K�{�j�z��`
Public Const TITLE_OUTPUT_PATH As String = "�o�̓p�X"                                 ' �y�o�̓p�X�z��`
Public Const TITLE_OUTPUT_FLG As String = "�쐬�t���O"                                 ' �y�쐬�t���O�z��`
Public Const TITLE_TABLE_NAME As String = "�e�[�u����"                                 ' �y�e�[�u�����z��`
Public Const TITLE_KOMOKU_ID As String = "����ID"                                     ' �y����ID�z��`

Public Sub CallMacro()

    CreateInsertSql

End Sub

' ���C������
Public Function CreateInsertSql()

    Dim TitleSheetName As Range: Set TitleSheetName = Range(searchCell(TITLE_SHEETNAME))

    Dim TitleOutputPath As Range: Set TitleOutputPath = Range(searchCell(TITLE_OUTPUT_PATH))

    Dim TitleOutputFlag As Range: Set TitleOutputFlag = Range(searchCell(TITLE_OUTPUT_FLG))

    Dim listSheetName As Object: Set listSheetName = getGroupListbySelectedValue(TITLE_SHEETNAME)

    Dim folderPath As String

    Dim curkey As Variant
     ' �ݒ肵��Dic���e�ɂāA���[�v����
    For Each curkey In listSheetName.keys
        If (listSheetName.Item(curkey)(2) = "��") Then
            ' �t�H���_�p�X�̎擾
            folderPath = getFolderPath(TITLE_OUTPUT_PATH)
            ' �V�[�g���̎擾
            Dim sheetName As String: sheetName = listSheetName.Item(curkey)(0)
            ' �V�[�g�����������Z�����擾
            Dim rangeTableName As Range: Set rangeTableName = Range(searchCell(TITLE_TABLE_NAME, sheetName))
            ' TBL���Z�����擾
            Dim tableName As String: tableName = getWorSheet(sheetName).Cells(rangeTableName.Row, rangeTableName.Column + 1).Value
            ' ���[�N�V�[�gObj���擾
            Dim targetWorkSheet As Worksheet: Set targetWorkSheet = getWorSheet(sheetName)
            ' ���ږ����X�g���擾
            Dim listItemsName As Object: Set listItemsName = getGroupListbyCellAddress(targetWorkSheet.Cells(rangeTableName.Row + 5, rangeTableName.Column).Address, sheetName)

            ' SQL���̒�`
            Dim sql As String:   sql = ""

            sql = sql + "INSERT INTO "
            sql = sql + tableName
            sql = sql + "("
            ' SQL���̒�`
            Dim firstKey As Variant
            ' �ꌏ�߂̂ݎ擾
            ' ���ږ�
            For Each firstKey In listItemsName.keys
                sql = sql + Join(listItemsName.Item(firstKey), ", ")
                Exit For
            Next

            ' SQL���̒�`
            sql = sql + ") values("
            ' ���ړ��e�̒l���擾
            Dim listValue As Object: Set listValue = getGroupListbyCellAddress(targetWorkSheet.Cells(rangeTableName.Row + 5, rangeTableName.Column).Address, sheetName, True, False)
            ' ����SQL���i�[����z����`(����)
            Dim sqlList() As String: ReDim sqlList(0)
            If (listValue.count > 0) Then
                 ReDim sqlList(listValue.count - 1)
            End If
            ' ���[�v�J�E���g�̒�`
            Dim cntLoop As Integer: cntLoop = 0
            ' ���ꏈ���̒�`
            Dim checkExist() As Variant: checkExist() = Array("user", "current_timestamp", "�� NULL ��")

            ' ��`�p�t�@�C���p�X�̐ݒ���擾
            Dim listSetpath As Object: Set listSetpath = getListDictionaryAsAddress(getWorSheet(TITLE_WORKSHEET_PATH_SETTING).Range(C8), TITLE_WORKSHEET_PATH_SETTING)
            Dim pathArray() As String: ReDim pathArray(listSetpath.count - 1)
            cntLoop = 0
            Dim loopKey As Variant
            For Each loopKey In listSetpath.keys
                pathArray(cntLoop) = listSetpath.Item(loopKey)
                cntLoop = cntLoop + 1
            Next
            cntLoop = 0

            ' ���͂������e���X�g�Ń��[�v����
            For Each firstKey In listValue.keys
                ' ���Ώۂ�SQL��
                Dim sqlValue As String: sqlValue = ""
                ' ���Ώۂ̒l���X�g
                Dim listValueDetail() As String: listValueDetail() = listValue.Item(firstKey)
                ' ���݃��[�v�ʒu�i�l�j
                Dim cntValuelocation As Integer
                ' ���ݒl
                Dim valueCol As String
                ' ���Ώۂ̒l���X�g�Ń��[�v����
                For cntValuelocation = LBound(listValueDetail()) To UBound(listValueDetail())
                    ' ���ݒl
                    valueCol = Trim(listValueDetail()(cntValuelocation))
                    ' ���ݒl������ݒ肪�K�v�ꍇ
                    If (checkExistArray(checkExist(), valueCol)) Then
                        ' ���ݒl���y�� NULL ��z�̏ꍇ
                        If (checkStringEqual(valueCol, "�� NULL ��")) Then
                            sqlValue = sqlValue + "NULL"
                        Else
                            sqlValue = sqlValue + valueCol
                        End If
                    ' SQL�̏ꍇ�y�h�z�����Ȃ�
                    ElseIf (isQuery(valueCol, pathArray) = True) Then
                        sqlValue = sqlValue + removeLeftStr(valueCol, 1)
                    Else
                    ' �N�G�����̏ꍇ�y�h�z������
                        sqlValue = sqlValue + "'" + Replace(valueCol, "'", "''") + "'"
                    End If
                    ' ���ݒl���Ō�̍��ڂł͂Ȃ��ꍇ
                    If (cntValuelocation <> UBound(listValueDetail())) Then
                        sqlValue = sqlValue + ","
                    End If
                Next
                ' ���Ώۂ�SQL��
                sqlList(cntLoop) = sql + sqlValue + ");"
                cntLoop = cntLoop + 1
            Next

            If (getArrayLength(sqlList) > 0) Then

                Dim outputSql As String
                outputSql = Join(sqlList, vbLf)

                Dim ret As Integer

                Dim startLineSql As String
                startLineSql = "/* delete */" + vbLf
                startLineSql = startLineSql + "DELETE FROM " + tableName + ";" + vbLf
                startLineSql = startLineSql + "/* insert */" + vbLf
                outputSql = startLineSql + outputSql + vbLf

                ret = CreateFileWithoutBom(folderPath, tableName + ".sql", outputSql)

            End If

        End If

    Next

    MsgBox "SQL�t�@�C���̍쐬���������܂����B"

End Function


