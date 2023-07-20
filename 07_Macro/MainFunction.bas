Attribute VB_Name = "MainFunction"

Option Explicit

Public Const TITLE_SHEETNAME As String = "シート名（必須）"                                 ' 【シート名（必須）】定義
Public Const TITLE_OUTPUT_PATH As String = "出力パス"                                 ' 【出力パス】定義
Public Const TITLE_OUTPUT_FLG As String = "作成フラグ"                                 ' 【作成フラグ】定義
Public Const TITLE_TABLE_NAME As String = "テーブル名"                                 ' 【テーブル名】定義
Public Const TITLE_KOMOKU_ID As String = "項目ID"                                     ' 【項目ID】定義

Public Sub CallMacro()

    CreateInsertSql

End Sub

' メイン処理
Public Function CreateInsertSql()

    Dim TitleSheetName As Range: Set TitleSheetName = Range(searchCell(TITLE_SHEETNAME))

    Dim TitleOutputPath As Range: Set TitleOutputPath = Range(searchCell(TITLE_OUTPUT_PATH))

    Dim TitleOutputFlag As Range: Set TitleOutputFlag = Range(searchCell(TITLE_OUTPUT_FLG))

    Dim listSheetName As Object: Set listSheetName = getGroupListbySelectedValue(TITLE_SHEETNAME)

    Dim folderPath As String

    Dim curkey As Variant
     ' 設定したDic内容にて、ループする
    For Each curkey In listSheetName.keys
        If (listSheetName.Item(curkey)(2) = "○") Then
            ' フォルダパスの取得
            folderPath = getFolderPath(TITLE_OUTPUT_PATH)
            ' シート名の取得
            Dim sheetName As String: sheetName = listSheetName.Item(curkey)(0)
            ' シート名を書いたセルを取得
            Dim rangeTableName As Range: Set rangeTableName = Range(searchCell(TITLE_TABLE_NAME, sheetName))
            ' TBL名セルを取得
            Dim tableName As String: tableName = getWorSheet(sheetName).Cells(rangeTableName.Row, rangeTableName.Column + 1).Value
            ' ワークシートObjを取得
            Dim targetWorkSheet As Worksheet: Set targetWorkSheet = getWorSheet(sheetName)
            ' 項目名リストを取得
            Dim listItemsName As Object: Set listItemsName = getGroupListbyCellAddress(targetWorkSheet.Cells(rangeTableName.Row + 5, rangeTableName.Column).Address, sheetName)

            ' SQL文の定義
            Dim sql As String:   sql = ""

            sql = sql + "INSERT INTO "
            sql = sql + tableName
            sql = sql + "("
            ' SQL文の定義
            Dim firstKey As Variant
            ' 一件めのみ取得
            ' 項目名
            For Each firstKey In listItemsName.keys
                sql = sql + Join(listItemsName.Item(firstKey), ", ")
                Exit For
            Next

            ' SQL文の定義
            sql = sql + ") values("
            ' 項目内容の値を取得
            Dim listValue As Object: Set listValue = getGroupListbyCellAddress(targetWorkSheet.Cells(rangeTableName.Row + 5, rangeTableName.Column).Address, sheetName, True, False)
            ' 結果SQLを格納する配列を定義(複数)
            Dim sqlList() As String: ReDim sqlList(0)
            If (listValue.count > 0) Then
                 ReDim sqlList(listValue.count - 1)
            End If
            ' ループカウントの定義
            Dim cntLoop As Integer: cntLoop = 0
            ' 特殊処理の定義
            Dim checkExist() As Variant: checkExist() = Array("user", "current_timestamp", "≪ NULL ≫")

            ' 定義用ファイルパスの設定を取得
            Dim listSetpath As Object: Set listSetpath = getListDictionaryAsAddress(getWorSheet(TITLE_WORKSHEET_PATH_SETTING).Range(C8), TITLE_WORKSHEET_PATH_SETTING)
            Dim pathArray() As String: ReDim pathArray(listSetpath.count - 1)
            cntLoop = 0
            Dim loopKey As Variant
            For Each loopKey In listSetpath.keys
                pathArray(cntLoop) = listSetpath.Item(loopKey)
                cntLoop = cntLoop + 1
            Next
            cntLoop = 0

            ' 入力した内容リストでループする
            For Each firstKey In listValue.keys
                ' 現対象のSQL文
                Dim sqlValue As String: sqlValue = ""
                ' 現対象の値リスト
                Dim listValueDetail() As String: listValueDetail() = listValue.Item(firstKey)
                ' 現在ループ位置（値）
                Dim cntValuelocation As Integer
                ' 現在値
                Dim valueCol As String
                ' 現対象の値リストでループする
                For cntValuelocation = LBound(listValueDetail()) To UBound(listValueDetail())
                    ' 現在値
                    valueCol = Trim(listValueDetail()(cntValuelocation))
                    ' 現在値が特殊設定が必要場合
                    If (checkExistArray(checkExist(), valueCol)) Then
                        ' 現在値が【≪ NULL ≫】の場合
                        If (checkStringEqual(valueCol, "≪ NULL ≫")) Then
                            sqlValue = sqlValue + "NULL"
                        Else
                            sqlValue = sqlValue + valueCol
                        End If
                    ' SQLの場合【”】をつけない
                    ElseIf (isQuery(valueCol, pathArray) = True) Then
                        sqlValue = sqlValue + removeLeftStr(valueCol, 1)
                    Else
                    ' クエリ文の場合【”】をつける
                        sqlValue = sqlValue + "'" + Replace(valueCol, "'", "''") + "'"
                    End If
                    ' 現在値が最後の項目ではない場合
                    If (cntValuelocation <> UBound(listValueDetail())) Then
                        sqlValue = sqlValue + ","
                    End If
                Next
                ' 現対象のSQL文
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

    MsgBox "SQLファイルの作成が完了しました。"

End Function


