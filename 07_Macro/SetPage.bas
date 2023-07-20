Attribute VB_Name = "SetPage"

Option Explicit

Public Const C8 As String = "C8"
Public Const TITLE_WORKSHEET_PATH_SETTING As String = "環境差異のある設定について"

Public Sub SetPageMethod()

    ' テーブル名のセルを取得
    Dim tableNameRange As Range: Set tableNameRange = Range(searchCell("テーブル名"))
    ' 項目名リストを取得
    Dim listItemsName As Object: Set listItemsName = getGroupListbySelectedValue(Cells(tableNameRange.Row + 5, tableNameRange.Column).Value)

    ' ループ用のキーを宣言
    Dim loopKey As Variant
    ' 項目数の取得
    Dim CntCol As Integer
    ' 一件めのみ取得
    For Each loopKey In listItemsName.keys
        Dim ary() As String
        ary() = listItemsName.Item(loopKey)
        CntCol = getArrayLength(ary())
        Exit For
    Next

    ' 定義用パスの設定を取得
    Dim listSetpath As Object: Set listSetpath = getListDictionaryAsAddress(getWorSheet(TITLE_WORKSHEET_PATH_SETTING).Range(C8), TITLE_WORKSHEET_PATH_SETTING)
    Dim setArray() As String: ReDim setArray(listSetpath.count - 1)
    Dim cntLoop As Integer: cntLoop = 0
    For Each loopKey In listSetpath.keys
        setArray(cntLoop) = listSetpath.Item(loopKey)
        cntLoop = cntLoop + 1
    Next

    ' プルダウンの内容を追加
    Dim listSet() As Variant: listSet() = Array("user", "current_timestamp", "≪ NULL ≫")

    ' 定義用パスの設定 と 追加プルダウンの内容を結合する
    setArray() = Split(Join(setArray(), ",") + "," + Join(listSet(), ","), ",")

    Dim cntCase As Integer
    cntCase = getCountCase("項目ID", "", 1)

    ' Yラインのカウント
    Dim cntY As Integer: cntY = 0
    ' Xラインのカウント
    Dim cntX As Integer: cntX = 0
    ' 設定内容
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

    Dim cellFlg As Range: Set cellFlg = Range(searchCell("作成フラグ"))
    Dim cntLoop As Integer: cntLoop = getCountCase("シート名（必須）", "", 0)

    Dim backFlg As Boolean: backFlg = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Dim cnt As Integer

    For cnt = 0 To cntLoop - 1

        Dim rangeNow As Range: Set rangeNow = Range(Cells(cellFlg.Row + 1 + cnt, cellFlg.Column).Address)
        rangeNow.Validation.Delete
        rangeNow.Validation.Add Type:=xlValidateList, Formula1:="○"

    Next
    Application.ScreenUpdating = backFlg

    setSqlPage = 0

End Function


