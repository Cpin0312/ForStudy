Attribute VB_Name = "AddOnFunction"
' 初期データ作成処理
Public Function CreateInitialData()

    If ActiveSheet.name = "SQL作成" Then

        CallMacro

    Else

        MsgBox "シート【SQL作成】に移動して、実行してください"

    End If

End Function

' バージョン表示
Public Function ShowVersion()

    Dim msg As String

    msg = msg + "Version 0.1 : 新規作成" + vbCrLf
    msg = msg + "Version 0.2 : 初期ページ作成機能を追加" + vbCrLf

    MsgBox msg

End Function

' 初期ページ作成処理
Public Function CreateInitialPage()

    Dim ws As Worksheet
    Dim flag As Boolean
    Dim createList() As Variant
    createList() = Array("変更履歴", "SQL作成", "使用方法の説明", "環境差異のある設定について")

    For Each ws In ActiveWorkbook.Worksheets
        flag = False

        If ws.name = "変更履歴" Or ws.name = "SQL作成" Or ws.name = "使用方法の説明" Or ws.name = "環境差異のある設定について" Then
            createList() = removeItemFromArray(createList(), ws.name)
        End If

    Next ws

    If createList(0) <> Empty Then
        Dim nameSheet As Variant
        For Each nameSheet In createList
            ThisWorkbook.Worksheets(nameSheet).Copy After:=ActiveWorkbook.Worksheets(Worksheets.count)

            If nameSheet = "SQL作成" Then

                 setSqlPage

            End If

        Next
        NewInitialPage
        MsgBox "作成完了しました。"
    Else

        MsgBox "作成可能シートがありません。"

    End If

End Function

' ページ追加処理
Public Function NewInitialPage()

    Dim ws As Worksheet
    Dim flag As Boolean
    Dim newSheetName As String: newSheetName = "サンプルテーブル"

    For Each ws In ActiveWorkbook.Worksheets
        flag = False

        If ws.name = newSheetName Then
            MsgBox newSheetName & "がすでに存在しています、作成できません。"
            flag = True
            Exit For
        End If

    Next ws

    If flag = False Then

        ThisWorkbook.Worksheets(newSheetName).Copy After:=ActiveWorkbook.Worksheets(Worksheets.count)
        SetPageMethod

    End If

End Function


