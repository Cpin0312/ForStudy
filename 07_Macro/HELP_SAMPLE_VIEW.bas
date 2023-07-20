Attribute VB_Name = "HELP_SAMPLE_VIEW"
'**helpview() ヘルプシート表示
'**sumpleview()　サンプルシート表示
'**dialogshow() ダイアログ表示


'==========================================================
'【関数名】helpview
'【概　要】ヘルプシートを表示
'【引　数】なし
'【戻り値】なし
'==========================================================

Sub helpview()
    Dim wnum As Integer
    '開かれているワークブックがあるか調べる
    'ワークブックが開かれていない場合ワークブックを追加してシート追加
    If Workbooks.Count = 0 Then
        Workbooks.Add
        Workbooks(TEMPLATE).Worksheets(HELPSHEET).Copy _
        before:=Workbooks(1).Worksheets(1)
        'メッセージをOFFにする
        Application.DisplayAlerts = False
        '余分なシートを削除する
        For i = 2 To Worksheets.Count
            Worksheets(2).Delete
        Next i
        Application.DisplayAlerts = True
    Else
        wnum = HaveSheet(ActiveWorkbook.name, HELPSHEET)
        If wnum > 0 Then
          ActiveWorkbook.Worksheets(wnum).Activate
        Else
            wnum = HaveSheet2(ActiveWorkbook.name, 9)
            If wnum > 0 Then
                ActiveWorkbook.Worksheets(wnum).Activate
            Else
                Workbooks(TEMPLATE).Worksheets(HELPSHEET).Copy _
                before:=Worksheets(1)
            End If
        End If
    End If
End Sub


'==========================================================
'【関数名】sampleview
'【概　要】サンプルシートを表示
'【引　数】なし
'【戻り値】なし
'==========================================================

Sub sampleview()
    Dim wnum As Integer
    '開かれているワークブックがあるか調べる
    'ワークブックが開かれていない場合ワークブックを追加してシート追加
    If Workbooks.Count = 0 Then
        Workbooks.Add
        GV_book = ActiveWorkbook.name
        Workbooks(TEMPLATE).Worksheets(SAMPLESHEET).Copy _
        before:=Workbooks(1).Worksheets(1)
        'メッセージをOFFにする
        Application.DisplayAlerts = False
        '余分なシートを削除する
        For i = 2 To Worksheets.Count
            Worksheets(2).Delete
        Next i
        Application.DisplayAlerts = True
    Else
        wnum = HaveSheet(ActiveWorkbook.name, SAMPLESHEET)
        If wnum > 0 Then
            ActiveWorkbook.Worksheets(wnum).Activate
        Else
            Workbooks(TEMPLATE).Worksheets(SAMPLESHEET).Copy _
            before:=Worksheets(1)
        End If
    End If
End Sub

Sub dialogshow()
    VOYAGERDIALOG.Show
    'ActiveWorkbook.DialogSheets("Dialog1").Show
End Sub

