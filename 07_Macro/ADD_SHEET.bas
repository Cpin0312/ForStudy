Attribute VB_Name = "ADD_SHEET"
'***NewBookT() 新規ブック作成
'***NewTableSheet() テーブル項目シート挿入
'***DeleteSheet シートのDELETE（不使用）
'***Runstop 処理中断

'==========================================================
'【プロシージャ名】NewBookT
'【概　要】テーブル定義書の追加
'【引　数】なし
'【戻り値】なし
'==========================================================
Sub NewBookT()
    Dim wnum As Integer
    Dim i As Integer
    Dim ans As Integer
    Dim sname As String
    
    'テーブルシート名取得
    Tb_SheetNMInp = ""
    Tb_SheetNMInp = UCase(InputBox("テーブルIDを入力してください"))
    
    If Len(Tb_SheetNMInp) > 31 Then
        MsgBox "テーブルIDを31桁以内で入力してください。"
        Exit Sub
    End If
    
    If Tb_SheetNMInp = "" Then
        Exit Sub
    Else
        'ワークブックが開かれているか調べる
        If Workbooks.Count = 0 Then
            Application.ScreenUpdating = False
            Workbooks().Add
            wnum = Worksheets.Count
            NewTableSheet (ActiveWorkbook.name)
            Application.DisplayAlerts = False   '警告メッセージ表示OFF
            For i = 1 To wnum
                Worksheets(2).Delete
            Next i
            Application.DisplayAlerts = True   '警告メッセージ表示ON
            Application.ScreenUpdating = True
        Else
            Application.ScreenUpdating = False
            If Worksheets().Count > 1 Then
                DeleteSheet (Tb_SheetNMInp)
                NewTableSheet (ActiveWorkbook.name)
            Else
                If HaveSheet(ActiveWorkbook.name, Tb_SheetNMInp) > 0 Then
                    Workbooks(ActiveWorkbook.name).Worksheets(Tb_SheetNMInp).name = Tb_SheetNMInp & "(1)"
                    NewTableSheet (ActiveWorkbook.name)
                    Workbooks(ActiveWorkbook.name).Worksheets(Tb_SheetNMInp).name = Tb_SheetNMInp & "(2)"
                    Workbooks(ActiveWorkbook.name).Worksheets(Tb_SheetNMInp & "(1)").name = Tb_SheetNMInp
                Else
                    NewTableSheet (ActiveWorkbook.name)
                End If
            Application.ScreenUpdating = True
            End If
        End If
    End If
End Sub

'==========================================================
'【プロシージャ名】NewTableSheet
'【概　要】テーブル定義書の追加
'【引　数】book名称
'【戻り値】なし
'==========================================================

Sub NewTableSheet(bname As String)
        Dim wkrange As String
        Workbooks("voyager.xla").Worksheets(Tb_SheetNm).Copy _
            before:=Workbooks(bname).Worksheets(1)
        Workbooks(bname).Worksheets(Tb_SheetNm).name = Tb_SheetNMInp
        Workbooks(bname).Worksheets(Tb_SheetNMInp).Activate
        'Cells(R_TblId, C_TblId).Value = Tb_SheetNMInp 'テーブルID
        Cells(R_TblId2, C_TblId2).Value = Tb_SheetNMInp 'テーブル名称
        Cells(R_COLNAME, C_COLNAME).Select 'カラムを物理名にセット
        kaktyoNV
End Sub


'==========================================================
'【プロシージャ名】DeleteSheet
'【概　要】シート削除
'【引　数】sheet名称
'【戻り値】なし
'==========================================================
Sub DeleteSheet(sname As String)
    Dim ret As Integer
    ret = HaveSheet(ActiveWorkbook.name, sname)
    If ret > 0 Then
        Beep
        answer = MsgBox(sname & " は既に作成されています " & Chr(13) & "上書きしてよろしいですか？", vbQuestion + vbOKCancel)
        If answer = vbOK Then
            Application.DisplayAlerts = False   '警告メッセージ表示OFF
            ActiveWorkbook.Worksheets(ret).Delete
            Application.DisplayAlerts = True   '警告メッセージ表示ON
        Else
            RunStop
        End If
    End If
End Sub


'==========================================================
'【プロシージャ名】RunStop
'【概　要】処理中断
'【引　数】なし
'【戻り値】なし
'==========================================================

Sub RunStop()
    Application.StatusBar = "処理を中断しました."
    Application.Cursor = xlNormal
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    End
    Stop
End Sub
