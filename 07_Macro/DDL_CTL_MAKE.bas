Attribute VB_Name = "DDL_CTL_MAKE"
'***PutDdl()  DDL出力
'***colname_seach(bname As String, sname As String) As Integer 物理名を探してカラムをセットする
'***spadd_r(i_st As String, i_len As Integer) As String 右側に指定された数のスペースをセットする

Sub PutDdlSheet()
    Dim abook As Workbook
    Dim sname As String
    
    'On Error GoTo err1
    Call Tb_Posget
    
    If Workbooks.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
        Exit Sub
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
        Exit Sub
    End If
        
    Set abook = ActiveWorkbook
    If abook.name <> "" Then
        If Cells(R_SheetId, C_SheetId).Value <> 2 Then
            Exit Sub
        End If
        Call MakeDdlSheet
    End If
End Sub

Sub PutDdlCpy()
    Dim abook As Workbook
    Dim sname As String
    
    On Error GoTo err1
    Set abook = ActiveWorkbook
    If abook.name <> "" Then
        If Cells(R_SheetId, C_SheetId).Value <> 2 Then
            Exit Sub
        End If
        Call MakeDdlCpy
    End If
    Exit Sub
err1:
    MsgBox ("シートを追加してください")
End Sub

' [DDL出力]->[ファイル]選択時
Sub PutDdlFil()
    Dim abook As Workbook
    Dim sname As String
    Dim tableListRow As Integer
    Dim tableListCol As Integer
    Dim tableName As String
    Dim dialog As FileDialog
    Dim rootFolderPath As String
    Dim retVal As Integer
    
    On Error GoTo err1
    
    Call Tb_Posget
    
    Set abook = ActiveWorkbook
    If abook.name <> "" Then
        
        '-----------------------------
        ' 出力先ルートフォルダ選択
        '-----------------------------
        Set dialog = Application.FileDialog(msoFileDialogFolderPicker)
        retVal = dialog.Show
        
        If retVal = -1 Then
            rootFolderPath = dialog.SelectedItems(1)
            ' 無ければ作る
            If Dir(rootFolderPath, vbDirectory) = "" Then
                MkDir (rootFolderPath)
            End If
            Set dialog = Nothing
        Else
            Set dialog = Nothing
            Exit Sub
        End If

        
        '-----------------------------
        ' 選択シートにより対象を変更
        '-----------------------------
        sname = abook.ActiveSheet.name
        
        If sname = "ＤＢ一覧" Then
            
            tableListRow = 5
            tableListCol = 2
            
            ' カラム参照ループ
            While abook.Worksheets(sname).Cells(tableListRow, tableListCol) <> ""
                ' テーブル名取得
                tableName = abook.Worksheets(sname).Cells(tableListRow, tableListCol)
                
                ' 同名のシートを選択
                abook.Worksheets(tableName).Select
                
'--- MOD Start 2019/06/20 SPC
'                If Cells(R_SheetId, C_SheetId).Value = 2 Then
                ' 「テーブルID」がビュー（"Z_"始まり）の場合終了
                If Cells(R_TblId2, C_TblId2).Value = "" _
                    Or (Len(Cells(R_TblId2, C_TblId2).Value) >= 2 And Left(Cells(R_TblId2, C_TblId2).Value, 2) = "Z_") Then
                Else
                    ' テーブル定義シートであればDDL作成実行
                    Call MakeDdlFil(rootFolderPath)
                End If
'--- MOD End 2019/06/20 SPC
                
                tableListRow = tableListRow + 1
            Wend
        Else
'--- MOD Start 2019/06/20 SPC
'            If Cells(R_SheetId, C_SheetId).Value <> 2 Then
            ' 「テーブルID」がビュー（"Z_"始まり）の場合終了
            If Cells(R_TblId2, C_TblId2).Value = "" _
                    Or (Len(Cells(R_TblId2, C_TblId2).Value) >= 2 And Left(Cells(R_TblId2, C_TblId2).Value, 2) = "Z_") Then
'--- MOD End 2019/06/20 SPC
                Exit Sub
            End If
            Call MakeDdlFil(rootFolderPath)
        End If
    End If
    
    MsgBox "完了！"
    
    Exit Sub
err1:
    MsgBox ("シートを追加してください。：" & tableName)
End Sub

Sub MakeDdlSheet()
    Dim bname As String
    Dim sname As String
    Dim ddl_sname As String
    Dim Dao As DataObject
    Dim wktext As String
    Dim MaxRow As Integer
    
    
    If ActiveWorkbook.ActiveSheet.Cells(R_DocId, C_DocId) <> 1 Then
        MsgBox ("テーブル定義書をアクティブにしてください")
        Exit Sub
    End If

    MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
    If MaxRow = 0 Then
        MsgBox ("項目IDに空欄があります")
        Exit Sub
    End If
    
    
    If Checkspace(R_COLNAME, C_kata, MaxRow) = 0 Then
        MsgBox ("型に空欄があります")
        Exit Sub
    End If
        
    '---MOD START 2010/07/27 OU
    'If Checkspace(R_COLNAME, C_keta, MaxRow) = 0 Then
    If Checkspaceketa(R_COLNAME, C_keta, MaxRow) = 0 Then
    '---MOD END
        MsgBox ("桁に空欄があります")
        Exit Sub
    End If
    
    Call seisho
    bname = ActiveWorkbook.name
    sname = ActiveWorkbook.ActiveSheet.name
'--- Mod Start 2012/03/05 TFC
    wktext = CreateDdl(bname, sname, DDL_KIND_ALL)
'--- Mod End 2012/03/05 TFC
    ddl_sname = "Create_" & sname
    DeleteSheet (ddl_sname)
    Set Dao = New DataObject
    Worksheets().Add
    Workbooks(ActiveWorkbook.name).ActiveSheet.name = ddl_sname
    ActiveWindow.DisplayGridlines = False
    Cells.Select
    With Selection.Font
        .name = "ＭＳ ゴシック"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
    Dao.SetText wktext
    Dao.PutInClipboard
    Cells(1.1).PasteSpecial
    Cells(1, 1).Select
End Sub

Sub MakeDdlCpy()
    Dim bname As String
    Dim sname As String
    Dim Dao As DataObject
    Dim wktext As String
    Dim MaxRow As Integer
    
    Call Tb_Posget
    
    If ActiveWorkbook.ActiveSheet.Cells(R_DocId, C_DocId) <> 1 Then
        MsgBox ("テーブル定義書をアクティブにしてください")
        Exit Sub
    End If

    MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
    If MaxRow = 0 Then
        MsgBox ("項目IDに空欄があります")
        Exit Sub
    End If
    
    
    If Checkspace(R_COLNAME, C_kata, MaxRow) = 0 Then
        MsgBox ("型に空欄があります")
        Exit Sub
    End If
    
    '---MOD START 2010/07/27 OU
    'If Checkspace(R_COLNAME, C_keta, MaxRow) = 0 Then
    If Checkspaceketa(R_COLNAME, C_keta, MaxRow) = 0 Then
    '---MOD END
        MsgBox ("桁に空欄があります")
        Exit Sub
    End If
    
    Call seisho
    bname = ActiveWorkbook.name
    sname = ActiveWorkbook.ActiveSheet.name
'--- Mod Start 2012/03/05 TFC
    wktext = CreateDdl(bname, sname, DDL_KIND_ALL)
'--- Mod End 2012/03/05 TFC
    Set Dao = New DataObject
    Dao.SetText wktext
    Dao.PutInClipboard
End Sub

Sub MakeDdlFil(rootFolderPath As String)
    Dim bname As String
    Dim sname As String
    Dim Dao As DataObject
    Dim wktext As String
    Dim strFilePath As String
    Dim intFileNo As Integer
    Dim MaxRow As Integer
    Dim folderPath As String
    Dim strTableId As String
    
    Call Tb_Posget
    
'--- DEL Start 2019/06/20 SPC
'    If ActiveWorkbook.ActiveSheet.Cells(R_DocId, C_DocId) <> 1 Then
'        MsgBox ("テーブル定義書をアクティブにしてください")
'        Exit Sub
'    End If
'--- DEL Start 2019/06/20 SPC

    MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
    If MaxRow = 0 Then
        MsgBox ("項目IDに空欄があります")
        Exit Sub
    End If
    
    
    If Checkspace(R_COLNAME, C_kata, MaxRow) = 0 Then
        MsgBox ("型に空欄があります")
        Exit Sub
    End If
        
    '---MOD START 2010/07/27 OU
    'If Checkspace(R_COLNAME, C_keta, MaxRow) = 0 Then
    If Checkspaceketa(R_COLNAME, C_keta, MaxRow) = 0 Then
    '---MOD END
        MsgBox ("桁に空欄があります")
        Exit Sub
    End If
    
'--- DEL Start 2019/06/20 SPC
'    Call seisho
'--- DEL Start 2019/06/20 SPC
    
    ' テーブルID保持
    strTableId = Cells(R_TblId2, C_TblId2).Value
    
    bname = ActiveWorkbook.name
    sname = ActiveWorkbook.ActiveSheet.name
    
    ' ----------------------------------
    ' テーブル定義出力
    ' ----------------------------------
    wktext = CreateDdl(bname, sname, DDL_KIND_TABLE)
    
    '保存先の取得
    folderPath = rootFolderPath & "\table"
    
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir (folderPath)
    End If
    
    strFilePath = folderPath & "\" & strTableId & ".sql"
    
    'ファイル出力
'--- MOD Start 2019/07/08 SPC
'    intFileNo = FreeFile
'    Open strFilePath For Output As #intFileNo
'    Print #intFileNo, wktext
'    Close #intFileNo
    Call outputUtf8File(strFilePath, wktext)
'--- MOD End 2019/07/08 SPC
    
    ' ----------------------------------
    ' PK、インデックス定義出力
    ' ----------------------------------
    wktext = CreateDdl(bname, sname, DDL_KIND_INDEX)
    
    If wktext <> "" Then
        '保存先の取得
        folderPath = rootFolderPath & "\index"
        
        If Dir(folderPath, vbDirectory) = "" Then
            MkDir (folderPath)
        End If
        
        strFilePath = folderPath & "\" & strTableId & ".sql"
        
        'ファイル出力
'--- MOD Start 2019/07/08 SPC
'        intFileNo = FreeFile
'        Open strFilePath For Output As #intFileNo
'        Print #intFileNo, wktext
'        Close #intFileNo
        Call outputUtf8File(strFilePath, wktext)
'--- MOD Start 2019/07/08 SPC
        
    End If
    
End Sub

'==========================================================
'【プロシージャ名】PutDdl
'【概　要】DDL文出力
'【引　数】なし
'【戻り値】なし
'==========================================================

Sub PutDdl()
    Dim wkmsg As String
    Dim wknum As Integer
    Dim ans As Integer
    Dim tbl_sname As String
    Dim sname As String
    Dim ddl_sname As String
    Dim tableName As String
    Dim wk_len As Integer
    Dim x As Integer
    Dim y As Integer
    Dim curr_l As Integer
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim TABLENAMEJ As String
    Dim pkey(32768) As Integer
    Dim pkey_seq(32768) As Integer
    Dim pkey_pos(32768) As Integer
    Dim wktablesp As String
    Dim tblsheet As Worksheet
    Dim wktype As String
    
    Tb_Posget
    seisho
    

    If Cells(R_COLNAME, C_COLNAME) = "" Then
        MsgBox ("項目IDが入力されていません")
        Exit Sub
    End If

    tbl_sname = Workbooks(ActiveWorkbook.name).ActiveSheet.Cells(R_TblId, C_TblId)
    ddl_sname = "CREATE文_" & tbl_sname
    DeleteSheet (ddl_sname)
    Worksheets().Add
    Workbooks(ActiveWorkbook.name).ActiveSheet.name = ddl_sname
    sname = Workbooks(ActiveWorkbook.name).ActiveSheet.name
    Workbooks(ActiveWorkbook.name).Worksheets(sname).name = ddl_sname
    ActiveWindow.DisplayGridlines = False

    
    Application.ScreenUpdating = False
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Select
    Cells.Select
    With Selection.Font
        .name = "ＭＳ ゴシック"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
        
    'コメント行1行目
    x = 1
    y = 1
    Cells(y, x).Value = "/**********************************************************/"
    
    'コメント行　TABLENAME
    wkmsg = DdlCom_TbId(tbl_sname)
    
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    y = y + 1
    Cells(y, x).Value = wkmsg
    
    'コメント行　TABLENAME
    wkmsg = DdlCom_TbNm(tbl_sname)
    
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    y = y + 1
    Cells(y, x).Value = wkmsg
    
    wkmsg = "/*     " & "作成日:" & Format(Date, "yyyy/mm/dd")
    
    wk_len = LenB(StrConv(wkmsg, vbFromUnicode))
    wknum = 57 - wk_len
    wkmsg = spadd_r(wkmsg, wknum)
    wkmsg = wkmsg & " */"
    
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    y = y + 1
    Cells(y, x).Value = wkmsg
    
'--- DEL Start 2015/02/27 TFC
'    y = y + 1
'    Cells(y, x).Value = "/**********************************************************/"
'    y = y + 1
'    Cells(y, x).Value = "/* エラーハンドリング */"
'    y = y + 1
'    Cells(y, x).Value = "WHENEVER OSERROR  EXIT OSCODE      ROLLBACK"
'    y = y + 1
'    Cells(y, x).Value = "WHENEVER SQLERROR EXIT SQL.SQLCODE ROLLBACK"
'--- DEL End

    y = y + 1
    Cells(y, x).Value = "/* CREATE 文 */"
    
    tableName = Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Cells(R_TblId2, C_TblId2).Value
    wkschima = Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Cells(R_Schima, C_Schima).Value
    If wkschima = "" Then
        wkmsg = "CREATE TABLE " & tableName & "("
    Else
        wkmsg = "CREATE TABLE " & wkschima & "." & tableName & "("
    End If
        
    y = y + 1
    Cells(y, x).Value = wkmsg
    curr_l = R_COLNAME '物理名1行目
    
    '項目名を物理名カラムから取得してDDL文を作成する
    While Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Cells(curr_l, C_COLNAME) <> ""
        Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Activate
        wkmsg = "       " & Cells(curr_l, C_COLNAME)
        wkmsg = wkmsg & " " & Cells(curr_l, C_kata)
        wkmsg = wkmsg & "(" & Cells(curr_l, C_keta)
        If Cells(curr_l, C_shou) <> "" Then
           If Cells(curr_l, C_kata) = "NUMBER" Then
                wkmsg = wkmsg & "," & Cells(curr_l, C_shou)
            End If
        End If
        wkmsg = wkmsg & ")"
        If Cells(curr_l, C_nnul).Value = "○" Then
            wkmsg = wkmsg & " NOT NULL"
        End If
        If Cells(curr_l, C_uniq).Value = "○" Then
            wkmsg = wkmsg & " UNIQUE"
        End If
        
        wktype = Cells(curr_l, C_kata).Value
        If Cells(curr_l, C_def).Value <> "" Then
            wkmsg = wkmsg & " DEFAULT "
            If (wktype = "CHAR") Or (wktype = "VARCHAR2") Then
                wkmsg = wkmsg & " '" & Cells(curr_l, C_def) & "'"
            Else
                wkmsg = wkmsg & Cells(curr_l, C_def)
            End If
        End If
        
        If Cells(curr_l + 1, C_COLNAME) <> "" Then
            wkmsg = wkmsg & ","
        End If
        
        
        Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
        y = y + 1
        Cells(y, x).Value = wkmsg
        curr_l = curr_l + 1
    Wend
    
    '最後の”）”を追加する
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    wkmsg = "       )"
    wktablesp = Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Cells(R_TblSp, C_TblSp).Value
    If wktablesp <> "" Then
        wkmsg = wkmsg & " TABLESPACE "
        wkmsg = wkmsg & wktablesp
        wkmsg = wkmsg & ";"
    Else
        wkmsg = wkmsg & ";"
    End If
    y = y + 1
    Cells(y, x).Value = wkmsg

        
    '主キー制約を追加する
    Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Activate
    j = 0
    For i = R_COLNAME To curr_l - 1
        If Cells(i, C_primary) <> "" Then
            pkey(j) = CInt(Cells(i, C_primary))
            pkey_pos(j) = i
            pkey_seq(j) = j + 1
            j = j + 1
        End If
    Next i
    
    wkpkey = ""
    If j > 0 Then
        For i = 0 To j - 1
            For k = 0 To j - 1
                If pkey_seq(i) = pkey(k) Then
                    If i < j - 1 And j > 0 Then
                        wkpkey = wkpkey & Cells(pkey_pos(k), C_COLNAME) & ","
                    Else
                        wkpkey = wkpkey & Cells(pkey_pos(k), C_COLNAME)
                    End If
                End If
            Next k
        Next i
    End If
    
    If wkpkey <> "" Then
        Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
        y = y + 1
        Cells(y, x).Value = "/* PRIMARY KEY */"
        y = y + 1
        Cells(y, x).Value = "ALTER TABLE " & tableName
        wkmsg = " ADD CONSTRAINT " & "PK_" & tableName & " PRIMARY KEY(" & wkpkey & ")"
        
        'TABLE表領域が指定されていれば追加する
        Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Activate
        If Cells(pkey_pos(0), C_IdxSp2).Value <> "" Then
        
            Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
            y = y + 1
            Cells(y, x).Value = wkmsg
            
            Workbooks(ActiveWorkbook.name).Worksheets(tbl_sname).Activate
            wktablesp = " USING INDEX TABLESPACE " & Cells(pkey_pos(0), C_IdxSp2)
            
            Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
            y = y + 1
            Cells(y, x).Value = wktablesp & ";"
        Else
            Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
            wkmsg = wkmsg & ";"
            y = y + 1
            Cells(y, x).Value = wkmsg
        End If
            
    End If
    
    Workbooks(ActiveWorkbook.name).Worksheets(ddl_sname).Activate
    Application.ScreenUpdating = True
    Range("A1").Select
End Sub

'==========================================================
'【関数名】DdlCom_TbId
'【概　要】テーブル名を編集
'【引　数】シート名
'【戻り値】コメント文
'==========================================================

Function DdlCom_TbId(sname As String) As String
    Dim tableName As String
    Dim wkmsg As String
    Dim wknum As Integer
    
    tableName = Workbooks(ActiveWorkbook.name).Worksheets(sname).Cells(R_TblId2, C_TblId2).Value
    wkmsg = "/*     TABLE NAME: " & tableName
    wk_len = Len(wkmsg)
    wknum = 57 - wk_len
    wkmsg = spadd_r(wkmsg, wknum)
    wkmsg = wkmsg & " */"
    
    DdlCom_TbId = wkmsg
    
End Function

'==========================================================
'【関数名】DdlCom_TbNm
'【概　要】テーブル名を編集
'【引　数】シート名
'【戻り値】コメント文
'==========================================================
Function DdlCom_TbNm(sname As String) As String
    Dim wkmsg As String
    Dim TABLENAMEJ As String
    Dim wk_len As Integer
    Dim wknum As Integer
    
    TABLENAMEJ = Workbooks(ActiveWorkbook.name).Worksheets(sname).Cells(R_TblNm, C_TblNm).Value
    wkmsg = "/*     " & TABLENAMEJ
    wk_len = LenB(StrConv(wkmsg, vbFromUnicode))
    
    wknum = 57 - wk_len
    wkmsg = spadd_r(wkmsg, wknum)
    wkmsg = wkmsg & " */"
    
    DdlCom_TbNm = wkmsg
    
End Function




'==========================================================
'【関数名】spadd_r
'【概　要】Stringの右側に半角Spaceを詰める
'【引　数】詰める前のString,詰める文字数
'【戻り値】詰めた後のString
'==========================================================

Function spadd_r(i_st As String, i_len As Integer) As String
    Dim i As Integer
    Dim wk_st As String
    wk_st = i_st
    For i = 1 To i_len
        wk_st = wk_st & " "
    Next i
    spadd_r = wk_st
End Function


'--- Add Start S.Iwanaga 2010/04/08
'==========================================================
'【関数名】menuExAreaView
'【概　要】拡張カラム表示/非表示処理
'【引　数】なし
'【戻り値】なし
'==========================================================
Public Sub menuExAreaView()

    Dim strSheet    As String
    
    If Workbooks.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name
        
        'シートタイプがテーブル項目シートの場合のみ処理
        If isTblDefSheet(strSheet) Then
            '拡張カラムの表示状態取得
            If chkExAreaState(strSheet) = 0 Then
                '非表示
                Call kaktyoV
            Else
                '表示
                Call kaktyoNV
            End If
        End If
    End If
        
End Sub

'==========================================================
'【関数名】menuConvLogicName
'【概　要】メニュー物理名変換処理
'【引　数】なし
'【戻り値】なし
'==========================================================
Public Sub menuConvLogicName()

    Dim strSheet    As String
    
    If Workbooks.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name
    
        '---Mod Start OU 2010/07/27
        'If LtoP(strSheet, 8, 20) = -1 Then
        If LtoP(strSheet, R_COLNAME, 70) = -1 Then
        '---Mod End
            '変換処理エラー
        End If
    End If
    
End Sub
'--- Add End

'--- Add Start S.Iwanaga 2010/04/16
'==========================================================
'【関数名】menuCtlToFile
'【概　要】CTLをファイルに出力
'【引　数】なし
'【戻り値】なし
'==========================================================
Public Sub menuCtlToFile()

    Dim intRtn      As Integer
    Dim intFileNo   As Integer
    Dim strSheet    As String
    Dim strFilePath As String
    Dim strCtlData  As String
    Dim MaxRow As Integer
    
    strCtlData = ""
    strFilePath = ""
    
    If Workbooks.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name

        'テーブル項目シートかチェック
        If Not isTblDefSheet(strSheet) Then
            MsgBox "テーブル項目シートをアクティブにして下さい。", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        
        MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
        If MaxRow = 0 Then
            MsgBox ("未入力の項目IDセルが存在します")
            Exit Sub
        End If
        
        'Add Start 2010/07/29 OU
        Dim strDataType As String
        strDataType = Trim(Cells(R_DataTyp, C_DataTyp).Value)
        If strDataType <> "CSV" Then
        'Add End
            If Checkspace(R_COLNAME, C_FFilePosition, MaxRow) = 0 Then
                MsgBox ("未入力のフラットファイル位置セルが存在します")
                kaktyoV
                Exit Sub
            End If
            
            If Checkspace(R_COLNAME, C_FFileLength, MaxRow) = 0 Then
                MsgBox ("未入力のフラットファイル桁セルが存在します")
                kaktyoV
                Exit Sub
            End If
        End If
        
        '物理名未入力チェック
        intRtn = chkBlankPName(strSheet)
        If intRtn = 1 Then
            MsgBox "未入力の項目IDセルが存在します。", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        
        '保存先の取得
        strFilePath = Application.GetSaveAsFilename(strSheet & ".ctl", "制御ファイル, *.ctl")
        
        '制御ファイル作成
        strCtlData = createCtl(strSheet)

        If Len(strCtlData) > 0 Then
            'ファイル出力
            intFileNo = FreeFile
            Open strFilePath For Output As #intFileNo
            Print #intFileNo, strCtlData
            Close #intFileNo
        End If
    End If

End Sub

'==========================================================
'【関数名】menuCtlToSheet
'【概　要】CTLをシートに出力
'【引　数】なし
'【戻り値】なし
'==========================================================
Public Sub menuCtlToSheet()

    Dim strSheet        As String
    Dim strCtlData      As String
    Dim strTableName    As String
    Dim strCtlShtName   As String
    Dim dd              As New DataObject
    Dim ws              As Worksheet
    Dim MaxRow As Integer
    
    strCtlData = ""
    
    If Workbooks.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("テーブル定義書を開いてください")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name

        'テーブル項目シートかチェック
        If Not isTblDefSheet(strSheet) Then
            MsgBox "テーブル項目シートを選択して下さい。", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        
        MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
        If MaxRow = 0 Then
            MsgBox ("未入力の項目IDセルが存在します")
            Exit Sub
        End If
        
        'Add Start 2010/07/29 OU
        Dim strDataType As String
        strDataType = Trim(Cells(R_DataTyp, C_DataTyp).Value)
        If strDataType <> "CSV" Then
        'Add End
            If Checkspace(R_COLNAME, C_FFilePosition, MaxRow) = 0 Then
                MsgBox ("未入力のフラットファイル位置セルが存在します")
                kaktyoV
                Exit Sub
            End If
            
            If Checkspace(R_COLNAME, C_FFileLength, MaxRow) = 0 Then
                MsgBox ("未入力のフラットファイル桁セルが存在します")
                kaktyoV
                Exit Sub
            End If
        End If
        
       '物理名未入力チェック
        intRtn = chkBlankPName(strSheet)
        If intRtn = 1 Then
            MsgBox "未入力の項目IDセルが存在します。", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
                
        '制御ファイル作成
        strCtlData = createCtl(strSheet)

        dd.SetText strCtlData
        dd.PutInClipboard

        strTableName = Trim(ActiveWorkbook.Sheets(strSheet).Cells(R_TblId, C_TblId).Value)
        strCtlShtName = "CTL_" & strTableName
        Call DeleteSheet(strCtlShtName)
        Set ws = ActiveWorkbook.Worksheets.Add()
        ws.name = strCtlShtName
        ws.Cells(1, 1).PasteSpecial
        'Add Start 2010/07/30 OU
        ActiveWindow.DisplayGridlines = False
        Cells(1, 1).Select
        'Add End
    End If

End Sub

'==========================================================
'【関数名】menuCtlToClipboard
'【概　要】CTLをクリップボードにコピー
'【引　数】なし
'【戻り値】なし
'==========================================================
Public Sub menuCtlToClipboard()

    Dim strSheet    As String
    Dim strCtlData  As String
    Dim dd          As New DataObject
    Dim MaxRow As Integer
    
    strCtlData = ""
    
    If Workbooks.Count < 1 Then
        MsgBox ("有効なテーブル定義書を開いてください")
    ElseIf ActiveWorkbook.Sheets.Count < 1 Then
        MsgBox ("有効なテーブル定義書を開いてください")
    Else
        Call Tb_Posget
        
        strSheet = ActiveWorkbook.ActiveSheet.name

        'テーブル項目シートかチェック
        If Not isTblDefSheet(strSheet) Then
            MsgBox "テーブル項目シートを選択して下さい。", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
        
        
        MaxRow = Checkspace(R_COLNAME, C_COLNAME, 0)
        If MaxRow = 0 Then
            MsgBox ("未入力の項目IDセルが存在します")
            Exit Sub
        End If
        
        'Add Start 2010/07/29 OU
        Dim strDataType As String
        strDataType = Trim(Cells(R_DataTyp, C_DataTyp).Value)
        If strDataType <> "CSV" Then
        'Add End
            If Checkspace(R_COLNAME, C_FFilePosition, MaxRow) = 0 Then
                MsgBox ("未入力のフラットファイル位置セルが存在します")
                kaktyoV
                Exit Sub
            End If
            
            If Checkspace(R_COLNAME, C_FFileLength, MaxRow) = 0 Then
                MsgBox ("未入力のフラットファイル桁セルが存在します")
                kaktyoV
                Exit Sub
            End If
        
        End If
        
        
        '物理名未入力チェック
        intRtn = chkBlankPName(strSheet)
        If intRtn = 1 Then
            MsgBox "未入力の項目IDセルが存在します。", vbExclamation + vbOKOnly, "Error"
            Exit Sub
        End If
                
        '制御ファイル作成
        strCtlData = createCtl(strSheet)

        dd.SetText strCtlData
        dd.PutInClipboard

    End If

End Sub
'--- Add End

'--- ADD Start 2019/07/08 SPC
Sub outputUtf8File(filePath As String, wtext As String)

    Dim stream As ADODB.stream
    Set stream = New ADODB.stream
    Dim byteData() As Byte
    
    stream.Type = adTypeText
    stream.Charset = "UTF-8"
    stream.LineSeparator = adCRLF
    stream.Open

    stream.WriteText wtext, adWriteLine
    
    '----------------
    ' BOM削除
    '----------------
    ' BOMコードを飛ばしてデータを取得する。
    stream.Position = 0
    stream.Type = adTypeBinary
    stream.Position = 3
    byteData = stream.Read
    
    ' 取得したデータを先頭から書き出し直す
    stream.Position = 0
    stream.Write byteData
    stream.SetEOS
    
    ' ファイル保存
    stream.SaveToFile filePath, adSaveCreateOverWrite
    stream.Close
    
End Sub
'--- ADD End 2019/07/08 SPC


