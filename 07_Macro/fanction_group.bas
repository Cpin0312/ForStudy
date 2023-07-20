Attribute VB_Name = "fanction_group"
'Function
'Checkspace(i_R As Integer, i_C As Integer, chk As Integer) As Integer
'HaveSheet(bname As String, sname As String) As Integer
'HaveSheet2(bname As String, sname As Integer) As Integer
'colposget(name As String) As Integer
'CreateDdl(bname As String, sname As String) As String
'Checkphy(bname As String, sname As String) As Integer
'HaveSheet(bname As String, sname As String) As Integer

'==========================================================
'【関数名】CheckSpace
'【概　要】カラムの空欄をチェック
'【引　数】i_R     :Row
'【引　数】i_C     :Column
'【戻り値】0:空欄あり n:最終行
'==========================================================

Function Checkspace(i_R As Integer, i_C As Integer, chk As Integer) As Integer
    Dim MaxRow As Integer
    Dim DmaxRow As Integer
    If Cells(i_R, i_C) = "" Then
        Checkspace = 0
        Exit Function
    End If
    If chk = 0 Then
        MaxRow = Cells(Rows.Count, i_C).End(xlUp).Row
    Else
        MaxRow = chk
    End If
    
    If Cells(i_R + 1, i_C).Value = "" Then
        DmaxRow = i_R
    Else
        DmaxRow = Range(Cells(i_R, i_C), Cells(i_R, i_C)).End(xlDown).Row
    End If
        
    If MaxRow <> DmaxRow Then
        Checkspace = 0
    Else
        Checkspace = MaxRow
    End If
    
End Function

'---ADD START 2010/07/27 OU
'==========================================================
'【関数名】Checkspaceketa
'【概　要】桁数カラムの空欄をチェック
'【引　数】i_R     :Row
'【引　数】i_C     :Column
'【戻り値】0:空欄あり n:最終行
'==========================================================
Function Checkspaceketa(i_R As Integer, i_C As Integer, chk As Integer) As Integer
    Dim MaxRow As Integer

    If chk = 0 Then
        MaxRow = Cells(Rows.Count, i_C).End(xlUp).Row
    Else
        MaxRow = chk
    End If
    
    Dim tmp As String
    Dim i_rl As Integer
    For i_rl = i_R To MaxRow
        If Cells(i_rl, C_kata).Value = "DATE" Or Cells(i_rl, C_kata).Value = "TIMESTAMP" Or Cells(i_rl, C_kata).Value = "BLOB" _
            Or Cells(i_rl, C_kata).Value = "INTEGER" Or Cells(i_rl, C_kata).Value = "BYTEA" Then
            Cells(i_rl, i_C).Value = ""
            Cells(i_rl, C_shou).Value = ""
        ElseIf Cells(i_rl, i_C).Value = "" Then
            Checkspaceketa = 0
            Exit Function
        End If
    Next
    
    Checkspaceketa = MaxRow
End Function
'--- ADD END

'==========================================================
'【関数名】HaveSheet
'【概　要】同名のシートがbook上にあるか調べる
'【引　数】bname     :book名
'【引　数】sname     :sheet名
'【戻り値】True=sheet Index値
'==========================================================
Function HaveSheet(bname As String, sname As String) As Integer
    Dim wnum As Integer        '登録ワークシート数
    Dim i As Integer           'カウンタ
    Dim ret As Integer
    Dim bookname As String
    
    ret = 0
    wnum = Workbooks(bname).Worksheets.Count     '登録ワークシート数を得る
    For i = 1 To wnum
        If UCase(Workbooks(bname).Worksheets(i).name) = UCase(sname) Then
            ret = i
            Exit For
        End If
    Next i
    HaveSheet = ret
End Function

'==========================================================
'【関数名】HaveSheet2
'【概　要】同じシートidがbook上にあるか調べる
'【引　数】bname     :book名
'【引　数】sname     :sheetid
'【戻り値】True=sheet Index値
'==========================================================
Function HaveSheet2(bname As String, sname As Integer) As Integer
    Dim wnum As Integer        '登録ワークシート数
    Dim i As Integer           'カウンタ
    Dim ret As Integer
    Dim bookname As String
    
    ret = 0
    wnum = Workbooks(bname).Worksheets.Count     '登録ワークシート数を得る
    For i = 1 To wnum
        If CInt(Workbooks(bname).Worksheets(i).Cells(R_SheetId, C_SheetId)) = sname Then
            ret = i
            Exit For
        End If
    Next i
    HaveSheet2 = ret
End Function


'==========================================================
'【関数名】colposget
'【概　要】テーブル項目シートのカラム位置を取得
'【引　数】項目名
'【戻り値】カラム位置
'==========================================================

Function colposget(name As String) As Integer
    Dim i As Integer
    For i = C_COLNAME To C_KeiEnd
        If ThisWorkbook.Worksheets("テーブル項目").Cells(R_COLNAME - 2, i).Value = name Then
            colposget = i
            Exit For
        End If
    Next i
    
End Function


'==========================================================
'【関数名】CreateDdl
'【概　要】クリエイト文を作成する
'【引　数】ブック名称、シート名称、
'          種別（DDL_KIND_ALL：テーブル＆インデックス
'                DDL_KIND_TABLE：テーブル定義のみ
'                DDL_KIND_INDEX：インデックス定義のみ）
'【戻り値】クリエイト文
'==========================================================
Function CreateDdl(bname As String, sname As String, ddlKind As String) As String
    Dim ddl_sname As String
    Dim MaxRow As Integer
    Dim Dao As DataObject
    Dim wktext As String
    Dim wkintext As String
    Dim wkset As String
    Dim curr_l
    Dim j As Integer
    Dim pkey(32768) As Integer
    Dim pkey_seq(32768) As Integer
    Dim pkey_pos(32768) As Integer
    Dim wkpkey As String
    '--- Add Start 2012/02/27 TFC
    Dim idx() As Integer
    Dim idx_seq() As Integer
    Dim idx_pos() As Integer
    '--- Add End 2012/02/27 TFC
    Dim TableId As String
    Dim wktype As String
    '--- Add Start 2010/07/30 OU
    Dim strComment As String
    Dim strCommentTable As String
    strComment = "/* COMMENT */" & vbCrLf
    Cells(R_TblSp, C_TblSp) = UCase(Cells(R_TblSp, C_TblSp))
    Cells(R_IdxSp, C_IdxSp) = UCase(Cells(R_IdxSp, C_IdxSp))
    '--- Add End
    '--- Add End 2012/02/27 TFC
    Dim pkeyName As String
    Dim strindex As String
    Dim intIndexC As Integer
    Dim intIndexR As Integer
    Dim strIndexSpace As String
    Dim indexNamePrefix As String
    Dim indexName As String
    Dim fileHeader As String
    '--- Add End 2012/02/27 TFC
    '--- ADD Start 2019/07/19 SPC
    Dim wkPartitionKind As String
    Dim wkPartitionKoumoku As String
    '--- ADD End 2019/07/19 SPC
    
    fileHeader = "/**********************************************************/" + vbCrLf
    '--- Mod Start 2010/09/21
    'TableId = ActiveWorkbook.Worksheets(sname).Cells(R_TblId, C_TblId).Value
    TableId = Cells(R_TblId2, C_TblId2).Value
    '--- Mod End
    wkset = "/*     TABLE NAME: " & TableId
    wkset = spadd_r(wkset, 57 - Len(wkset))
    fileHeader = fileHeader & wkset & " */" + vbCrLf
    
    wkintext = Workbooks(bname).Worksheets(sname).Cells(R_TblNm, C_TblNm).Value
    wkset = "/*     テーブル名：" & wkintext
        
    '--- Add Start OU 2019/10/15
    strComment = strComment & "COMMENT ON TABLE " & Cells(R_TblId2, C_TblId2).Value & " IS '" & wkintext & "';" & vbCrLf
    
    wkset = spadd_r(wkset, 57 - LenB(StrConv(wkset, vbFromUnicode)))
    fileHeader = fileHeader & wkset & " */" + vbCrLf
    
'--- DEL Start 2019/06/20 SPC
'    wkset = "/*     " & "作成日:" & Format(Date, "yyyy/mm/dd")
'
'    wkset = spadd_r(wkset, 57 - LenB(StrConv(wkset, vbFromUnicode)))
'    fileHeader = fileHeader & wkset & " */" + vbCrLf
'--- DEL End 2019/06/20 SPC
    
    fileHeader = fileHeader & "/**********************************************************/" + vbCrLf
'--- DEL Start 2015/02/27 TFC
'    wktext = wktext & "/* エラーハンドリング */" + vbCrLf
'    wktext = wktext & "WHENEVER OSERROR  EXIT OSCODE      ROLLBACK" + vbCrLf
'    wktext = wktext & "WHENEVER SQLERROR EXIT SQL.SQLCODE ROLLBACK" + vbCrLf
'--- DEL End 2015/02/27 TFC
    
    ' テーブルID取得
    wkintext = Cells(R_TblId2, C_TblId2).Value
    If Cells(R_Schima, C_Schima).Value <> "" Then
        wkintext = Cells(R_Schima, C_Schima).Value + "." + wkintext
    End If
    strCommentTable = wkintext
    
    
    If ddlKind = DDL_KIND_ALL Or ddlKind = DDL_KIND_TABLE Then
        
        wktext = wktext & "/* CREATE 文 */" + vbCrLf
        
        wktext = wktext + "CREATE TABLE " + wkintext + "(" + vbCrLf
        
        curr_l = R_COLNAME '物理名1行目
        
        While Cells(curr_l, C_COLNAME) <> ""
            wktext = wktext + "       " & Cells(curr_l, C_COLNAME)
            '--- Add Start 2010/07/30
            Cells(curr_l, C_kata) = UCase(Cells(curr_l, C_kata))
            wktype = Cells(curr_l, C_kata)
            If wktype = "INTEGER" Then
                wktype = "INT"
            End If
            '--- Add End
            wktext = wktext & " " & wktype
            '--- Add START 2010/07/27 OU
            If wktype <> "DATE" And wktype <> "TIMESTAMP" And wktype <> "BLOB" And wktype <> "INT" And wktype <> "BYTEA" Then
                wktext = wktext & "(" & Cells(curr_l, C_keta)
            End If
            '--- Add END
            If (Cells(curr_l, C_kata) = "NUMBER" Or Cells(curr_l, C_kata) = "NUMERIC") And Cells(curr_l, C_shou) <> "" Then
                wktext = wktext & "," & Cells(curr_l, C_shou)
            End If
            '---ADD START 2010/07/27 OU
            If wktype <> "DATE" And wktype <> "TIMESTAMP" And wktype <> "BLOB" And wktype <> "INT" And wktype <> "BYTEA" Then
                wktext = wktext & ")"
            End If
            '---ADD END
            
            '--- Add Start OU 2010/07/26
            'デフォルト値の設定
            If Cells(curr_l, C_def) <> "" Then
                wktext = wktext & " DEFAULT "
                If (wktype = "CHAR") Or (wktype = "VARCHAR2") Or (wktype = "VARCHAR") Then
                    wktext = wktext & "'" & Cells(curr_l, C_def) & "'"
                Else
                    wktext = wktext & Cells(curr_l, C_def)
                End If
            End If
            
            
            strComment = strComment & "COMMENT ON COLUMN " & strCommentTable & "." & Cells(curr_l, C_COLNAME) & " IS '" & Cells(curr_l, C_ITEMNAME) & "';" & vbCrLf
            '---Add End
            If Cells(curr_l, C_uniq) = "○" Then
                wktext = wktext + " UNIQUE"
            End If
            If Cells(curr_l, C_nnul) = "○" Then
                wktext = wktext + " NOT NULL"
            End If
            If Cells(curr_l + 1, C_COLNAME) <> "" Then
                wktext = wktext & ","
            End If
            wktext = wktext + vbCrLf
            curr_l = curr_l + 1
        Wend
        wktext = wktext + "       )"
        
        '--- ADD Start 2019/07/19 SPC
        wkPartitionKind = Cells(R_PartitionKind, C_PartitionKind).Value
        wkPartitionKoumoku = Cells(R_PartitionKoumoku, C_PartitionKoumoku).Value
        ' パーティション定義が存在する場合
        If wkPartitionKind <> "" And wkPartitionKoumoku <> "" Then
            wktext = wktext & vbCrLf
            wktext = wktext & "       "
            wktext = wktext & "PARTITION BY " & wkPartitionKind & " (" & wkPartitionKoumoku & ")" & vbCrLf
            wktext = wktext & "       "
        End If
        '--- ADD End 2019/07/19 SPC
        
        wkintext = Cells(R_TblSp, C_TblSp).Value
        If wkintext <> "" Then
            wktext = wktext & "TABLESPACE "
            wktext = wktext & wkintext
            wktext = wktext & ";" + vbCrLf
        Else
            wktext = wktext & ";" + vbCrLf
        End If
        
        '--- Add Start 2010/07/30 OU
        wktext = wktext & vbCrLf & strComment
        '--- Add End
    End If
    
    
    If ddlKind = DDL_KIND_ALL Or ddlKind = DDL_KIND_INDEX Then
            
        '=====================================
        ' プライマリキー作成
        '=====================================
        j = 0
        intIndexR = R_COLNAME '物理名1行目
        
        ' カラム参照ループ
        While Cells(intIndexR, C_COLNAME) <> ""
            ' 主キー列に設定がある場合
            If Cells(intIndexR, C_primary) <> "" Then
                pkey(j) = CInt(Cells(intIndexR, C_primary)) ' 主キー設定値（数値変換後）格納
                pkey_pos(j) = intIndexR                     ' 行番号格納
                pkey_seq(j) = j + 1                 ' シーケンス番号格納
                j = j + 1
            End If
            
            intIndexR = intIndexR + 1
        Wend
        
        wkpkey = ""
        ' 主キー設定データを保持した場合
        If j > 0 Then
            ' シーケンス番号参照ループ
            For i = 0 To j - 1
                ' 主キー設定値参照ループ
                For k = 0 To j - 1
                    If pkey_seq(i) = pkey(k) Then
                        ' 続きがある場合
                        If i < j - 1 And j > 0 Then
                            wkpkey = wkpkey & Cells(pkey_pos(k), C_COLNAME) & ","
                        Else
                            wkpkey = wkpkey & Cells(pkey_pos(k), C_COLNAME)
                        End If
                    End If
                Next k
            Next i
        End If
        
        ' CREATE文作成
        If wkpkey <> "" Then
            wktext = wktext & "/* PRIMARY KEY */" + vbCrLf
            '--- Mod Start 2010/09/21
            'wktext = wktext & "ALTER TABLE " & TableId + vbCrLf
            wktext = wktext & "ALTER TABLE " & strCommentTable + vbCrLf
            '--- Mod End
            
            '--- Mod Start 2012/02/27 TFC
            ' 主キー定義名作成
            pkeyName = "PK_" & TableId
            pkeyName = Left(pkeyName, 30)
            
            wktext = wktext & " ADD CONSTRAINT " & pkeyName & " PRIMARY KEY(" & wkpkey & ")"
            '--- Mod End 2012/02/27 TFC
        End If
        
        If pkey_pos(0) = 0 Then
            ElseIf Cells(pkey_pos(0), C_IdxSp2).Value <> "" Then
                '--- Add Start 2010/07/30
                Cells(pkey_pos(0), C_IdxSp2).Value = UCase(Cells(pkey_pos(0), C_IdxSp2).Value)
                '--- Add End
                wktext = wktext & " USING INDEX TABLESPACE " & Cells(pkey_pos(0), C_IdxSp2)
                wktext = wktext & ";" + vbCrLf
            ElseIf Cells(R_IdxSp, C_IdxSp).Value <> "" Then
                wktext = wktext & " USING INDEX TABLESPACE " & Cells(R_IdxSp, C_IdxSp).Value
                wktext = wktext & ";" + vbCrLf
            Else
                wktext = wktext & ";" + vbCrLf
        End If
        
        If wkpkey <> "" Then
            wktext = wktext & vbCrLf
        End If
        '--- Add Start 2010/07/29 OU
        '=====================================
        'インデックス作成
        '=====================================
        ' インデックス定義列参照ループ
        For intIndexC = C_IndexStart To C_IndexEnd Step 2
        
            strindex = ""
            strIndexSpace = ""
            intIndexR = R_COLNAME '物理名1行目
            
            '--- ADD Start 2019/06/21 SPC
            If (Cells(R_COLNAME - 1, intIndexC) <> "") Then
            '--- ADD End 2019/06/21 SPC
                '--- Mod Start 2012/02/27 TFC
                If (Cells(R_COLNAME - 1, intIndexC) = "FNC") Then
                
                    '=======================================
                    ' ■ファンクションインデックスの場合
                    ' インデックス定義領域に直接記入された文字列を
                    ' 文字連結する。
                    '=======================================
                
                    ' カラム参照ループ
                    While Cells(intIndexR, C_COLNAME) <> ""
                        
                        'インデックスが設定されている場合
                        If Cells(intIndexR, intIndexC) <> "" Then
                        
                            If strindex <> "" Then
                                strindex = strindex & ","
                            End If
                            ' ★記入された関数文字列をそのまま設定する。
                            strindex = strindex & Cells(intIndexR, intIndexC)
                        End If
                        
                        ' 表領域名未取得、かつ
                        ' インデックス定義枠に表領域リストが定義されている場合
                        If strIndexSpace = "" And Cells(intIndexR, C_IdxSp2) <> "" Then
                            ' 最初に出現した表領域名を使用する
                            strIndexSpace = Cells(intIndexR, C_IdxSp2)
                        End If
                        
                        intIndexR = intIndexR + 1
                    Wend
                
                Else
                    
                    '=======================================
                    ' ■ファンクションインデックス以外の場合
                    ' インデックス定義領域に設定された数値順に、
                    ' カラムIDを文字連結する。
                    '=======================================
                
                    ReDim idx(32768) As Integer
                    ReDim idx_seq(32768) As Integer
                    ReDim idx_pos(32768) As Integer
                    j = 0
                    
                    '-----------------
                    ' インデックス定義位置の保持と、表領域名取得
                    '-----------------
                    ' カラム参照ループ
                    While Cells(intIndexR, C_COLNAME) <> ""
                        
                        'インデックスが設定されている場合
                        If Cells(intIndexR, intIndexC) <> "" Then
                            
                            idx(j) = CInt(Cells(intIndexR, intIndexC))  ' インデックス設定値（数値変換後）格納
                            idx_pos(j) = intIndexR                      ' 行番号格納
                            idx_seq(j) = j + 1                          ' シーケンス番号格納
                            j = j + 1
                        End If
                        
                        ' 表領域名未取得、かつ
                        ' インデックス定義枠に表領域リストが定義されている場合
                        If strIndexSpace = "" And Cells(intIndexR, C_IdxSp2) <> "" Then
                            ' 最初に出現した表領域名を使用する
                            strIndexSpace = Cells(intIndexR, C_IdxSp2)
                        End If
                        
                        intIndexR = intIndexR + 1
                    Wend
                    
                    '-----------------
                    ' インデックス設定の項目ID文字列を作成
                    '-----------------
                    strindex = ""
                    
                    ' インデックス設定データを保持した場合
                    If j > 0 Then
                        ' シーケンス番号参照ループ
                        For i = 0 To j - 1
                            ' インデックス設定値参照ループ
                            For k = 0 To j - 1
                                If idx_seq(i) = idx(k) Then
                                    ' 続きがある場合
                                    If i < j - 1 And j > 0 Then
                                        strindex = strindex & Cells(idx_pos(k), C_COLNAME) & ","
                                    Else
                                        strindex = strindex & Cells(idx_pos(k), C_COLNAME)
                                    End If
                                End If
                            Next k
                        Next i
                    End If
                End If
                '--- Mod End 2012/02/27 TFC
            '--- ADD Start 2019/06/21 SPC
            End If
            '--- ADD End 2019/06/21 SPC
            
            ' CREATE文作成
            If strindex <> "" Then
                If intIndexC = C_IndexStart Then
                    wktext = wktext & "/* INDEX */" + vbCrLf
                End If
                
                '--- Mod Start 2012/02/27 TFC
                wktext = wktext & "CREATE"
                '-----------------
                ' インデックスの種類に応じて構文を変更
                '-----------------
                ' ユニークインデックス
                If (Cells(R_COLNAME - 1, intIndexC) = "UNQ") Then
                    wktext = wktext & " UNIQUE INDEX"
                    
                ' ビットマップインデックス
                ElseIf (Cells(R_COLNAME - 1, intIndexC) = "BMP") Then
                    wktext = wktext & " BITMAP INDEX"
                
                ' ノーマルインデックス、ファンクションインデックス
                Else
                    wktext = wktext & "        INDEX"
                    
                End If
                
                '-----------------
                ' インデックス定義名作成
                '-----------------
                ' ユニークインデックス
                If (Cells(R_COLNAME - 1, intIndexC) = "UNQ") Then
                    indexNamePrefix = "UDX"
                    
                ' ビットマップインデックス
                ElseIf (Cells(R_COLNAME - 1, intIndexC) = "BMP") Then
                    indexNamePrefix = "BDX"
                
                ' ファンクションインデックス
                ElseIf (Cells(R_COLNAME - 1, intIndexC) = "FNC") Then
                    indexNamePrefix = "FDX"
                    
                ' ノーマルインデックス
                Else
                    indexNamePrefix = "IDX"
                
                End If
                
                indexName = indexNamePrefix & CStr((intIndexC - C_IndexStart) / 2 + 1) & "_" & TableId
                ' 30バイト以内に調整
                indexName = Left(indexName, 30)
                wktext = wktext & " " & indexName
                
                wktext = wktext & " ON " & strCommentTable
                wktext = wktext & "(" & strindex & ")"
                
                '-----------------
                ' ローカルインデックスの場合はLOCAL句を追加
                '-----------------
                If (Cells(R_COLNAME - 2, C_IndexStart) = "LOCAL INDEX") Then
                    wktext = wktext & " LOCAL"
                End If
                '--- Mod End 2012/02/27 TFC
                
                If strIndexSpace <> "" Then
                    ' インデックス定義枠の表領域定義を使用する。
                    wktext = wktext & " TABLESPACE " & strIndexSpace
                ElseIf Cells(R_IdxSp, C_IdxSp).Value <> "" Then
                    ' ヘッダのインデックス表領域定義を使用する。
                    wktext = wktext & " TABLESPACE " & Cells(R_IdxSp, C_IdxSp).Value
                End If
                wktext = wktext & ";" + vbCrLf
            End If
            
        Next
        '--- Add End
    End If
    
    If ddlKind = DDL_KIND_INDEX And wktext = "" Then
        CreateDdl = ""
    Else
        CreateDdl = fileHeader & wktext
    End If
End Function

'--- Mod Start S.Iwanaga 2010/04/13
'==========================================================
'【関数名】LtoP
'【概　要】論理名を物理名に変換
'【引　数】strSheet     :処理対象シート名
'　　　　　lngStartRow  :処理開始行
'　　　　　lngRepeatCnt :処理を繰り返す回数
'【戻り値】0=正常終了　-1=異常終了
'==========================================================
Public Function LtoP(ByVal strSheet As String, _
                    ByVal lngStartRow As Long, _
                    ByVal lngRepeatCnt As Long) As Integer
On Error GoTo Err_Handler
    Dim intDiv      As Integer
    Dim intNoEntry  As Integer
    Dim lngFrom     As Long
    Dim strSrc      As String
    Dim strTmp      As String
    Dim strResult   As String
    Dim strConvFile As String
    Dim valRtn      As Variant
    Dim wsCurrent   As Worksheet
    Dim wbConv      As Workbook
    Dim wsConv      As Worksheet
    
    '処理中は画面の更新を止める
    Application.ScreenUpdating = False
    
    lngFrom = lngStartRow
    Set wsCurrent = ActiveWorkbook.Sheets(strSheet)
    
    '使用する変換テーブルファイルの指定
    If Len(ConvFilePath) > 0 Then
        strConvFile = ConvFilePath
    Else
        strConvFile = Workbooks(TEMPLATE).Path & "\" & CONVERT_LIST_FILE
    End If
    
    '変換テーブルのファイルを開く
    If Len(Dir(strConvFile)) = 0 Then
        '変換テーブルファイルが存在しない
        MsgBox " 変換用定義ファイルが見つかりません。" & vbCrLf & strConvFile, vbExclamation + vbOKOnly, "Error"
        GoTo Exit_Handler
    End If
    Set wbConv = Workbooks.Open(strConvFile, 0, True)
    Set wsConv = wbConv.Sheets(CONVERT_LIST_SHEET)

    '処理開始行から処理終了行までの論理名を物理名に変換
    With wsCurrent
        Do While lngStartRow + lngRepeatCnt >= lngFrom
            
            '物理名が入力されていない項目のみ処理する
            If .Cells(lngFrom, C_COLNAME).Value = "" Then
                strSrc = .Cells(lngFrom, C_ITEMNAME).Value
                intDiv = 0
                intNoEntry = 0
                strTmp = ""
                strResult = ""
                valRtn = Null
                While 0 < Len(strSrc)
                    strTmp = strSrc
                    Do While 0 < Len(strTmp)
                        valRtn = Application.VLookup(strTmp, wsConv.Range(DEFINITION_NAME), 2, False)
                        If (Not IsError(valRtn)) Then
                            If (0 < Len(strResult)) Then
                                strResult = strResult + "_"
                            End If
                            strResult = strResult + CStr(valRtn)
                            strSrc = Right(strSrc, Len(strSrc) - Len(strTmp))
                            intDiv = 1
                            Exit Do
                        End If
                        strTmp = Left(strTmp, Len(strTmp) - 1)
                    Loop
                    If (0 = Len(strTmp)) Then
                        If (1 = intDiv) Then
                            strResult = strResult + "_"
                        End If
                        strResult = strResult + Left(strSrc, 1)
                        strSrc = Right(strSrc, Len(strSrc) - 1)
                        intDiv = 0
                        intNoEntry = 1  '登録されていない項目名が存在する
                    End If
                Wend
                .Cells(lngFrom, C_COLNAME).Value = UCase(strResult)
                
                '物理名登録チェック
                If intNoEntry = 0 Then
                    '入力した全ての論理名が登録済み
                    .Cells(lngFrom, C_COLNAME).Interior.ColorIndex = xlColorIndexNone
                Else
                    '未登録の論理名が含まれる
                    .Cells(lngFrom, C_COLNAME).Interior.ColorIndex = 46
                End If
                
                '物理名byte数チェック
                If LenB(StrConv(strResult, vbFromUnicode)) > 30 Then
                    .Cells(lngFrom, C_COLNAME).Interior.ColorIndex = 3
                End If

            End If
            
            lngFrom = lngFrom + 1
        Loop
    End With
    
    LtoP = 0
    GoTo Exit_Handler
    
Err_Handler:
    LtoP = -1
    MsgBox Err.Description, vbCritical + vbOKOnly, "Error"
    
Exit_Handler:
    If wsConv Is Nothing = False Then Set wsConv = Nothing
    If wbConv Is Nothing = False Then wbConv.Close: Set wbConv = Nothing
    If wsCurrent Is Nothing = False Then Set wsCurrent = Nothing

    '画面の更新停止解除
    Application.ScreenUpdating = True
    
End Function
'--- Mod End

'--- Add Start S.Iwanaga 2010/04/08
'==========================================================
'【関数名】isTblDefSheet
'【概　要】対象シートがテーブル項目シートかチェックする
'【引　数】strSheet     :処理対象シート名
'【戻り値】True=テーブル項目シート　False=テーブル項目シート以外
'==========================================================
Public Function isTblDefSheet(ByVal strSheet As String) As Integer

    If ActiveWorkbook.Sheets(strSheet).Cells(R_SheetId, C_SheetId) = 2 Then
        isTblDefSheet = True
    Else
        isTblDefSheet = False
    End If

End Function

'==========================================================
'【関数名】chkExAreaState
'【概　要】拡張カラムの状態を取得
'【引　数】strSheet     :処理対象シート名
'【戻り値】0=非表示　1=表示
'==========================================================
Public Function chkExAreaState(ByVal strSheet As String) As Integer

    Dim ws  As Worksheet
    
    Set ws = ActiveWorkbook.Sheets(strSheet)
    
    If ws.Columns(C_HideSNm & ":" & C_HideENm).EntireColumn.Hidden = True Then
        chkExAreaState = 0
    Else
        chkExAreaState = 1
    End If
    
End Function
'--- Add End

'--- Add Start S.Iwanaga 2010/04/16
'==========================================================
'【関数名】chkBlankPName
'【概　要】物理名に未入力のセルがないかチェック
'【引　数】strSheet     :処理対象シート名
'【戻り値】0=未入力セルなし 1=未入力セルあり
'==========================================================
Public Function chkBlankPName(ByVal strSheet As String) As Integer

    'TODO: 未入力チェック処理を作成する

    chkBlankPName = 0
    
End Function

'==========================================================
'【関数名】setFFileData
'【概　要】型と桁数を基にフラットファイルの位置と桁情報をセットする
'【引　数】strSheet     :処理対象シート名
'【戻り値】なし
'==========================================================
Public Function setFFileData(ByVal strSheet As String)

    'TODO: フラットファイル情報セット処理を作成する

    
End Function



'==========================================================
'【関数名】createCtl
'【概　要】SQL*Loader制御ファイルデータ作成
'【引　数】strSheet     :処理対象シート名
'【戻り値】作成した制御ファイル文字列
'==========================================================
Public Function createCtl(ByVal strSheet As String) As String

    Dim intI            As Integer
    Dim intIndex        As Integer
    Dim strRtn          As String
    Dim strTableName    As String   'テーブル名
    Dim strLoadType     As String   'ロードオプション
    Dim strPName        As String   '物理名
    Dim strPosStart     As String   'フラットファイル開始位置
    Dim strPosEnd       As String   'フラットファイル終了位置
    Dim strFFileData()  As String
    '--- Add Start 2010/07/29 OU
    Dim strDataType     As String   'データタイプ
    Dim strDataTypeCSV As String
    strDataTypeCSV = "CSV"
    Dim strDataHead As String


    '--- Add End
    
    createCtl = ""
    
    'テーブル項目書に罫線を引き直す
    Call seisho
    
    'フラットファイル情報のセット
    Call setFFileData(strSheet)
    
    intIndex = 0
    intI = 0
    With ActiveWorkbook.Sheets(strSheet)
        strTableName = Trim(.Cells(R_TblId, C_TblId).Value)
        strLoadType = Trim(.Cells(R_LdOp, C_LdOp).Value)
        'Add Start 2010/07/29 OU
        strDataType = Trim(.Cells(R_DataTyp, C_DataTyp).Value)
        'Add End
        
        ReDim strFFileData(intIndex)
        strFFileData(intIndex) = getCTLHead & " " & strLoadType & " INTO TABLE " & strTableName & vbCrLf
        'Add Start 2010/07/29 OU
        If strDataType = strDataTypeCSV Then
            strFFileData(intIndex) = strFFileData(intIndex) & " FIELDS TERMINATED BY "",""" & vbCrLf
        End If
        'Add End
        strFFileData(intIndex) = strFFileData(intIndex) & "  (" & vbCrLf
        
        '物理名セルの値が空になるまで繰り返す
        Do While .Cells(R_COLNAME + intI, C_COLNAME).Value <> ""
            
            strPName = Trim(.Cells(R_COLNAME + intI, C_COLNAME).Value)
            intIndex = UBound(strFFileData) + 1
            ReDim Preserve strFFileData(intIndex)
            
            strFFileData(intIndex) = "    " & strPName

            If strDataType <> strDataTypeCSV Then
                strPosStart = Trim(.Cells(R_COLNAME + intI, C_FFilePosition).Value)
                strPosEnd = CStr(CInt(Trim(.Cells(R_COLNAME + intI, C_FFilePosition).Value)) + (CInt(Trim(.Cells(R_COLNAME + intI, C_FFileLength).Value) - 1)))
                strFFileData(intIndex) = strFFileData(intIndex) & "  POSITION(" & strPosStart & ":" & strPosEnd & ")"
            End If
            
            
            intI = intI + 1
            If .Cells(R_COLNAME + intI, 1).Value = "" Then
                strFFileData(intIndex) = strFFileData(intIndex) & vbCrLf
            Else
                strFFileData(intIndex) = strFFileData(intIndex) & "," & vbCrLf
            End If
        Loop
        
        intIndex = UBound(strFFileData) + 1
        ReDim Preserve strFFileData(intIndex)
        
        strFFileData(intIndex) = "  )"

    End With
    
    strRtn = ""
    For intI = 0 To UBound(strFFileData)
        strRtn = strRtn & strFFileData(intI)
    Next intI
    
    createCtl = strRtn

End Function
'--- Add End

'--- Add Start 2010/07/29 OU
Function getCTLHead() As String
    Dim strDirectPath As String
    strDirectPath = Trim(Cells(R_DirectPath, C_DirectPath).Value)

    strDataHead = strDataHead & " -- *****************************************************" & vbCrLf
    strDataHead = strDataHead & " -- SQL*LOADER 制御ファイル" & vbCrLf
    strDataHead = strDataHead & " -- テーブルID :" & Cells(R_TblId2, C_TblId2).Value & vbCrLf
    strDataHead = strDataHead & " -- テーブル名称 :" & Cells(R_TblNm, C_TblNm).Value & vbCrLf
    strDataHead = strDataHead & " -- 作成日 : " & Format(Date, "yyyy/mm/dd") & " Ver.1.0" & vbCrLf
    strDataHead = strDataHead & " -- 更新履歴 : " & vbCrLf
    strDataHead = strDataHead & " -- *****************************************************" & vbCrLf
    If UCase(strDirectPath) = "TRUE" Then
        strDataHead = strDataHead & " OPTIONS (ERRORS=0,DIRECT=TRUE)" & vbCrLf
        strDataHead = strDataHead & " UNRECOVERABLE" & vbCrLf
    Else
        strDataHead = strDataHead & " OPTIONS (ERRORS=0,DIRECT=FALSE)" & vbCrLf
    End If
    
    strDataHead = strDataHead & " LOAD DATA" & vbCrLf
    strDataHead = strDataHead & " INFILE "".DAT""" & vbCrLf
    
    getCTLHead = strDataHead
    
End Function
'--- Add End
