Attribute VB_Name = "Global_PosGet"

'==========================================================
'【プロシージャ名】Tb_Posget
'【概　要】テーブル定義書のカラムポジションを取得
'【引　数】なし
'【戻り値】なし
'==========================================================

Sub Tb_Posget()
    Dim bname As String
    Dim sname As String
    '--- MOD Start 2015/02/27 TFC
    bname = ThisWorkbook.name
    '--- MOD End 2015/02/27 TFC
    sname = Workbooks(bname).Worksheets("プロパティ").name
    Tb_SheetNm = CStr(Workbooks(bname).Worksheets(sname).Cells(4, 5).Value) 'テーブル項目　シート名
    R_COLNAME = CInt(Workbooks(bname).Worksheets(sname).Cells(9, 8).Value) 'テーブル物理名　列     項目ID
    C_COLNAME = CInt(Workbooks(bname).Worksheets(sname).Cells(10, 8).Value) 'テーブル物理名　カラム   項目ID
    R_ITEMNAME = CInt(Workbooks(bname).Worksheets(sname).Cells(56, 8).Value) 'テーブル項目名　行  項目名カラムPOS
    C_ITEMNAME = CInt(Workbooks(bname).Worksheets(sname).Cells(57, 8).Value) 'テーブル項目名　カラム   項目名カラムPOS
    R_TblId = CInt(Workbooks(bname).Worksheets(sname).Cells(3, 8).Value) 'テーブルID位置(隠し)
    C_TblId = CInt(Workbooks(bname).Worksheets(sname).Cells(4, 8).Value) 'テーブルID位置(隠し)
    R_TblNm = CInt(Workbooks(bname).Worksheets(sname).Cells(6, 8).Value) 'テーブル名位置(隠し)
    C_TblNm = CInt(Workbooks(bname).Worksheets(sname).Cells(7, 8).Value) 'テーブル名位置(隠し)
    C_KeiEnd = CInt(Workbooks(bname).Worksheets(sname).Cells(13, 8).Value)
    C_HideSta = CInt(Workbooks(bname).Worksheets(sname).Cells(16, 8).Value)
    C_HideEnd = CInt(Workbooks(bname).Worksheets(sname).Cells(16, 8).Value)
    R_TblId2 = CInt(Workbooks(bname).Worksheets(sname).Cells(47, 8).Value) 'テーブルID位置(見出し）
    C_TblId2 = CInt(Workbooks(bname).Worksheets(sname).Cells(48, 8).Value) 'テーブルID位置(見出し）
    R_Schima = CInt(Workbooks(bname).Worksheets(sname).Cells(21, 8).Value) 'スキーマ名
    C_Schima = CInt(Workbooks(bname).Worksheets(sname).Cells(22, 8).Value) 'スキーマ名
    R_TblSp = CInt(Workbooks(bname).Worksheets(sname).Cells(24, 8).Value) 'テーブル表領域
    C_TblSp = CInt(Workbooks(bname).Worksheets(sname).Cells(25, 8).Value) 'テーブル表領域
    R_DataTyp = CInt(Workbooks(bname).Worksheets(sname).Cells(27, 8).Value) 'データタイプ
    C_DataTyp = CInt(Workbooks(bname).Worksheets(sname).Cells(28, 8).Value) 'データタイプ
    R_LdOp = CInt(Workbooks(bname).Worksheets(sname).Cells(30, 8).Value) 'ロードオプション
    C_LdOp = CInt(Workbooks(bname).Worksheets(sname).Cells(31, 8).Value) 'ロードオプション
    R_IdxSp = CInt(Workbooks(bname).Worksheets(sname).Cells(33, 8).Value) 'INDEX表領域
    C_IdxSp = CInt(Workbooks(bname).Worksheets(sname).Cells(34, 8).Value) 'INDEX表領域
    R_Create = CInt(Workbooks(bname).Worksheets(sname).Cells(53, 8).Value) '作成日
    C_Create = CInt(Workbooks(bname).Worksheets(sname).Cells(54, 8).Value) '作成日
    R_TblNm2 = CInt(Workbooks(bname).Worksheets(sname).Cells(50, 8).Value) 'テーブル名位置(見出し）
    C_TblNm2 = CInt(Workbooks(bname).Worksheets(sname).Cells(51, 8).Value) 'テーブル名位置(見出し）
    C_printsta = Trim(Workbooks(bname).Worksheets(sname).Cells(63, 8).Value) '印刷範囲列
    C_printend = Trim(Workbooks(bname).Worksheets(sname).Cells(64, 8).Value) '印刷範囲列
    '--- Add Start OU 2010/07/29
    R_DirectPath = CInt(Workbooks(bname).Worksheets(sname).Cells(66, 8).Value) 'ダイレクトパス
    C_DirectPath = CInt(Workbooks(bname).Worksheets(sname).Cells(67, 8).Value) 'ダイレクトパス
    C_IndexStart = CInt(Workbooks(bname).Worksheets(sname).Cells(69, 8).Value) 'Indexキー開始列
    C_IndexEnd = CInt(Workbooks(bname).Worksheets(sname).Cells(71, 8).Value) 'Indexキー最終列
    '--- Add End
    '--- ADD Start 2019/07/19 SPC
    R_PartitionKind = CInt(Workbooks(bname).Worksheets(sname).Cells(73, 8).Value) ' パーティション種類行
    C_PartitionKind = CInt(Workbooks(bname).Worksheets(sname).Cells(74, 8).Value) ' パーティション種類列
    R_PartitionKoumoku = CInt(Workbooks(bname).Worksheets(sname).Cells(76, 8).Value) ' パーティション対象項目行
    C_PartitionKoumoku = CInt(Workbooks(bname).Worksheets(sname).Cells(77, 8).Value) ' パーティション対象項目列
    '--- ADD End 2019/07/19 SPC

    '--- Add Start S.Iwanaga 2010/04/08
    'ドキュメントID位置情報取得
    R_DocId = CInt(Workbooks(bname).Worksheets(sname).Cells(2, 11).Value)
    C_DocId = CInt(Workbooks(bname).Worksheets(sname).Cells(3, 11).Value)
    'シートID位置情報取得
    R_SheetId = CInt(Workbooks(bname).Worksheets(sname).Cells(5, 11).Value)
    C_SheetId = CInt(Workbooks(bname).Worksheets(sname).Cells(6, 11).Value)
    '非表示カラム開始終了列名情報取得
    C_HideSNm = Trim(Workbooks(bname).Worksheets(sname).Cells(18, 8).Value)
    C_HideENm = Trim(Workbooks(bname).Worksheets(sname).Cells(19, 8).Value)
    '--- Add End
    '物理名変換テーブルファイルパス
    ConvFilePath = Trim(Workbooks(bname).Worksheets(sname).Cells(1, 14).Value)  '--- Add S.Iwanaga 2010/04/13

    '--- Add Start S.Iwanaga 2010/04/16
    'フラットファイル位置
    C_FFilePosition = Trim(Workbooks(bname).Worksheets(sname).Cells(59, 8).Value)
    'フラットファイル桁
    C_FFileLength = Trim(Workbooks(bname).Worksheets(sname).Cells(61, 8).Value)
    '--- Add End

    'テーブル定義書各項目見出しから内容をセットする為のカラム位置を取得する
    C_kata = colposget("型")
    C_keta = colposget("桁数")
    C_shou = colposget("小数")
    C_primary = colposget("主キー")
    C_uniq = colposget("一意")
    C_nnul = colposget("必須")
    C_check = colposget("チェック制約")
    C_def = colposget("デフォルト値")
    'インデックス表領域だけは行が違うので関数を使わない
    For i = C_COLNAME To C_KeiEnd
        If Workbooks(bname).Worksheets("テーブル項目").Cells(R_COLNAME - 1, i).Value = "表領域" Then
            C_IdxSp2 = i
            Exit For
        End If
    Next i
End Sub


