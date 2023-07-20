Attribute VB_Name = "Global_declea"
'グローバル変数定義　不要なものあり

Public Type byteData
    DataBody() As Byte
    DataLen As Long
End Type
Public R_COLNAME As Integer
Public C_COLNAME As Integer
Public R_ITEMNAME As Integer    '項目名カラムPOS(行)
Public C_ITEMNAME As Integer    '項目名カラムPOS(列)
Public R_TblId As Integer
Public C_TblId As Integer
Public R_TblNm As Integer
Public C_TblNm As Integer
Public R_TblNm2 As Integer
Public C_TblNm2 As Integer
Public C_KeiEnd As Integer
Public C_HideSta As Integer
Public C_HideEnd As Integer
Public R_Schima As Integer
Public C_Schima As Integer
Public R_TblSp As Integer
Public C_TblSp As Integer
Public R_DataTyp As Integer
Public C_DataTyp As Integer
Public R_LdOp As Integer
Public C_LdOp As Integer
Public R_IdxSp As Integer
Public C_IdxSp As Integer
Public R_TblId2 As Integer
Public C_TblId2 As Integer
Public C_kata As Integer
Public C_keta As Integer
Public C_shou As Integer
Public C_primary As Integer
Public C_uniq As Integer
Public C_nnul As Integer
Public C_check As Integer
Public C_def As Integer
Public C_IdxSp2 As Integer
Public R_Create As Integer
Public C_Create As Integer
'--- Add Start S.Iwanaga 2010/04/08
Public R_DocId      As Integer  'ドキュメントID位置行番号
Public C_DocId      As Integer  'ドキュメントID位置列番号
Public R_SheetId    As Integer  'シートID位置行番号
Public C_SheetId    As Integer  'シートID位置列番号
Public C_HideSNm    As String   '非表示カラムの開始列名
Public C_HideENm    As String   '非表示カラムの終了列名
'--- Add End
Public ConvFilePath As String   '物理名変換テーブルファイルパス --- Add S.Iwanaga 2010/04/13
'--- Add Start S.Iwanaga 2010/04/16
Public C_FFilePosition  As Integer  'フラットファイル位置列番号
Public C_FFileLength    As Integer  'フラットファイル桁列番号
'--- Add End

Public Tb_SheetNm As String
Public Tb_SheetNMInp As String
Public C_printsta As String
Public C_printend As String

'--- Add Start OU 2010/07/29
Public R_DirectPath As Integer 'ダイレクトパス
Public C_DirectPath As Integer 'ダイレクトパス
Public C_IndexStart As Integer 'Indexキー開始列
Public C_IndexEnd  As Integer 'Indexキー最終列
'--- Add End

'--- ADD Start 2019/07/19 SPC
Public R_PartitionKind As Integer ' パーティション種類行
Public C_PartitionKind As Integer ' パーティション種類列
Public R_PartitionKoumoku As Integer ' パーティション対象項目行
Public C_PartitionKoumoku As Integer ' パーティション対象項目列
'--- ADD End 2019/07/19 SPC


' ********* 変数 *********
Public GV_book As String
Public GV_recLeng As Integer
Public GV_txtRecLeng As Integer
Public GV_compCalc As Boolean
Public GV_itemType As String
Public GV_itemKeta1 As Integer
Public GV_itemKeta2 As Integer
Public GV_itemByte As Integer
Public GV_saveFont As String
Public GV_saveFontSize As Integer
Public GV_fileBuff As Integer
Public GV_grpArray(3) As String
Public GV_occArray(3) As String
Public GV_subNo(2) As Integer
Public GV_readRow As Integer
Public GV_convFlag As String
Public GV_lv1Leng As Integer
Public GV_lv2Leng As Integer
Public GV_pgmid As String
Public GV_dataStop As byteData
' ********* 定数 *********
Public Const TEMPLATE As String = "voyager.xla"
Public Const FILEFILTER_ALL As String = "すべてのﾌｧｲﾙ (*.*),*.*"
Public Const GV_TBNM_COL As Integer = 5
Public Const GV_TBNM_ROW As Integer = 2

' *** ファイル項目 ***
Public Const FILEITEM As String = "ファイル項目"
Public Const HELPSHEET As String = "ヘルプ"
Public Const SAMPLESHEET As String = "sample"
Public Const F_STARTROW As Integer = 5
Public Const F_ROW_PAGE As Integer = 19
Public Const F_EXIST_COL As Integer = 1 'A  NULLなら項目定義は終了。
Public Const F_LV1_COL As Integer = 2 'B
Public Const F_LV2_COL As Integer = 3 'C
Public Const F_LV3_COL As Integer = 4 'D
Public Const F_OCC_COL As Integer = 5 'E
Public Const F_ATTR_COL As Integer = 6 'F
Public Const F_BYTE_COL As Integer = 7 'G
Public Const F_KEY_COL As Integer = 8 'H
Public Const F_HISSU_COL As Integer = 9 'I
Public Const F_SETUMEI_COL As Integer = 10 'J
Public Const F_NAME_COL As Integer = 11 'K
Public Const F_KAISI_COL As Integer = 12 'L
Public Const F_TYPE_COL As Integer = 13 'M
Public Const F_SEISU_COL As Integer = 14 'N
Public Const F_SYOSU_COL As Integer = 15 'O
Public Const F_RECSIZE_CELL As String = "L2"
Public Const F_TXTPGM_CELL As String = "N2"
Public Const F_TXRPGM_CELL As String = "O2"
' *** ファイル項目 (2) or ワークシート***
Public Const FILEITEM2 As String = "ファイル項目 (2)"
Public Const F_WKS_GONLY As Integer = 0
Public Const F_WKS_ITEMS As Integer = 1
' *** ファイルレイアウト ***
Public Const FILELAYOUT As String = "ファイルレイアウト"
Public Const FILELAYOUT_WK As String = "レイアウトワーク"
Public Const L_STARTROW As Integer = 1
Public Const L_STARTCOL As Integer = 4  'D
Public Const L_ENDCOL As Integer = 53   'BA
Public Const L_DAN_PAGE As Integer = 5
Public Const L_ROW_DAN As Integer = 9
Public Const L_BYTE_DAN As Integer = 50
Public Const L_DATE_ROW As Integer = 1
Public Const L_AUTHOR_ROW As Integer = 3
Public Const L_FILEID_ROW As Integer = 5
Public Const L_FILENAME_ROW As Integer = 7
Public Const L_RECSIZE_ROW As Integer = 12
Public Const L_COPY_ROW As Integer = 16
Public Const L_LABEL_COL As Integer = 55 'BC
Public Const L_TITLE_COL As Integer = 56 'BD
Public Const L_COPYSRC_ROWS As String = "28:36"
Public Const L_COLIDX_STARTROW As Integer = 8
' *** 登録集 ***
Public Const CBLCOPY As String = "登録集"
' *** データ入力 ***
Public Const DATAINPUT As String = "データ入力"
Public Const DATASHEET_WK As String = "データシートワーク"
Public Const D_EXIST_ROW As Integer = 4 '  NULLならデータ終了
Public Const D_STARTCOL As Integer = 16 'P
Public Const D_STARTROW As Integer = 5
Public Const D_FILENAME_CELL As String = "A2"
Public Const D_FILEID_CELL As String = "F2"
Public Const DATASTOP As String = "~~__!! DataWrite Stop ##@@^^"
' *** データ表示 ***
Public Const DATADISPLAY As String = "データ表示"
' *** テキスト化プログラム ***
Public Const TXTCONVPGM As String = "テキスト化PGM"
Public Const TXTREVPGM As String = "逆テキスト化PGM"
Public Const TXTCONVSKL As String = "テキスト化SKL"
Public Const TXTCONV_WK As String = "テキスト化ワーク"
Public Const X_STARTROW As Integer = 2
Public Const X_TITLE_ROW As Integer = 3
Public Const X_PGMAIM_ROW As Integer = 5
Public Const X_FILEID_ROW As Integer = 6
Public Const X_DATE_ROW As Integer = 8
Public Const X_PROGRAMID_ROW As Integer = 15
Public Const X_DATEWRITTEN_ROW As Integer = 17
Public Const X_IRECCHARS_ROW As Integer = 35
Public Const X_INREC_ROW As Integer = 39
Public Const X_ORECCHARS_ROW As Integer = 43
Public Const X_OUTREC_ROW As Integer = 47
Public Const X_CLRREC_ROW As Integer = 70
Public Const X_BUFFLEN_ROW As Integer = 69
Public Const X_FILEVAR_ROW As Integer = 71
Public Const X_PGMIDVAR_ROW As Integer = 72
Public Const X_PREGEN_ROW As Integer = 73
Public Const X_WKS_GROUP1 As String = " (2)"
Public Const X_WKS_GROUP2 As String = " (3)"
Public Const X_WKS_NUMBER As String = " (4)"
Public Const X_ORG_DEF As Integer = 1
Public Const X_TXT_DEF As Integer = 2
' *** その他 ***
Public Const F_GROUP As String = "G" ' 集合項目
Public Const F_XTYPE As String = "X" ' 文字項目
Public Const F_NTYPE As String = "N" ' 2BYTE文字項目
Public Const F_9TYPE As String = "9" ' 数値項目（通常Z,P,Cを使用)
Public Const F_ZTYPE As String = "Z" ' ZONE
Public Const F_PTYPE As String = "P" ' PACKED , COMP-3
Public Const F_CTYPE As String = "C" ' BINARY , COMP
Public Const F_SZTYPE As String = "SZ" ' 符号付きZONE
Public Const F_SPTYPE As String = "SP" ' 符号付きPACK
Public Const F_SCTYPE As String = "SC" ' 符号付きCOMP
' *** COBOL属性チェック結果
Public Const F_OK As Integer = 0 'All Right
Public Const F_ERR_BYTE As Integer = 1 'BYTES ERROR
Public Const F_ERR_TYPE As Integer = 2 'TYPE ERROR
Public Const F_ERR_KETA1 As Integer = 3 'SEISU-KETA ERROR
Public Const F_ERR_KETA2 As Integer = 4 'SYOSU-KETA ERROR
Public Const F_COMPATIBLE As Integer = -1 'TYPE NOT EQUAL, BUT COMPATIBLE
' *** ColorIndex
Public Const COLOR_AUTO As Integer = xlAutomatic
Public Const COLOR_NONE As Integer = 0
Public Const COLOR_WHITE As Integer = 2
Public Const COLOR_LIGHTYELLOW As Integer = 36

' ********* 変数 *********
Public GV_charOra As String
Public GV_unloadBuff As Integer
Public GV_msgtitle As String
Public GV_oracleVersion As Integer
Public GV_ddlSheet As String
Public GV_sheetBreak As Integer
Public GV_unlim As Boolean
Public GV_primary As String
Public GV_ddlBunkatu As Boolean

' ********* 定数 *********
Public Const SHEET_MAX_ROW As Integer = 16384
Public Const SHEET_BREAK_ROW As Integer = 16000
Public Const SHEET_BREAK_NAME As String = "継続"
Public Const SHEET_CONTINUE As String = "/*** TO BE CONTINUED NEXT SHEET ***/"

Public Const DDL_64K_LINES As Integer = 560
Public Const DDL_PART_LINES As Integer = 12

' *** テーブル項目 ***
Public Const TABLEITEM As String = "テーブル項目"
Public Const TABLEITEM_WK As String = "テーブルワーク"
Public Const TABLEITEM_FILE As String = "テーブル項目選択"
Public Const T_STARTROW As Integer = 5
Public Const T_ROW_PAGE As Integer = 19
Public Const T_EXIST_COL As Integer = 1 'A
Public Const T_ITEM_COL As Integer = 2 'B
Public Const T_TYPE_COL As Integer = 3 'C
Public Const T_KETA_COL As Integer = 4 'D
Public Const T_PKEY_COL As Integer = 5 'E
Public Const T_FKEY_COL As Integer = 6 'F
Public Const T_HISSU_COL As Integer = 7 'G
Public Const T_SETUMEI_COL As Integer = 8 'H
Public Const T_CHECK_COL As Integer = 9 'I
Public Const T_NAME_COL As Integer = 10 'J
Public Const T_COBNAME_COL As Integer = 11 'K
Public Const T_COBATTR_COL As Integer = 12 'L
Public Const T_COBBYTE_COL As Integer = 13 'M
Public Const T_COBTYPE_COL As Integer = 14 'N
Public Const T_COBSEISU_COL As Integer = 15 'O
Public Const T_COBSYOSU_COL As Integer = 16 'P
Public Const T_COLSIZE_COL As Integer = 17 'Q
Public Const T_COBRECSIZE_CELL As String = "L2"
Public Const T_ROW_SPACE_CELL As String = "Q2"
Public Const T_UNLPGM_CELL As String = "K2"
Public Const T_VERSION_CELL As String = "Z1"
Public Const T_STORAGEID_CELL As String = "AQ2"
Public Const T_VARLEN_COL As Integer = 26 'Z
Public Const T_SKEY1_COL As Integer = 27 'AA
Public Const T_SKEY2_COL As Integer = 28 'AB
Public Const T_SKEY3_COL As Integer = 29 'AC
Public Const T_SKEY4_COL As Integer = 30 'AD
Public Const T_SKEY5_COL As Integer = 31 'AE
Public Const T_SKEY6_COL As Integer = 32 'AF
Public Const T_SKEY7_COL As Integer = 33 'AG
Public Const T_SKEY8_COL As Integer = 34 'AH
Public Const T_SKEY9_COL As Integer = 35 'AI
Public Const T_STORAGELABEL_COL As Integer = 41 'AO
Public Const T_TBLSTORAGE_COL As Integer = 42 'AP
Public Const T_PKSTORAGE_COL As Integer = 43 'AQ
Public Const T_SK1STORAGE_COL As Integer = 44 'AR
Public Const T_SK2STORAGE_COL As Integer = 45 'AS
Public Const T_SK3STORAGE_COL As Integer = 46 'AT
Public Const T_SK4STORAGE_COL As Integer = 47 'AU
Public Const T_SK5STORAGE_COL As Integer = 48 'AV
Public Const T_SK6STORAGE_COL As Integer = 49 'AW
Public Const T_SK7STORAGE_COL As Integer = 50 'AX
Public Const T_SK8STORAGE_COL As Integer = 51 'AY
Public Const T_SK9STORAGE_COL As Integer = 52 'AZ
Public Const T_BLOCKSIZE_ROW As Integer = 5
Public Const T_BYTEROW_ROW As Integer = 6
Public Const T_ROWBLOCK_ROW As Integer = 7
Public Const T_INITROW_ROW As Integer = 8
Public Const T_MAXROW_ROW As Integer = 9
Public Const T_PCTFREE_ROW As Integer = 10
Public Const T_PCTUSED_ROW As Integer = 11
Public Const T_EXPAND_ROW As Integer = 12
Public Const T_TABLESPACE_ROW As Integer = 13
Public Const T_INITIAL_ROW As Integer = 14
Public Const T_NEXT_ROW As Integer = 15
Public Const T_MAXEXTENTS_ROW As Integer = 16
Public Const T_PARTITION_ROW As Integer = 17
Public Const T_INITVOLUME_ROW As Integer = 18
Public Const T_MAXVOLUME_ROW As Integer = 19
Public Const T_MAXEXTVOL_ROW As Integer = 20
Public Const T_INITLAYER_ROW As Integer = 21
Public Const T_MAXLAYER_ROW As Integer = 22
Public Const T_ITEMNUM_ROW As Integer = 23

' *** ＤＤＬ ***
Public Const S_STARTROW As Integer = 1
Public Const S_TABLENAME_ROW As Integer = 2
Public Const S_TABLEID_ROW As Integer = 3
Public Const S_MAKE_ROW As Integer = 4
Public Const S_PREGEN_ROW As Integer = 6
Public Const S_TABLE_STR As String = "/*** テーブル作成部 ***/"
Public Const S_INDEX_STR As String = "/*** インデックス作成部 ***/"


' *** アンロードプログラム ***
Public Const UNLOADPGM As String = "アンロードPGM"
Public Const UNLOADSKL As String = "アンロードSKL"
Public Const U_STARTROW As Integer = 2
Public Const U_TITLE_ROW As Integer = 3
Public Const U_TABLEID_ROW As Integer = 6
Public Const U_DATE_ROW As Integer = 8
Public Const U_PROGRAMID_ROW As Integer = 15
Public Const U_DATEWRITTEN_ROW As Integer = 17
Public Const U_RECCHARS_ROW As Integer = 38
Public Const U_OUTREC_ROW As Integer = 42
Public Const U_CLRREC_ROW As Integer = 63
Public Const U_BUFFLEN_ROW As Integer = 58
Public Const U_TABLEVAR_ROW As Integer = 65
Public Const U_PGMIDVAR_ROW As Integer = 66
Public Const U_PREGEN_ROW As Integer = 67
Public Const U_TAB1 As String = "            "
Public Const U_TAB2 As String = "                "
Public Const U_COB_DEF As Integer = 1
Public Const U_HCOB_DEF As Integer = 2
Public Const U_TABLE_SELECT As Integer = 3
Public Const U_HCOB_SET As Integer = 4
' *** ロード制御ファイル ***
Public Const LOADCTL As String = "ロードCTL"
Public Const LOADSKL As String = "ロードSKL"
Public Const P_STARTROW As Integer = 1
Public Const P_TABLEID_ROW As Integer = 3
Public Const P_TABLENAME_ROW As Integer = 4
Public Const P_RECSIZE_ROW As Integer = 5
Public Const P_MAKE_ROW As Integer = 6
Public Const P_PREGEN_ROW As Integer = 8
' *** エクスポートパラメータ ***
Public Const EXPPAR As String = "EXPORTパラメータ"
' *** エクスポートパラメータ ***
Public Const IMPPAR As String = "IMPORTパラメータ"
' *** その他 ***
Public Const T_CHAR As String = "CHAR"
Public Const T_VARCHAR2 As String = "VARCHAR2"
Public Const T_NUMBER As String = "NUMBER"
Public Const T_INTEGER As String = "INTEGER"
Public Const T_DATE As String = "DATE"
Public Const T_RAW As String = "RAW"
Public Const T_ROWID As String = "ROWID"
Public Const T_UB1 As Integer = 1
Public Const T_UB4 As Integer = 4
Public Const T_SB2 As Integer = 2
Public Const T_KDBT As Integer = 4
Public Const T_KDBH As Integer = 14
Public Const T_KTBIT As Integer = 24
Public Const T_KTBBH As Integer = 48
Public Const T_KCBH As Integer = 20
Public Const T_DB_BLOCK_SIZE As String = "8K"
Public Const T_PCTFREE As Integer = 10
Public Const T_PCTUSED As Integer = 40
Public Const T_INITRANS_TBL As Integer = 4
Public Const T_INITRANS_IDX As Integer = 2
Public Const T_FIXED_HEADER As Integer = 113
Public Const T_ENTRY_HEADER As Integer = 2
Public Const T_BTREE As String = "B*TREE"
Public Const T_CLUSTERED As String = "CLUSTERED"
Public Const T_BITMAP As String = "BITMAP"
Public Const T_UNIQUE As String = "UNIQUE"
Public Const T_LOCAL As String = "LOCAL"
Public Const T_UNLIMITED As Long = -1
Public Const T_PRIMARY As String = "PRIMARY"
Public Const T_PRIMARY_INDEX As String = "PRIMARY_INDEX"
Public Const T_EXTENT_BASE As Long = 1024
'********************* 変数 ***
Public GV_listBook As String
Public GV_multiThread As Boolean
'Public GV_totalVolK As Long
'Public GV_totalAvailK As Long
Public GV_partNaming As String

'********************* 定数 ***
' *** パーティション ***
Public Const PARTITION As String = "パーティション"
Public Const R_STARTROW As Integer = 6
Public Const R_EXIST_COL As Integer = 1 'A
Public Const R_IDNAME_COL As Integer = 2 'B
Public Const R_KIND_COL As Integer = 3 'C
Public Const R_PARTITEM_COL As Integer = 4 'D
Public Const R_PARTVAL_COL As Integer = 5 'E
Public Const R_PARTNAME_COL As Integer = 6 'F
Public Const R_MAXROW_COL As Integer = 7 'G
Public Const R_MAXVOL_COL As Integer = 8 'H
Public Const R_TBLSPACE_COL As Integer = 9 'I
Public Const R_PCTFREE_COL As Integer = 10 'J
Public Const R_PCTUSED_COL As Integer = 11 'K
Public Const R_INITEXT_COL As Integer = 12 'L
Public Const R_NEXTEXT_COL As Integer = 13 'M
Public Const R_MAXEXT_COL As Integer = 14 'N
Public Const R_MAXEXTVOL_COL As Integer = 15 'O
Public Const R_MAXLAYER_COL As Integer = 16 'P
Public Const R_KINDCODE_COL As Integer = 17 'Q
Public Const R_SOURCE_COL As Integer = 18 'R
Public Const R_SUBNAME_CELL As String = "B2"
Public Const R_TBLNAME_CELL As String = "E2"
Public Const R_VERSION_CELL As String = "P1"
Public Const R_PART_STR As String = "  (Partition)"
Public Const R_PART_STR2 As String = " Partitions"

'*** テーブルサイズ一覧 ***
Public Const TABLELIST As String = "テーブルサイズ一覧"
Public Const TABLELIST2 As String = "テーブルサイズ一覧 (2)"
Public Const TABLELIST_WK As String = "テーブルサイズ一覧(BAK)"
Public Const Z_STARTROW As Integer = 6
Public Const Z_EXIST_COL As Integer = 1 'A
Public Const Z_TBLNAME_COL As Integer = 2 'B
Public Const Z_TBLID_COL As Integer = 3 'C
Public Const Z_TBLSPACE_COL As Integer = 4 'D
Public Const Z_ROWSIZE_COL As Integer = 5 'E
Public Const Z_PCTFREE_COL As Integer = 6 'F
Public Const Z_PCTUSED_COL As Integer = 7 'G
Public Const Z_INITEXT_COL As Integer = 8 'H
Public Const Z_NEXTEXT_COL As Integer = 9 'I
Public Const Z_MAXEXT_COL As Integer = 10 'J
Public Const Z_MAXROW_COL As Integer = 11 'K
Public Const Z_MAXVOL_COL As Integer = 12 'L
Public Const Z_MAXEXTVOL_COL As Integer = 13 'M
Public Const Z_OBJTYPE_COL As Integer = 14 'N
Public Const Z_TBLIDX_COL As Integer = 15 'O
Public Const Z_SUBNAME_CELL As String = "B2"
Public Const Z_TOTALVOL_CELL As String = "D2"
Public Const Z_BLOCKSIZE_CELL As String = "E2"
Public Const Z_TOTALEXTVOL_CELL As String = "G2"
Public Const Z_TABLEEXTVOL_CELL As String = "I2"


'*** テーブルスペース別オブジェクト一覧 ***
Public Const TSTABLELIST As String = "スペース別オブジェクト一覧"
Public Const TSTABLELIST2 As String = "スペース別オブジェクト一覧 (2)"
Public Const TSTABLELIST_WK As String = "スペース別オブジェクトワーク"
Public Const TSTABLELIST_INPUT As String = "RBS・TMP定義"
Public Const V_STARTROW As Integer = 6
Public Const V_EXIST_COL As Integer = 1 'A
Public Const V_TBLSPACE_COL As Integer = 2 'B
Public Const V_OBJID_COL As Integer = 3 'C
Public Const V_OBJTYPE_COL As Integer = 4 'D
Public Const V_TABLEVOL_COL As Integer = 5 'E
Public Const V_TBLSPACEVOL_COL As Integer = 6 'F
Public Const V_PCTFREE_COL As Integer = 7 'G
Public Const V_PCTUSED_COL As Integer = 8 'H
Public Const V_INITEXT_COL As Integer = 9 'I
Public Const V_NEXTEXT_COL As Integer = 10 'J
Public Const V_MINEXT_COL As Integer = 11 'K
Public Const V_MAXEXT_COL As Integer = 12 'L
Public Const V_MISC_COL As Integer = 13 'M
Public Const V_OBJGROUP_COL As Integer = 14 'N
Public Const V_PARTSEQ_COL As Integer = 15 'O
Public Const V_TBLIDX_COL As Integer = 16 'P
Public Const V_TSGROUP_COL As Integer = 17 'Q
Public Const V_TSSEQ_COL As Integer = 18 'R
Public Const V_SUBNAME_CELL As String = "B2"
Public Const V_DBNAME_CELL As String = "D2"
Public Const V_BLOCKSIZE_CELL As String = "F2"
Public Const V_TOTALEXTVOL_CELL As String = "I2"

' *** その他CREATE ***
Public Const DDLROLLBACK As String = "ロールバックCREATE文"
Public Const M_STARTROW As Integer = 1
Public Const M_DBNAME_ROW As Integer = 2
Public Const M_MAKE_ROW As Integer = 3
Public Const M_PREGEN_ROW As Integer = 5

'*** テーブルスペース一覧 ***
Public Const TSPACELIST As String = "テーブルスペース一覧"
Public Const W_STARTROW As Integer = 6
Public Const W_EXIST_COL As Integer = 1 'A
Public Const W_TBLSPACE_COL As Integer = 2 'B
Public Const W_TEMP_COL As Integer = 3 'C
Public Const W_DATFILE_COL As Integer = 4 'D
Public Const W_FILEVOL_COL As Integer = 5 'E
Public Const W_TBLSPACEVOL_COL As Integer = 6 'F
Public Const W_PCTFREE_COL As Integer = 7 'G
Public Const W_PCTUSED_COL As Integer = 8 'H
Public Const W_MINIMUM_COL As Integer = 9 'I
Public Const W_INITEXT_COL As Integer = 10 'J
Public Const W_NEXTEXT_COL As Integer = 11 'K
Public Const W_MINEXT_COL As Integer = 12 'L
Public Const W_MAXEXT_COL As Integer = 13 'M
Public Const W_MISC_COL As Integer = 14 'N
Public Const W_OBJGROUP_COL As Integer = 15 'O
Public Const W_PARTSEQ_COL As Integer = 16 'P
Public Const W_TBLIDX_COL As Integer = 17 'Q
Public Const W_TSGROUP_COL As Integer = 18 'R
Public Const W_TSSEQ_COL As Integer = 19 'S
Public Const W_SUBNAME_CELL As String = "B2"
Public Const W_DBNAME_CELL As String = "E2"
Public Const W_BLOCKSIZE_CELL As String = "F2"
Public Const W_TOTALVOL_CELL As String = "I2"
Public Const W_AUTHOR_CELL As String = "K2"
Public Const W_DATE_CELL As String = "M2"
Public Const W_VERSION_CELL As String = "G1"

' *** 表領域CREATE ***
Public Const DDLTBLSPACE As String = "表領域CREATE文"
Public Const Y_STARTROW As Integer = 1
Public Const Y_DBNAME_ROW As Integer = 2
Public Const Y_MAKE_ROW As Integer = 3
Public Const Y_PREGEN_ROW As Integer = 5

' *** ログ・制御ファイル ***
Public Const LOGCTLLIST As String = "ログ・制御ファイル一覧"
Public Const Q_STARTROW As Integer = 6
Public Const Q_SYSTEM_ROW As Integer = 7
Public Const Q_EXIST_COL As Integer = 1 'A
Public Const Q_FILEOBJ_COL As Integer = 2 'B
Public Const Q_OBJTYPE_COL As Integer = 3 'C
Public Const Q_LOGTHREAD_COL As Integer = 4 'D
Public Const Q_LOGGROUP_COL As Integer = 5 'E
Public Const Q_DATFILE_COL As Integer = 6 'F
Public Const Q_FILEVOL_COL As Integer = 7 'G
Public Const Q_SETUMEI_COL As Integer = 8 'H
Public Const Q_MISC_COL As Integer = 9 'I
Public Const Q_DBNAME_CELL As String = "F2"
Public Const Q_SYSTEM_OBJ As String = "SYSTEM表領域"

' *** DB_CREATE ***
Public Const DDLDATABASE As String = "DB_CREATE文"
Public Const O_STARTROW As Integer = 1
Public Const O_DBNAME_ROW As Integer = 2
Public Const O_MAKE_ROW As Integer = 3
Public Const O_PREGEN_ROW As Integer = 5

'****
Public Const NA_STR As String = "-"
Public Const SAME_STR As String = "〃"
Public Const R_TABLE As String = "TABLE"
Public Const R_TABLE_PARTITION As String = "TABLE_P"
Public Const R_INDEX As String = "INDEX"
Public Const R_INDEX_LOCAL As String = "INDEX_L"
'Public Const R_INDEX_LOCAL_PREFIX As String = "INDEX_LP"
'Public Const R_INDEX_LOCAL_NONPREFIX As String = "INDEX_LNP"
Public Const R_INDEX_GLOBAL As String = "INDEX_G"
Public Const R_PARTITION As String = "PARTITION"
Public Const R_PARTITION_TABLE As String = "PARTITION_T"
Public Const R_PARTITION_LOCAL As String = "PARTITION_L"
Public Const R_PARTITION_GLOBAL As String = "PARTITION_G"
Public Const R_LOCAL As String = "LOCAL"
Public Const R_GLOBAL As String = "GLOBAL"
Public Const R_TEMPORARY As String = "TEMPORARY"
Public Const R_TABLESPACE_SHARE As String = "$"
Public Const R_GLOBAL_PK_TS As String = "ts_pk_ptables"
Public Const V_TABLE As String = "TBL"
Public Const V_INDEX As String = "IDX"
Public Const V_ROLLBACK As String = "RBS"
Public Const V_TEMP As String = "TMP"
Public Const V_PARTITION As String = "PRT"
Public Const V_PARTITION_SOURCE As String = "PRS"
Public Const V_TABLESPACE As String = "TSP"
Public Const Q_SYSTEM As String = "SYS"
Public Const Q_REDOLOG As String = "LOG"
Public Const Q_CONTROL As String = "CTL"

'for WIN32API
'定数
Public Const MAX_PATH = 260
Public Const ERROR_INVALID_HANDLE = 6&
Public Const ERROR_NO_MORE_FILES = 18&
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10

'型
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

'関数
Public Declare Function GetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
'*
Public GV_srcBook As String
Public GV_interactive As String
Public GV_partSave As String
Public GV_ignore As String
Public GV_updated As Boolean

'*** ERwin連携 ***
Public Const E_STARTROW As Integer = 2
Public Const E_ATTR_COL As Integer = 1 'A
Public Const E_DEFINE_COL As Integer = 2 'B
Public Const E_INDEX_COL As Integer = 3 'C
Public Const E_PKEY_COL As Integer = 4 'D
Public Const E_ENTITY_COL As Integer = 5 'E
Public Const E_TABLE_COL As Integer = 6 'F
Public Const E_COLUMN_COL As Integer = 7 'G
Public Const E_TYPE_COL As Integer = 8 'H
Public Const E_NULL_COL As Integer = 9 'I

'*** テーブル件数一覧 ***
Public Const KENSULIST As String = "テーブル件数一覧"
Public Const K_STARTROW As Integer = 6
Public Const K_EXIST_COL As Integer = 1  'A
Public Const K_TBLNAME_COL As Integer = 2  'B
Public Const K_INITNUM_COL As Integer = 3  'C
Public Const K_MAXNUM_COL As Integer = 4  'D
Public Const K_ROWSIZE_COL As Integer = 5  'E
Public Const K_PARTITION_COL As Integer = 6  'F
Public Const K_TBLVOL_COL As Integer = 7  'G
Public Const K_INDEX_COL As Integer = 8  'H
Public Const K_IDXVOL_COL As Integer = 9  'I
Public Const K_SUMVOL_COL As Integer = 10  'J
Public Const K_MISC_COL As Integer = 11  'K
Public Const K_BOOK_COL As Integer = 12  'L
Public Const K_TABLEID_COL As Integer = 13  'M
Public Const K_TSTABLE_COL As Integer = 14  'N
Public Const K_TSINDEX_COL As Integer = 15  'O
Public Const K_TBLVOL_CELL As String = "D2"
Public Const K_IDXVOL_CELL As String = "F2"
Public Const K_SUMVOL_CELL As String = "H2"
Public Const K_BLOCKSIZE_CELL As String = "L2"
Public Const K_VERSION_CELL As String = "M1"
Public Const KENSULIST_WK As String = "ファイル一覧並べ替え"
Public Const KWK_DIR_ROW As Integer = 1
Public Const KWK_STARTROW As Integer = 3
Public Const KWK_ORDER_COL As Integer = 1   'A
Public Const KWK_PATH_COL As Integer = 2    'B
Public Const KWK_COMMENT_COL As Integer = 3 'C

'*** 分割比率 ***
Public Const BUNKATUHIRITU As String = "分割比率"
Public Const B_STARTROW As Integer = 6
Public Const B_EXIST_COL As Integer = 1 'A
Public Const B_TOTAL_COL As Integer = 2 'B
Public Const B_B1VAL_COL As Integer = 3 'C
Public Const B_B1DEF_COL As Integer = 4 'D
Public Const B_B1PCT_COL As Integer = 5 'E
Public Const B_B2VAL_COL As Integer = 6 'F
Public Const B_B2DEF_COL As Integer = 7 'G
Public Const B_B2PCT_COL As Integer = 8 'H
Public Const B_B3VAL_COL As Integer = 9 'I
Public Const B_B3DEF_COL As Integer = 10 'J
Public Const B_B3PCT_COL As Integer = 11 'K
Public Const B_BUNKATU_COL As Integer = 12 'L
Public Const B_B1ITEM_CELL As String = "D4"
Public Const B_B2ITEM_CELL As String = "G4"
Public Const B_B3ITEM_CELL As String = "J4"
Public Const B_BKIND_CELL As String = "I2"
Public Const B_BUNKATU_0 As String = "N/A"
Public Const B_BUNKATU_1 As String = "B"
Public Const B_BUNKATU_2 As String = "BB"
Public Const B_BUNKATU_3 As String = "BBB"
Public Const B_BUNKATU_E As String = "E"
Public Const B_BUNKATU_EB As String = "EB"
Public Const B_BUNKATU_EBB As String = "EBB"

'--- Add Start S.Iwanaga 2010/04/08
Public Const CONVERT_LIST_FILE As String = "name.xls"      '変換情報ファイル名
Public Const CONVERT_LIST_SHEET As String = "name"         '変換情報テーブルシート名
Public Const DEFINITION_NAME As String = "DIC"             '変換情報テーブル範囲の定義名
'--- Add End

'--- Add Start 2012/03/05 TFC
Public Const DDL_KIND_ALL As String = "0"
Public Const DDL_KIND_TABLE As String = "1"
Public Const DDL_KIND_INDEX As String = "2"
'--- Add Start 2012/03/05 TFC
