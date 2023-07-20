Option Explicit

' DB系の定義
Dim HOST                                            ' DBホスト
Dim PORT                                            ' DBポート
Dim DATABASE                                        ' DB名称
Dim USER                                            ' DBユーザ名
Dim PASSWORD                                        ' DBパスワード
Dim DB_JAR                                          ' DB実行jar
Dim ADM_USER                                        ' DB実行jar
Dim ADM_PASSWORD                                    ' DB実行jar

' 内容の定義
Dim ABATOR_CONFIG_BATCH                             ' Abator系定義(バッチ)
Dim ABATOR_CONFIG_API                               ' Abator系定義(API)
Dim ABATOR_CONFIG_TEXT                              ' Abator系定義(内容)
Dim SQL_MAP_BASE_CONFIG_TBL_API                     ' SQLMAP系定義(API)
Dim SQL_MAP_BASE_CONFIG_TBL_BATCH                   ' SQLMAP系定義(BATCH)
Dim SQL_MAP_BASE_CONFIG_TXT                         ' SQLMAP系定義(内容)
Dim APPLICATION_CONTEXT                             ' APPLICATION系定義(タイトル)
Dim APPLICATION_CONTEXT_CONTENT                     ' APPLICATION系定義(内容)

' Seqの定義内容
Dim SEQ_BATCH_BASEDAO                               ' SeqバッチBaseDao定義
Dim SEQ_BATCH_SQLMAP                                ' SeqバッチSqlMap定義
Dim SEQ_ONLINE_BASEDAO                              ' SeqリアルBaseDao定義
Dim SEQ_ONLINE_BASEDAOIMP                           ' SeqリアルBaseDaoImp定義
Dim SEQ_ONLINE_SQLMAP                               ' SeqリアルSqlMap定義

Dim LIST_SEQ_NAME_CAMEL                             ' SEQ_キャメルケース文字列
Dim LIST_SEQ_NAME_CLASS_NAME                        ' SEQ_クラス名
Dim LIST_SEQ_NAME_ID_ORACLE                         ' SEQ_SQLID(Oracle用)
Dim LIST_SEQ_NAME_ID_POSTGRES                       ' SEQ_SQLID(PostgreSQL用)
Dim LIST_SEQ_NAME_ID_UCASE                          ' SEQ_シーケンスID（大文字）
Dim LIST_SEQ_NAME_ID_LCASE                          ' SEQ_シーケンスID（小文字）
Dim LIST_SEQ_NAME_ID_LCASE_SQLMAP                   ' SEQ_シーケンスID（小文字）
Dim LIST_SEQ_NAME_ID_LCASE_SQLMAP_BASESQLMAP        ' SEQ_シーケンスID（小文字）
Dim LIST_SEQ_NAME_COMMENT                           ' SEQ_シーケンスコメント
Dim ARRAY_SEQ_DETAIL()                              ' SEQ_配列
Dim LIST_ARRAY_SEQ_DETAIL                           ' SEQ配列を格納するリスト

' パスの定義
Dim CUR_PATH                                        ' 現在パス

' INPUT系
Dim INPUT_PATH                                      ' INPUTパス
Dim INPUT_TBL_DLL_PATH                              ' INPUTパス(テーブルDLL)
Dim INPUT_INDEX_DLL_PATH                            ' INPUTパス(インデックスDLL)
Dim INPUT_SEQ_DLL_PATH                              ' INPUTパス(SEQDLL)

' WORK系
Dim WORK_PATH                                       ' WORKパス
Dim WORK_SQL_PATH                                   ' WORKパス(SQL)
Dim WORK_SQL_OUTPUT_PATH                            ' WORKパス(OUPUT)
Dim WORK_BAT_PATH                                   ' WORKパス(bat)
Dim WORK_VBS_PATH                                   ' WORKパス(vbs)
Dim WORK_OUTPUT_XML_PATH                            ' WORKパス(xml)
Dim WORK_OUTPUT_TXT_PATH                            ' WORKパス(txt)
Dim WORK_SQL_GET_TABLE_LIST                         ' WORKパス(テーブル名取得SQL)
Dim WORK_SQL_GET_SEQ_LIST                           ' WORKパス(Seq名取得SQL)
Dim WORK_TEMP_REAL_BASEDAO_PATH                     ' RealCommonBaseDao
Dim WORK_TEMP_REAL_SQLMAP_PATH                      ' RealCommonSqlMap
Dim WORK_TEMP_BATCH_BASEDAO_PATH                    ' BatchCommonBaseDao
Dim WORK_TEMP_BATCH_SQLMAP_PATH                     ' BatchCommonSqlMap
Dim WORK_TEMP_BUILD_PATH                            ' Buildファイル
Dim WORK_TEMP_LIB_PATH                              ' libフォルダ

' OUTPUT系
Dim OUTPUT_PATH                                     ' OUTPUTパス
Dim OUTPUT_SQLMAP_PATH_API                          ' OUTPUTパス(作成先API_SQLMAP)
Dim OUTPUT_SQLMAP_PATH_BATCH                        ' OUTPUTパス(作成先BAT_SQLMAP)
Dim OUTPUT_APP_CONT_PATH                            ' OUTPUTパス(作成先APPLICATION)
Dim OUTPUT_DAO_PATH                                 ' OUTPUTパス(作成済DAO)
Dim OUTPUT_SQLMAP_PATH                              ' OUTPUTパス(Batch_SQLMAP)
Dim OUTPUT_REAL_DAO_PATH                            ' OUTPUTパス(作成済RealDAO)
Dim OUTPUT_REAL_SQLMAP_PATH                         ' OUTPUTパス(Online_SQLMAP)
Dim OUTPUT_PISCOMMON_PATH                           ' OUTPUTパス(PisCommon)
Dim OUTPUT_PISCOMMON_CORE_PATH                      ' OUTPUTパス(PisCommon_Src_Core)

' 共通系
Dim OBJECT_FOR_ALL                                  ' 共通のオブジェクト
Dim WORKBOOK                                        ' 共通のワークブック
Dim MESSAGE                                         ' 共通のメッセージ変数
Dim objProgressMsg                                  ' Makes the object a Public object (Critical!)
' ======================処理開始======================

showProcessBar (0)

' 初期パス設定
Call SetPath

showProcessBar (5)

' 処理前確認メッセージ ,【OK:1】の場合のみ実行
MESSAGE = ""
MESSAGE = MESSAGE & "処理開始します。よろしいですか？" & vbCrLf
MESSAGE = MESSAGE & vbCrLf
MESSAGE = MESSAGE & "20191201 更新内容 : 初期作成" & vbCrLf
MESSAGE = MESSAGE & "20191205 更新内容 : Seqの対応可能" & vbCrLf
MESSAGE = MESSAGE & "20191211 更新内容 : Abatorの自動実行" & vbCrLf
MESSAGE = MESSAGE & "20191211 更新内容 : IEを依存しない" & vbCrLf
if showMsgOKCancel (MESSAGE,"確認") = 1 then

' 異常の場合、続行する
On Error Resume Next

    'テーブルDLLが存在しない場合
    if execGetFileCountBatch(INPUT_TBL_DLL_PATH) = 0 and execGetFileCountBatch(INPUT_SEQ_DLL_PATH) = 0 then
        showMsg "登録可能TBLが存在しません!!!" & vbCrLf & "今回の処理は終了します。"

    else

        ' 処理パスを作成
        Call execCreatePathBatch (CUR_PATH)
        ' 進捗状態の設定(10%)
        showProcessBar(10)

        ' 【設定内容.xlsx】から初期内容を取得
        Call ReadInitialFile
        ' 進捗状態の設定(20%)
        showProcessBar(20)

        ' Db接続確認
        Call execCheckDb

        ' DBの削除
        Call execRemoveDb
        ' 進捗状態の設定(30%)
        showProcessBar(30)

        ' DBの登録
        ' テーブル
        Call execRegistDb (INPUT_TBL_DLL_PATH)
        ' インデックス
        Call execRegistDb (INPUT_INDEX_DLL_PATH)
        ' SEQ (文字化けしますが、問題ありません)
        Call execRegistDb (INPUT_SEQ_DLL_PATH)

        ' 進捗状態の設定(40%)
        showProcessBar(40)

        ' バッチ実行（DBの全TBL名を取得）
        Call execGetTableNameListBatch
        ' 進捗状態の設定(50%)
        showProcessBar(50)

        ' バッチ実行（DBの全Seq名を取得）
        Call execGetSeqNameListBatch
        Call setSeqListName
        showProcessBar(55)

        ' Abator初期設定内容を取得して、置き換えする
        ' API
        ABATOR_CONFIG_BATCH = ReplaceAstarConfWithUcase(ABATOR_CONFIG_BATCH,ABATOR_CONFIG_TEXT, "@REWRITEHERE_BATCH@")
        ' Batch
        ABATOR_CONFIG_API   = ReplaceAstarConfWithUcase(ABATOR_CONFIG_API, ABATOR_CONFIG_TEXT, "@REWRITEHERE_REAL@")
        ' 進捗状態の設定(60%)
        showProcessBar(60)

        ' AbatorConfig系を作成
        Call createConfigFile
        ' 進捗状態の設定(70%)
        showProcessBar(70)

        ' Abator.jarの実行
        Call execAbator4JFK
        ' Abator.jarの実行結果を移動する
        Call execMoveFolder(WORK_BAT_PATH & "java ",OUTPUT_PATH & "java\")
        ' 進捗状態の設定(80%)
        showProcessBar(80)

        ' sqlMapの作成(置き換えする)
        ' Onlineの対応
        Dim SQLMAP_FOR_API: SQLMAP_FOR_API = SQL_MAP_BASE_CONFIG_TXT
        Call ReplaceAstarConf(WORKBOOK, SQLMAP_FOR_API, SQL_MAP_BASE_CONFIG_TBL_API, "@SQLMAP@")
        Call createFile(OUTPUT_SQLMAP_PATH_API, SQLMAP_FOR_API)

        ' Batchの対応
        Dim SQLMAP_FOR_BATCH: SQLMAP_FOR_BATCH = SQL_MAP_BASE_CONFIG_TXT
        Call ReplaceAstarConf(WORKBOOK, SQLMAP_FOR_BATCH, SQL_MAP_BASE_CONFIG_TBL_BATCH, "@SQLMAP@")
        Call createFile(OUTPUT_SQLMAP_PATH_BATCH, SQLMAP_FOR_BATCH)
        ' 進捗状態の設定(90%)
        showProcessBar(90)

        ' Dao名取得
        ' applicationContext情報作成(置き換えする)
        ' Seqの対応
        Dim APPLICATION_CONTEXT_CONTENT_SEQ : APPLICATION_CONTEXT_CONTENT_SEQ = APPLICATION_CONTEXT_CONTENT
        APPLICATION_CONTEXT_CONTENT_SEQ = replaceAstarBySeqDao(APPLICATION_CONTEXT_CONTENT_SEQ)
        APPLICATION_CONTEXT = Replace(APPLICATION_CONTEXT, "@APPCONTENT_SEQ@", APPLICATION_CONTEXT_CONTENT_SEQ)
        ' Tblの対応
        APPLICATION_CONTEXT_CONTENT = replaceAstarByBaseDao(APPLICATION_CONTEXT_CONTENT)
        APPLICATION_CONTEXT = Replace(APPLICATION_CONTEXT, "@APPCONTENT@", APPLICATION_CONTEXT_CONTENT)
        ' applicationContextファイルの作成
        Call createFile(OUTPUT_APP_CONT_PATH, APPLICATION_CONTEXT)
        ' Seqファイルの作成
        Call createSeqAllFile

        ' CommonBaseDaoをコピー
        Call execCopyFolder (WORK_TEMP_REAL_BASEDAO_PATH, OUTPUT_REAL_DAO_PATH)
        Call execCopyFolder (WORK_TEMP_BATCH_BASEDAO_PATH, OUTPUT_DAO_PATH)
        Call execCopyFolder (WORK_TEMP_REAL_SQLMAP_PATH, OUTPUT_REAL_SQLMAP_PATH)
        Call execCopyFolder (WORK_TEMP_BATCH_SQLMAP_PATH, OUTPUT_SQLMAP_PATH)

        ' PisCommonへコピー
        Call execCopyFolder (OUTPUT_PATH & "xml\", OUTPUT_PISCOMMON_PATH)
        Call execCopyFolder (OUTPUT_PATH & "java\", OUTPUT_PISCOMMON_CORE_PATH)
        Call execCopyFolder (WORK_TEMP_BUILD_PATH, OUTPUT_PISCOMMON_PATH)
        Call execCopyFolder (WORK_TEMP_LIB_PATH, OUTPUT_PISCOMMON_PATH)

        ' 進捗状態の設定(100%)
        showProcessBar(100)
    End if

    ' 異常発生の場合
    if err <> 0 then
        MESSAGE = ""
        MESSAGE = MESSAGE + "処理途中に作成失敗しました。"
        MESSAGE = MESSAGE + vbCrLf
        MESSAGE = MESSAGE + "再実行してください!!!"
        showMsg MESSAGE
        ' 質問したくない場合、コメントアウト可能
        if showMsg ("作成済の物を削除しますか？",vbOKCancel,"確認") = 1 then
            err = 0
            Call execDelFileorFolder (OUTPUT_PATH)
            Call execDelFileorFolder (WORK_BAT_PATH & "java")
        End if
    else
        ' 成功の場合
        MESSAGE = ""
        MESSAGE = MESSAGE + "おめでとう！！！"
        MESSAGE = MESSAGE + vbCrLf
        MESSAGE = MESSAGE + "作成成功しました！！！"
        showMsg MESSAGE
    End if
else
    showMsg "処理終了します！！！"
End if

' 処理終了する
WScript.Quit 0

' ======================処理終了======================

' 初期パスの設定
Sub SetPath()

    Set OBJECT_FOR_ALL            = CreateObject("WScript.Shell")
    CUR_PATH                      = OBJECT_FOR_ALL.CurrentDirectory & "\"

    INPUT_PATH                    = CUR_PATH & "INPUT\"
    INPUT_TBL_DLL_PATH            = CUR_PATH & "INPUT\DB_TABLE_DLL\"
    INPUT_INDEX_DLL_PATH          = CUR_PATH & "INPUT\DB_INDEX_DLL\"
    INPUT_SEQ_DLL_PATH            = CUR_PATH & "INPUT\DB_SEQ_DLL\"

    OUTPUT_PATH                   = CUR_PATH & "OUTPUT\"
    OUTPUT_APP_CONT_PATH          = OUTPUT_PATH & "xml\conf\jar\spring\applicationContext-online-dao.xml"
    OUTPUT_DAO_PATH               = OUTPUT_PATH & "java\jp\hitachisoft\jfk\batch\common\db\dao\"
    OUTPUT_SQLMAP_PATH            = OUTPUT_PATH & "java\jp\hitachisoft\jfk\batch\common\db\sqlmap\"
    OUTPUT_SQLMAP_PATH_API        = OUTPUT_PATH & "xml\conf\jar\spring\sqlMapBaseConfig.xml"
    OUTPUT_SQLMAP_PATH_BATCH      = OUTPUT_PATH & "xml\src\core\java\jp\hitachisoft\jfk\batch\common\db\sqlmap\BatchBaseSqlMapConfig.xml"
    OUTPUT_REAL_DAO_PATH          = OUTPUT_PATH & "java\jp\hitachisoft\jfk\online\common\db\dao\"
    OUTPUT_REAL_SQLMAP_PATH       = OUTPUT_PATH & "java\jp\hitachisoft\jfk\online\common\db\sqlmap\"
    OUTPUT_PISCOMMON_PATH         = OUTPUT_PATH & "PisCommonDao\"
    OUTPUT_PISCOMMON_CORE_PATH    = OUTPUT_PATH & "PisCommonDao\src\core\java\"

    WORK_PATH                     = CUR_PATH & "WORK\"
    WORK_BAT_PATH                 = WORK_PATH & "bat\"
    WORK_SQL_PATH                 = WORK_PATH & "sql\"
    WORK_VBS_PATH                 = WORK_PATH & "vbs\"
    WORK_OUTPUT_TXT_PATH          = WORK_PATH & "output\txt\"
    WORK_OUTPUT_XML_PATH          = WORK_PATH & "output\xml\"
    WORK_SQL_OUTPUT_PATH          = WORK_PATH & "output\sql\"
    WORK_SQL_GET_TABLE_LIST       = WORK_OUTPUT_TXT_PATH & "TABLE_NAME_LIST.txt"
    WORK_SQL_GET_SEQ_LIST         = WORK_OUTPUT_TXT_PATH & "SEQ_NAME_LIST.txt"

    WORK_TEMP_REAL_BASEDAO_PATH          = WORK_PATH & "temp\source\online\basedao\"
    WORK_TEMP_BATCH_BASEDAO_PATH         = WORK_PATH & "temp\source\batch\basedao\"
    WORK_TEMP_REAL_SQLMAP_PATH           = WORK_PATH & "temp\source\online\sqlmap\"
    WORK_TEMP_BATCH_SQLMAP_PATH          = WORK_PATH & "temp\source\batch\sqlmap\"
    WORK_TEMP_BUILD_PATH                 = WORK_PATH & "temp\source\build\"
    WORK_TEMP_LIB_PATH                   = WORK_PATH & "temp\source\lib"

    ' 解放する
    Set OBJECT_FOR_ALL = Nothing

End Sub

' 初期ファイルを読み込み
Sub ReadInitialFile()

' 異常なしの場合
if err = 0 then

    ' ワークブックを読み取り
    Dim initialXlsxPath : initialXlsxPath = CUR_PATH & "設定内容.xlsx"
    Dim objExcel : Set objExcel = CreateObject("Excel.Application")

    ' ワークブックの取得
    Set WORKBOOK = objExcel.Workbooks.Open(initialXlsxPath)
    ' 初期データの設定
    Call ReadInitialData(WORKBOOK, "初期設定", 3, 2)
    ' ABATOR系の設定
    ABATOR_CONFIG_BATCH = ReadInitialDataWithReplace(WORKBOOK, "abatorConfigBatch", 3, 2)
    ABATOR_CONFIG_API = ReadInitialDataWithReplace(WORKBOOK, "abatorConfigReal", 2, 2)
    ABATOR_CONFIG_TEXT = ReadInitialDataWithReplace(WORKBOOK, "ConfigText", 2, 2)
    ' SQLMAP系の設定
    SQL_MAP_BASE_CONFIG_TXT = ReadInitialDataWithLoopCnt(WORKBOOK, "sqlMapBaseConfig_001", 18)
    SQL_MAP_BASE_CONFIG_TBL_API = ReadSheetOneCellOnly(WORKBOOK, "sqlMapBaseConfig_API", 2, 2)
    SQL_MAP_BASE_CONFIG_TBL_BATCH = ReadSheetOneCellOnly(WORKBOOK, "sqlMapBaseConfig_Batch", 2, 2)
    ' APPLICATION系の設定
    APPLICATION_CONTEXT = ReadInitialDataWithLoopCnt(WORKBOOK, "applicationContext", 54)
    APPLICATION_CONTEXT_CONTENT = ReadInitialDataWithLoopCnt(WORKBOOK, "applicationContext_Content", 22)

    ' Seq内容の取得
    ' SeqバッチBaseDao定義
    SEQ_BATCH_BASEDAO = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_BatchBaseDAO", 1)
    ' SeqバッチSqlMap定義
    SEQ_BATCH_SQLMAP = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_BatchSqlMap", 1)
    ' SeqリアルBaseDao定義
    SEQ_ONLINE_BASEDAO = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_OnlineBaseDao", 1)
    ' SeqリアルBaseDaoImp定義
    SEQ_ONLINE_BASEDAOIMP = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_OnlineBaseDaoImp", 1)
    ' SeqリアルSqlMap定義
    SEQ_ONLINE_SQLMAP = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_OnlineSqlMap", 1)

    ' ワークブックを解放
    objExcel.Quit
End if

End Sub

' 初期データを読み込み
' 引数1  : ワークブック（オブジェクト）
' 引数2  : シート名称   (文字列)
' 引数3  : 開始列
' 引数4  : 開始行
' 戻り値 : なし
Sub ReadInitialData(objWorkbook, sheetName, offsetRow, offsetCol)

' 異常なしの場合
if err = 0 then

    ' シートの取得
    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    ' 開始列の設定
    Dim intRow: intRow = offsetRow
    ' 行が空白までループして、取得する
    Do Until objWorkSheet.Cells(intRow, offsetCol).Value = ""
        Call setInitialData (objWorkSheet.Cells(intRow, offsetCol), objWorkSheet.Cells(intRow, offsetCol + 1))
        intRow = intRow + 1
    Loop

End if

End Sub

' 初期データを代入する
' 引数1  : キー値
' 引数2  : Value値
' 戻り値 : なし
Sub setInitialData(iKey, iValue)

    If (iKey = "ホスト") Then

        HOST = iValue

    ElseIf (iKey = "ポート") Then

        PORT = iValue

    ElseIf (iKey = "DB名") Then

        DATABASE = iValue

    ElseIf (iKey = "ユーザー") Then

        USER = iValue

    ElseIf (iKey = "パスワード") Then

        PASSWORD = iValue

    ElseIf (iKey = "PostgresJar") Then

        DB_JAR = iValue

    ElseIf (iKey = "上位ユーザー") Then

        ADM_USER = iValue

    ElseIf (iKey = "上位パスワード") Then

        ADM_PASSWORD = iValue

    End If

End Sub

' 処理内容を読み込んで、置き換えする
' 引数1  : ワークブック（オブジェクト）
' 引数2  : シート名称   (文字列)
' 引数3  : 開始列
' 引数4  : 開始行
' 戻り値 : 置き換え後文字列
Function ReadInitialDataWithReplace(objWorkbook, sheetName, offsetRow, offsetCol)

' 異常なしの場合
if err = 0 then

    ' シートの取得
    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    ' 開始列の設定
    Dim intRow: intRow = offsetRow
    ' 戻り値の宣言
    Dim ConfStr : ConfStr = ""
    ' 行が空白までループして、取得する
    Do Until objWorkSheet.Cells(intRow, offsetCol).Value = ""
        ConfStr = ConfStr & objWorkSheet.Cells(intRow, offsetCol) & vbCrLf
        intRow = intRow + 1
    Loop

    ' 取得した内容を上書きする
    ReadInitialDataWithReplace = replaceStr(ConfStr)

End if

End Function

' 文字を置き換え
' 引数1  : 入力内容
' 戻り値 : 置き換え後文字列
Function replaceStr(inStr)

' 異常なしの場合
if err = 0 then

    replaceStr = inStr
    If (HOST <> "") Then
        replaceStr = Replace(replaceStr, "@HOST@", HOST)
    End If
    If (PORT <> "") Then
        replaceStr = Replace(replaceStr, "@PORT@", PORT)
    End If
    If (DATABASE <> "") Then
        replaceStr = Replace(replaceStr, "@DBNAME@", DATABASE)
    End If
    If (USER <> "") Then
        replaceStr = Replace(replaceStr, "@USERID@", USER)
    End If
    If (PASSWORD <> "") Then
        replaceStr = Replace(replaceStr, "@PASSWORD@", PASSWORD)
    End If
    If (DB_JAR <> "") Then
        replaceStr = Replace(replaceStr, "@POSTGRESDLLPATH@", DB_JAR)
    End If

End if

End Function

' X回ループで内容を取得
' 引数1  : ワークブック（オブジェクト）
' 引数2  : シート名称   (文字列)
' 引数3  : ループ回数
' 戻り値 : 取得した文字列
Function ReadInitialDataWithLoopCnt(objWorkbook, sheetName, loopCnt)

' 異常なしの場合
if err = 0 then

    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    Dim ConfStr: ConfStr = ""
    Dim cnt
    ' 内容が2行目から。【設定内容.xlsx】にて確認
    For cnt = 2 To loopCnt+1
        ConfStr = ConfStr & objWorkSheet.Cells(cnt, 2) & vbCrLf
    Next
    ' 読み込んだ内容
    ReadInitialDataWithLoopCnt = ConfStr
End if

End Function

' X回ループで内容を取得
' 引数1  : ワークブック（オブジェクト）
' 引数2  : シート名称   (文字列)
' 引数3  : ループ回数
' 戻り値 : 取得した文字列
Function ReadInitialDataWithLoopCntReplaceVbLf(objWorkbook, sheetName, loopCnt)

' 異常なしの場合
if err = 0 then

    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    Dim ConfStr: ConfStr = ""
    Dim cnt
    ' 内容が2行目から。【設定内容.xlsx】にて確認
    For cnt = 2 To loopCnt+1
        ConfStr = ConfStr & objWorkSheet.Cells(cnt, 2) & vbCrLf
    Next

    ConfStr = Replace(ConfStr,vbCrLf,"@XXXX@")
    ConfStr = Replace(ConfStr,vbCr,"@XXXX@")
    ConfStr = Replace(ConfStr,vbLf,"@XXXX@")
    ConfStr = Replace(ConfStr,"@XXXX@",vbCrLf)
    ' 読み込んだ内容
    ReadInitialDataWithLoopCntReplaceVbLf = ConfStr
End if

End Function

' セルを読み込む
' 引数1  : ワークブック（オブジェクト）
' 引数2  : シート名称   (文字列)
' 引数3  : 列番号
' 引数4  : 行番号
' 戻り値 : 取得した文字列
Function ReadSheetOneCellOnly(objWorkbook, sheetName, xRow, yCol)

' 異常なしの場合
if err = 0 then

    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    Dim ConfStr: ConfStr = ConfStr & objWorkSheet.Cells(xRow, yCol) & vbCrLf
    ' 読み込んだ内容
    ReadSheetOneCellOnly = ConfStr

End if

End Function

' BaseDaoのリストを取得し、対象文字【****】を置き換え
' 引数1  : 入力文字列
' 戻り値 : 置き換えした文字列
Function replaceAstarByBaseDao(inpurStr)

' 異常なしの場合
if err = 0 then
    ' BaseDaoのリストを取得
    Dim arryFileName: Set arryFileName = getBaseDaoList

    Dim oStr : oStr = ""
    if Not (arryFileName Is Nothing ) then
        Dim filename
        For Each filename In arryFileName
            ' 置き換えする
            oStr = oStr & Replace(inpurStr, "****", filename)
        Next
        ' 戻り値
        replaceAstarByBaseDao = oStr
    End if
End if

End Function

' BaseDaoのリストを取得し、対象文字【****】を置き換え
' 引数1  : 入力文字列
' 戻り値 : 置き換えした文字列
Function replaceAstarBySeqDao(inpurStr)

' 異常なしの場合
if err = 0 then
    ' BaseDaoのリストを取得
    Dim arryFileName: Set arryFileName = LIST_SEQ_NAME_CLASS_NAME

    Dim oStr : oStr = ""
    Dim filename
    For Each filename In arryFileName
        ' 置き換えする
        oStr = oStr & Replace(inpurStr, "****", filename)
    Next
    ' 戻り値
    replaceAstarBySeqDao = oStr
End if

End Function

' テーブルリストを取得し、対象文字を置き換え
' 引数1  : 入力文字列
' 引数2  : 置き換えしたい内容
' 引数3  : 置き換えキー
' 戻り値 : 置き換えした文字列（大文字）
Function ReplaceAstarConfWithUcase(outputStr, repSrc, repKey)

' 異常なしの場合
if err = 0 then

    Set OBJECT_FOR_ALL = WScript.CreateObject("Scripting.FileSystemObject")
    Dim contentStr: contentStr = ""

    Dim lineStr

    ' 読み込みファイルの指定
    Dim inputFile_db: Set inputFile_db = OBJECT_FOR_ALL.OpenTextFile(WORK_SQL_GET_TABLE_LIST, 1, False, 0)
    ' 読み込みファイルから1行ずつ読み込み、書き出しファイルに書き出すのを最終行まで繰り返す
    Do Until inputFile_db.AtEndOfStream
        lineStr = Trim(inputFile_db.ReadLine)
        If (Len(lineStr) > 0) Then
            lineStr = Replace(repSrc, "****", UCase(lineStr))
        End If
        contentStr = contentStr & "    " & lineStr
    Loop

    ' 置き換えする
    outputStr = Replace(outputStr, repKey, contentStr)

    Set OBJECT_FOR_ALL = Nothing
    ReplaceAstarConfWithUcase = outputStr

End if

End Function

' テーブルリストを取得し、対象文字を置き換え
' 引数1  : 入力文字列
' 引数2  : 置き換えしたい内容
' 引数3  : 置き換えキー
' 戻り値 : 置き換えした文字列
Sub ReplaceAstarConf(objWorkbook, outputStr, repSrc, repKey)

' 異常なしの場合
if err = 0 then

    Set OBJECT_FOR_ALL = WScript.CreateObject("Scripting.FileSystemObject")
    Dim contentStrSql: contentStrSql = ""
    Dim contentStr: contentStr = ""
    ' Seqの内容を代入する
    Dim lineStr
    Dim item
    For Each item In LIST_SEQ_NAME_ID_LCASE_SQLMAP
        lineStr = item
        lineStr = Replace(repSrc, "****", lineStr)
        contentStrSql = contentStrSql & "    " & lineStr
    Next
    outputStr = Replace(outputStr, "@SQLMAPSEQ@", contentStrSql)

    if (OBJECT_FOR_ALL.FileExists(WORK_SQL_GET_TABLE_LIST)) then
        ' Tblの内容を代入する
        Dim inputFile_db: Set inputFile_db = OBJECT_FOR_ALL.OpenTextFile(WORK_SQL_GET_TABLE_LIST, 1, False, 0)

        ' 読み込みファイルから1行ずつ読み込み、書き出しファイルに書き出すのを最終行まで繰り返す
        Do Until inputFile_db.AtEndOfStream
            lineStr = Trim(inputFile_db.ReadLine)
            If (Len(lineStr) > 0) Then
                lineStr = Replace(repSrc, "****", lineStr)
            End If
            contentStr = contentStr & "    " & lineStr
        Loop
    End if
    outputStr = Replace(outputStr, repKey, contentStr)

    Dim sqlmapComm : sqlmapComm = Replace(repSrc, "****", "common")
    outputStr = Replace(outputStr, "@SQLMAPCOMM@", sqlmapComm)
    Set OBJECT_FOR_ALL = Nothing

End if

End Sub

' Abatorファイルの作成
Sub createConfigFile()

' 異常なしの場合
if err = 0 then

    Dim outputFileName(1)
    outputFileName(0) = WORK_OUTPUT_XML_PATH & "abatorConfigBatch.xml"
    outputFileName(1) = WORK_OUTPUT_XML_PATH & "abatorConfigReal.xml"

    Dim outputStr:  outputStr = ""

    Set OBJECT_FOR_ALL = WScript.CreateObject("Scripting.FileSystemObject")
    Dim cnt

    For cnt = LBound(outputFileName) To UBound(outputFileName)

        ' 書き出しファイルの指定 (今回は新規作成する)
        Dim outputFile: Set outputFile = OBJECT_FOR_ALL.OpenTextFile(outputFileName(cnt), 2, True)
        If (cnt = 0) Then
            outputFile.WriteLine ABATOR_CONFIG_BATCH
        Else
            outputFile.WriteLine ABATOR_CONFIG_API
        End If
        ' バッファを Flush してファイルを閉じる
        outputFile.Close

    Next
    Set OBJECT_FOR_ALL = Nothing

End if

End Sub

' ファイルの作成
Sub createFile(path, contents)

' 異常なしの場合
if err = 0 then

    Dim outputStr:  outputStr = ""
    Set OBJECT_FOR_ALL = WScript.CreateObject("Scripting.FileSystemObject")

    createPath(OBJECT_FOR_ALL.GetParentFolderName(path))
    Dim outputFile: Set outputFile = OBJECT_FOR_ALL.OpenTextFile(path, 2, True)
    outputFile.WriteLine contents
    ' バッファを Flush してファイルを閉じる
    outputFile.Close
    Set OBJECT_FOR_ALL = Nothing
End if

End Sub

' ファイルの作成
Sub createFile_sjis(path, contents)

' 異常なしの場合
if err = 0 then
    ' 書き出しファイルの指定 (今回は新規作成する)
    Set OBJECT_FOR_ALL = WScript.CreateObject("ADODB.Stream")
    OBJECT_FOR_ALL.Type = 2
    OBJECT_FOR_ALL.Charset = "Shift-JIS"
    OBJECT_FOR_ALL.Open
    OBJECT_FOR_ALL.WriteText contents
    OBJECT_FOR_ALL.SaveToFile path, 1
    ' バッファを Flush してファイルを閉じる
    OBJECT_FOR_ALL.Close
    Set OBJECT_FOR_ALL = Nothing
End if

End Sub

' バッチ実行(ファイル・フォルダ削除)
' 引数1  : 削除対象パス
Sub execDelFileorFolder(path)

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "path_delete.bat "
    cmd = cmd & path
    execBatch cmd

End Sub

' バッチ実行(パス作成)
' 引数1  : 現在パス
Sub execCreatePathBatch(path)

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "path_createPath.bat "
    cmd = cmd & path

    execBatch cmd
end if

End Sub

' バッチ実行(ファイル数の取得)
' 引数1  : 現在パス
Function execGetFileCountBatch(path)

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "file_getCount.bat "
    cmd = cmd & path

    execGetFileCountBatch = execBatchWithResponce (cmd)

End Function

' バッチ実行(DB確認)
' 引数1  : DLLパス
Sub execCheckDb

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "db_checkDatabase.bat "
    cmd = cmd & CUR_PATH & " "
    cmd = cmd & HOST & " "
    cmd = cmd & PORT & " "
    cmd = cmd & DATABASE & " "
    cmd = cmd & USER & " "
    cmd = cmd & PASSWORD

    execBatch cmd

    if err <> 0 then
        showMsg "データベースに接続できません！！！" & vbCrLf & "今回の処理は終了します。"
    End if
End if

End Sub

' バッチ実行(DB登録)
' 引数1  : DLLパス
Sub execRegistDb(path)

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "db_runSqlFile.bat "
    cmd = cmd & CUR_PATH & " "
    cmd = cmd & HOST & " "
    cmd = cmd & PORT & " "
    cmd = cmd & DATABASE & " "
    cmd = cmd & USER & " "
    cmd = cmd & PASSWORD & " "
    cmd = cmd & path

    execBatch cmd
End if

End Sub

' バッチ実行(DB登録)_SJIS
' 引数1  : DLLパス
Sub execRegistDb_Sjis(path)

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "db_runSqlFile_Sjis.bat "
    cmd = cmd & CUR_PATH & " "
    cmd = cmd & HOST & " "
    cmd = cmd & PORT & " "
    cmd = cmd & DATABASE & " "
    cmd = cmd & USER & " "
    cmd = cmd & PASSWORD & " "
    cmd = cmd & path

    execBatch cmd
End if

End Sub

' バッチ実行(DB登録)
Sub execRemoveDb()

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "db_remove.bat "
    cmd = cmd & CUR_PATH & " "
    cmd = cmd & HOST & " "
    cmd = cmd & PORT & " "
    cmd = cmd & DATABASE & " "
    cmd = cmd & USER & " "
    cmd = cmd & PASSWORD & " "
    cmd = cmd & WORK_SQL_PATH & " "
    cmd = cmd & WORK_SQL_OUTPUT_PATH & " "
    cmd = cmd & ADM_USER & " "
    cmd = cmd & ADM_PASSWORD

    execBatch cmd
End if

End Sub

' バッチ実行(DBの全TBL名を取得)
Sub execGetTableNameListBatch()

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "db_getDBSelectList.bat "
    cmd = cmd & CUR_PATH & " "
    cmd = cmd & HOST & " "
    cmd = cmd & PORT & " "
    cmd = cmd & DATABASE & " "
    cmd = cmd & USER & " "
    cmd = cmd & PASSWORD & " "
    cmd = cmd & WORK_SQL_PATH & "getTableNameQuery.sql "
    cmd = cmd & WORK_SQL_GET_TABLE_LIST

    execBatch cmd
End if

End Sub

' バッチ実行(DBの全Seq名を取得)
Sub execGetSeqNameListBatch()

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "db_getDBSelectList.bat "
    cmd = cmd & CUR_PATH & " "
    cmd = cmd & HOST & " "
    cmd = cmd & PORT & " "
    cmd = cmd & DATABASE & " "
    cmd = cmd & USER & " "
    cmd = cmd & PASSWORD & " "
    cmd = cmd & WORK_SQL_PATH & "getSeqNameQuery.sql "
    cmd = cmd & WORK_SQL_GET_SEQ_LIST

    execBatch cmd
End if

End Sub

' バッチ実行(既存Abator4JFK)
Sub execAbator4JFK()

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "java_abator4JFK.bat "
    cmd = cmd & WORK_BAT_PATH & "abator4JFK.jar "
    cmd = cmd & WORK_OUTPUT_XML_PATH & "abatorConfigReal.xml "
    cmd = cmd & WORK_OUTPUT_XML_PATH & "abatorConfigBatch.xml "

    execBatch cmd
End if

End Sub

' バッチ実行(フォルダ移動)
' 引数1  : 移動元
' 引数2  : 移動先
Sub execMoveFolder(pathSrc, pathDest)

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "path_move.bat "
    cmd = cmd & pathSrc & " "
    cmd = cmd & pathDest

    execBatch cmd

End if

End Sub

' バッチ実行(フォルダコピー)
' 引数1  : 移動元
' 引数2  : 移動先
Sub execCopyFolder(pathSrc, pathDest)

' 異常なしの場合
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "path_copy.bat "
    cmd = cmd & pathSrc & " "
    cmd = cmd & pathDest

    execBatch cmd

End if

End Sub

' バッチ実行
' 引数1  : コマンド
Function execBatch(cmd)

' 異常なしの場合
if err = 0 then

    ' WshShellオブジェクトを作成する
    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")

    ' batファイルを実行する
    err = WshShell.Run (cmd, 1, True)
    ' オブジェクトを開放する
    Set WshShell = Nothing
End if

End Function

' バッチ実行（隠すタイプ）
' 引数1  : コマンド
Function execBatchWithResponce(cmd)

    ' WshShellオブジェクトを作成する
    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")
    ' batファイルを実行する
    execBatchWithResponce = WshShell.Run (cmd, 1, True)
    ' オブジェクトを開放する
    Set WshShell = Nothing
End Function

' BaseDaoのリストを取得
' 戻り値  : BaseDaoの名称リスト
Function getBaseDaoList()

' 異常なしの場合
if err = 0 then

    Set OBJECT_FOR_ALL = CreateObject("Scripting.FileSystemObject")
    If (OBJECT_FOR_ALL.FolderExists(OUTPUT_DAO_PATH)) then
        Dim folder:Set folder = OBJECT_FOR_ALL.getFolder(OUTPUT_DAO_PATH)
        Dim ArrayList: Set ArrayList = CreateObject("System.Collections.ArrayList")
        ' ファイル一覧
        Dim file
        For Each file In folder.Files
            ArrayList.Add (OBJECT_FOR_ALL.getbasename(file.Name))
        Next
        Set getBaseDaoList = ArrayList
    Else
        Set getBaseDaoList = Nothing
    End if
    Set OBJECT_FOR_ALL = Nothing

End if

End Function

Function ProperCase(sText)
'*** Converts text to proper case e.g.  ***'
'*** surname = Surname                  ***'
'*** o'connor = O'Connor                ***'
    Dim a, iLen, bSpace, tmpX, tmpFull
    iLen = Len(sText)
    For a = 1 To iLen
    If a <> 1 Then 'just to make sure 1st character is upper and the rest lower'
        If bSpace = True Then
            tmpX = UCase(mid(sText,a,1))
            bSpace = False
        Else
        tmpX=LCase(mid(sText,a,1))
            If tmpX = "_" Or tmpX = "'" Then bSpace = True
        End if
    Else
        tmpX = UCase(mid(sText,a,1))
    End if
    tmpFull = tmpFull & tmpX
    Next
    ProperCase = tmpFull

End Function

' SEQ各名称の設定
Sub setSeqListName()
' 異常なしの場合
if err = 0 then
    Set LIST_SEQ_NAME_CAMEL = CreateObject("System.Collections.ArrayList")
    Set LIST_SEQ_NAME_CLASS_NAME = CreateObject("System.Collections.ArrayList")
    Set LIST_SEQ_NAME_ID_ORACLE = CreateObject("System.Collections.ArrayList")
    Set LIST_SEQ_NAME_ID_POSTGRES = CreateObject("System.Collections.ArrayList")
    Set LIST_SEQ_NAME_ID_UCASE = CreateObject("System.Collections.ArrayList")
    Set LIST_SEQ_NAME_ID_LCASE = CreateObject("System.Collections.ArrayList")
    Set LIST_SEQ_NAME_ID_LCASE_SQLMAP = CreateObject("System.Collections.ArrayList")
    Set LIST_SEQ_NAME_ID_LCASE_SQLMAP_BASESQLMAP = CreateObject("System.Collections.ArrayList")
    Set LIST_SEQ_NAME_COMMENT = CreateObject("System.Collections.ArrayList")

    ReDim LIST_SEQ_DETAIL(6)
    Set LIST_ARRAY_SEQ_DETAIL = CreateObject("System.Collections.ArrayList")

    Dim inputFile_db: Set inputFile_db = CreateObject("ADODB.Stream")
    inputFile_db.Type = 2    ' 1：バイナリ・2：テキスト
    inputFile_db.Charset = "UTF-8"    ' 文字コード指定
    inputFile_db.Open    ' Stream オブジェクトを開く
    inputFile_db.LoadFromFile WORK_SQL_GET_SEQ_LIST    ' ファイルを読み込む

    ' 読み込みファイルから1行ずつ読み込み、書き出しファイルに書き出すのを最終行まで繰り返す
    Do Until inputFile_db.EOS

        Dim lineStr : lineStr = Trim(inputFile_db.ReadText(-2))
        Dim str_EN
        Dim str_JP
        If (Len(lineStr) > 0) Then
            Dim nameAry : nameAry = Split(lineStr,",")
            str_EN = nameAry(0)
            str_JP = nameAry(1)
            ' SEQ_キャメルケース文字列
            LIST_SEQ_NAME_CAMEL.Add (Replace(ProperCase(str_EN),"_",""))
            ' SEQ_クラス名 = 【@CamelBaseDao@】
            LIST_SEQ_NAME_CLASS_NAME.Add (Replace(ProperCase(str_EN),"_","") & "BaseDao")
            ' SEQ_SQLID(Oracle用) = 【@SeqOracle@)】
            LIST_SEQ_NAME_ID_ORACLE.Add (UCase(str_EN) & ".nextvalue1")
            ' SEQ_SQLID(PostgreSQL用) = 【@SeqPostgres@】
            LIST_SEQ_NAME_ID_POSTGRES.Add (UCase(str_EN) & ".nextvalue")
            ' SEQ_シーケンスID（大文字）【@SeqUpperId@】
            LIST_SEQ_NAME_ID_UCASE.Add (UCase(str_EN))
            ' SEQ_シーケンスID（小文字）
            LIST_SEQ_NAME_ID_LCASE.Add (LCase(str_EN))
            ' SEQ_シーケンスID（小文字）
            LIST_SEQ_NAME_ID_LCASE_SQLMAP_BASESQLMAP.Add ( LCase(str_EN) &  "_BaseSqlMap")
            ' SEQ_シーケンスID（小文字）
            LIST_SEQ_NAME_ID_LCASE_SQLMAP.Add ( LCase(str_EN) )
            ' SEQ_シーケンスコメント【@SeqComment@】
            LIST_SEQ_NAME_COMMENT.Add (str_JP)

            LIST_SEQ_DETAIL(0) = (Replace(ProperCase(str_EN),"_",""))               ' SEQ_キャメルケース文字列
            LIST_SEQ_DETAIL(1) = (Replace(ProperCase(str_EN),"_","") & "BaseDao")   ' SEQ_クラス名
            LIST_SEQ_DETAIL(2) = (str_JP)                                           ' SEQ_シーケンスコメント
            LIST_SEQ_DETAIL(3) = (UCase(str_EN) & ".nextvalue1")                    ' SEQ_SQLID(Oracle用)
            LIST_SEQ_DETAIL(4) = (UCase(str_EN) & ".nextvalue")                     ' SEQ_SQLID(PostgreSQL用)
            LIST_SEQ_DETAIL(5) = (UCase(str_EN))                                    ' SEQ_シーケンスID（大文字）
            LIST_SEQ_DETAIL(6) = (LCase(str_EN))                                    ' SEQ_シーケンスID（小文字）

            LIST_ARRAY_SEQ_DETAIL.Add Join(LIST_SEQ_DETAIL,",")

        End If

    Loop
End if

End Sub

' Seqのファイルを作成
Sub createSeqAllFile()
' 異常なしの場合
if err = 0 then
    If LIST_ARRAY_SEQ_DETAIL.Count > 0 then
        Dim fileName : fileName = ""
        For Each fileName In LIST_ARRAY_SEQ_DETAIL
            Dim seqCamelName : seqCamelName = Split(fileName,",")(0)
            Dim seqCamelBaseDao : seqCamelBaseDao = Split(fileName,",")(1)
            Dim seqComment : seqComment = Split(fileName,",")(2)
            Dim seqOracle : seqOracle = Split(fileName,",")(3)
            Dim seqPostgres : seqPostgres = Split(fileName,",")(4)
            Dim seqUpper : seqUpper = Split(fileName,",")(5)
            Dim seqLower : seqLower = Split(fileName,",")(6)

            Dim t_SEQ_BATCH_BASEDAO : t_SEQ_BATCH_BASEDAO = SEQ_BATCH_BASEDAO
            Dim t_SEQ_BATCH_SQLMAP : t_SEQ_BATCH_SQLMAP = SEQ_BATCH_SQLMAP
            Dim t_SEQ_ONLINE_BASEDAO : t_SEQ_ONLINE_BASEDAO = SEQ_ONLINE_BASEDAO
            Dim t_SEQ_ONLINE_BASEDAOIMP : t_SEQ_ONLINE_BASEDAOIMP = SEQ_ONLINE_BASEDAOIMP
            Dim t_SEQ_ONLINE_SQLMAP : t_SEQ_ONLINE_SQLMAP = SEQ_ONLINE_SQLMAP

            if Len(seqComment) = 0 then
                seqComment = seqCamelName
            end if

            t_SEQ_BATCH_BASEDAO =  Replace(t_SEQ_BATCH_BASEDAO, "@CamelBaseDao@", seqCamelBaseDao)
            t_SEQ_BATCH_BASEDAO =  Replace(t_SEQ_BATCH_BASEDAO, "@SeqOracle@", seqOracle)
            t_SEQ_BATCH_BASEDAO =  Replace(t_SEQ_BATCH_BASEDAO, "@SeqPostgres@", seqPostgres)
            t_SEQ_BATCH_BASEDAO =  Replace(t_SEQ_BATCH_BASEDAO, "@SeqUpperId@", seqUpper)
            t_SEQ_BATCH_BASEDAO =  Replace(t_SEQ_BATCH_BASEDAO, "@SeqComment@", seqComment)
            Call createFile(OUTPUT_DAO_PATH & seqCamelBaseDao & ".java", t_SEQ_BATCH_BASEDAO)

            t_SEQ_BATCH_SQLMAP =  Replace(t_SEQ_BATCH_SQLMAP, "@CamelBaseDao@", seqCamelBaseDao)
            t_SEQ_BATCH_SQLMAP =  Replace(t_SEQ_BATCH_SQLMAP, "@SeqOracle@", seqOracle)
            t_SEQ_BATCH_SQLMAP =  Replace(t_SEQ_BATCH_SQLMAP, "@SeqPostgres@", seqPostgres)
            t_SEQ_BATCH_SQLMAP =  Replace(t_SEQ_BATCH_SQLMAP, "@SeqUpperId@", seqUpper)
            t_SEQ_BATCH_SQLMAP =  Replace(t_SEQ_BATCH_SQLMAP, "@SeqComment@", seqComment)
            Call createFile(OUTPUT_SQLMAP_PATH & seqLower & "_BaseSqlMap.xml", t_SEQ_BATCH_SQLMAP)

            t_SEQ_ONLINE_BASEDAO =  Replace(t_SEQ_ONLINE_BASEDAO, "@CamelBaseDao@", seqCamelBaseDao)
            t_SEQ_ONLINE_BASEDAO =  Replace(t_SEQ_ONLINE_BASEDAO, "@SeqOracle@", seqOracle)
            t_SEQ_ONLINE_BASEDAO =  Replace(t_SEQ_ONLINE_BASEDAO, "@SeqPostgres@", seqPostgres)
            t_SEQ_ONLINE_BASEDAO =  Replace(t_SEQ_ONLINE_BASEDAO, "@SeqUpperId@", seqUpper)
            t_SEQ_ONLINE_BASEDAO =  Replace(t_SEQ_ONLINE_BASEDAO, "@SeqComment@", seqComment)
            Call createFile(OUTPUT_REAL_DAO_PATH & seqCamelBaseDao & ".java", t_SEQ_ONLINE_BASEDAO)

            t_SEQ_ONLINE_BASEDAOIMP =  Replace(t_SEQ_ONLINE_BASEDAOIMP, "@CamelBaseDao@", seqCamelBaseDao)
            t_SEQ_ONLINE_BASEDAOIMP =  Replace(t_SEQ_ONLINE_BASEDAOIMP, "@SeqOracle@", seqOracle)
            t_SEQ_ONLINE_BASEDAOIMP =  Replace(t_SEQ_ONLINE_BASEDAOIMP, "@SeqPostgres@", seqPostgres)
            t_SEQ_ONLINE_BASEDAOIMP =  Replace(t_SEQ_ONLINE_BASEDAOIMP, "@SeqUpperId@", seqUpper)
            t_SEQ_ONLINE_BASEDAOIMP =  Replace(t_SEQ_ONLINE_BASEDAOIMP, "@SeqComment@", seqComment)
            Call createFile(OUTPUT_REAL_DAO_PATH & seqCamelBaseDao & "Impl.java", t_SEQ_ONLINE_BASEDAOIMP)

            t_SEQ_ONLINE_SQLMAP =  Replace(t_SEQ_ONLINE_SQLMAP, "@CamelBaseDao@", seqCamelBaseDao)
            t_SEQ_ONLINE_SQLMAP =  Replace(t_SEQ_ONLINE_SQLMAP, "@SeqOracle@", seqOracle)
            t_SEQ_ONLINE_SQLMAP =  Replace(t_SEQ_ONLINE_SQLMAP, "@SeqPostgres@", seqPostgres)
            t_SEQ_ONLINE_SQLMAP =  Replace(t_SEQ_ONLINE_SQLMAP, "@SeqUpperId@", seqUpper)
            t_SEQ_ONLINE_SQLMAP =  Replace(t_SEQ_ONLINE_SQLMAP, "@SeqComment@", seqComment)
            Call createFile(OUTPUT_REAL_SQLMAP_PATH & seqLower & "_BaseSqlMap.xml", t_SEQ_ONLINE_SQLMAP)
        Next
    End If
End If

End Sub

' フォルダの作成（親フォルドも作成対象）
' パラメタ : 作成するパス
' 戻り値   : なし
Function createPath(intPath)
    if(intPath <> "") then

        Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
        ' 親フォルダの取得
        Dim parentPath : parentPath = objFso.GetParentFolderName(intPath)
        ' 対象親フォルダの確認
        if parentPath <> "" and objFso.FolderExists(parentPath) = false then
            ' 親フォルダの作成(無限ループ的な感じ)
            createPath(parentPath)
        end if

        ' 対象フォルダの確認
        if objFso.FolderExists(intPath) = false then
            ' 対象フォルダの作成
            objFso.CreateFolder(intPath)
        end if
        ' 後始末
        Set objFso = Nothing
    end if

end function

' OKCANCELメッセージBox
Function showMsgOKCancel( strMsg, strTitle)

    ProgressMsg "", "実行中。。。"
    showMsgOKCancel = MsgBox (strMsg, vbOKCancel , strTitle)

End function

' メッセージBox
Function showMsg( strMsg)

    ProgressMsg "", "実行中。。。"
    MsgBox strMsg

End function

' 進捗メッセージBox
Function showProcessBar(intPercentage)

    ProgressMsg "", "実行中。。。"
    Const SOLID_BLOCK_CHARACTER = "■"
    Const EMPTY_BLOCK_CHARACTER = "□"
    Const COUNT_BAR = 30
    Dim progress : progress= Round(( intPercentage / 100) * COUNT_BAR)
    Dim cnt
    Dim setBar : setBar = ""
    For cnt = 1 To COUNT_BAR
        if (cnt <= progress )then
            setBar = setBar + SOLID_BLOCK_CHARACTER
        else
            setBar = setBar + EMPTY_BLOCK_CHARACTER
        end if
    Next

    Dim msg
    msg = setBar
    ProgressMsg msg, "実行中。。。" & intPercentage & "%"

End function

' 並列メッセージ
Function ProgressMsg( strMessage, strWindowTitle )
' Written by Denis St-Pierre
' Displays a progress message box that the originating script can kill in both 2k and XP
' If StrMessage is blank, take down previous progress message box
' Using 4096 in Msgbox below makes the progress message float on top of things
' CAVEAT: You must have   Dim ObjProgressMsg   at the top of your script for this to work as described

Dim wshShell,strTEMP,objFSO,strTempVBS,objTempMessage
    Set wshShell = CreateObject( "WScript.Shell" )
    strTEMP = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    If strMessage = "" Then
        ' Disable Error Checking in case objProgressMsg doesn't exists yet
        On Error Resume Next
        ' Kill ProgressMsg
        objProgressMsg.Terminate( )
        ' Re-enable Error Checking
        On Error Goto 0
        Exit Function
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strTempVBS = strTEMP + "\" & "Message.vbs"     'Control File for reboot

    ' Create Message.vbs, True=overwrite
    Set objTempMessage = objFSO.CreateTextFile( strTempVBS, True )
    objTempMessage.WriteLine( "MsgBox""" & strMessage & """, " & 4096 & ", """ & strWindowTitle & """" )
    objTempMessage.Close

    ' Disable Error Checking in case objProgressMsg doesn't exists yet
    On Error Resume Next
    ' Kills the Previous ProgressMsg
    objProgressMsg.Terminate( )
    ' Re-enable Error Checking
    On Error Goto 0

    ' Trigger objProgressMsg and keep an object on it
    Set objProgressMsg = WshShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS)
    Set wshShell = Nothing
    Set objFSO   = Nothing
End Function