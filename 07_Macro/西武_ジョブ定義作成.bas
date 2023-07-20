Attribute VB_Name = "西武_ジョブ定義作成"
Option Explicit

Public Const CELL_JIKOU_FILE_NAME = "実行ファイル名"                   ' 【実行ファイル名】定義
Public Const CELL_JOB As String = "I5"                                 ' 【ジョブ】定義
Public Const NAME_JOB_ID As String = "ジョブID"                        ' 【ジョブID】定義
Public Const NAME_DB_CONDITIONAL As String = "DB削除条件"              ' 【削除条件】定義
Public Const NAME_JIKO_SYORI As String = "実行処理"                    ' 【実行処理】定義
Public Const NAME_KINOU_ID As String = "機能ID"                        ' 【機能ID】定義
Public Const NAME_KINOU_RENBAN As String = "機能内連番"                ' 【機能内連番】定義
Public Const NAME_KOUBAN As String = "項番"                            ' 【項番】定義
Public Const NAME_OUTPUT_PATH As String = "出力パス(絶対パス)"         ' 【出力パス(絶対パス)】定義
Public Const NAME_SYORI_SYUBETSU As String = "処理種別"                ' 【処理種別】定義
Public Const NAME_HULFT_SYUBETSU As String = "HULFT種別"               ' 【HULFT種別】定義
Public Const NAME_ACMS As String = "ACMS"                              ' 【ACMS】定義
Public Const OUTPUT_ADD_NAME As String = "Config"                      ' 【出力ファイル付加名】定義
Public Const OUTPUT_FILE_TYPE As String = ".sh"                        ' 【出力ファイル種類】定義
Public Const STR_BIN_BASH As String = "#!/bin/bash"                    ' 【SHファイルのヘッダ】定義
Public Const STR_ACMS_USER As String = "ACMS_USER_ID"                  ' 【ACMSユーザ [取引先略称]】定義
Public Const STR_ACMS_FILE As String = "ACMS_FILE_ID"                  ' 【ACMSファイル [ファイル略称]】定義
Public Const WS_CELL_SETTING_BUILD_FLG As String = "作成可フラグ"      ' シート【項目設定】の【作成可フラグ】セル定義
Public Const WS_CELL_SETTING_CONSTANT As String = "固定名"             ' シート【項目設定】の【固定名】セル定義
Public Const WS_CELL_SETTING_PRC_PATH As String = "実行パス"           ' シート【項目設定】の【実行パス】セル定義
Public Const WS_CELL_SETTING_SETSUZOKU_SAKI As String = "SFTP接続先"   ' シート【項目設定】の【SFTP接続先】セル定義
Public Const WS_CELL_SETTING_SYORI_SYUBETSU As String = "処理区分"     ' シート【項目設定】の【処理区分】セル定義
Public Const WS_NAME_SHEET_KOMOKU_SETTING As String = "項目設定"       ' シート【項目設定】の 定義


' メイン処理
' 戻り値   : なし
Public Function CreateShFile_Seibu()

    Dim DictShFile As Object: Set DictShFile = createDictionary                       ' 【シェル内容】リストを定義(Dictionary)
    Dim backFlg As Boolean                                                            ' 画面更新機能フラグを一時停止
    Dim cnt As Integer: cnt = 0                                                       ' ループ回数
    Dim curKey As Variant                                                             ' Dicのループ中キーを取得
    Dim folderPath As String: folderPath = getFolderPath(NAME_OUTPUT_PATH)            ' フォルダパスの定義
    Dim key As String                                                                 ' Dicのループ中キーを取得（String）

    Application.ScreenUpdating = False                                                ' 画面更新機能フラグを一時停止
    backFlg = Application.ScreenUpdating                                              ' 画面更新機能フラグをバックアップ
    Set DictShFile = getShText                                                        ' 出力リストを取得
    Application.ScreenUpdating = backFlg                                              ' 画面更新機能フラグを元に戻す

    For Each curKey In DictShFile                                                     ' 設定したDic内容にて、ループする
        key = curKey                                                                  ' DicのkeyはVariantタイプです。Stringに変換
        key = getFileName(key, OUTPUT_FILE_TYPE, OUTPUT_ADD_NAME)
        cnt = cnt + CreateFileWithoutBom(folderPath, key, DictShFile.item(curKey))
    Next

    ' 出力完了メッセージ
    MsgBox "シェルファイルの出力を完了しました。" & vbCrLf & "出力数 : " & DictShFile.count

     ' 作成するものが存在する場合、フォルダを開く
    If cnt > 0 Then
        Dim filepath As Range
        Set filepath = Range(searchCell(NAME_OUTPUT_PATH))
        Shell "C:\Windows\Explorer.exe " & Cells(filepath.Row + 1, filepath.Column), vbNormalFocus
    End If

End Function

' パラメタ : なし
' 戻り値   : 出力内容
Private Function getShText() As Object

    Dim DETAIL_JIKKO_SYORI As String:                                                                                                                      ' 実行処理セルの内容を取得
    Dim DictShFile As Object: Set DictShFile = createDictionary                                                                                            ' 【シェル内容】リストを定義(Dictionary)
    Dim ListHulftSyubetsu As Object: Set ListHulftSyubetsu = getGroupList(NAME_HULFT_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING, True)                         ' Hulft種別リストを取得
    Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getGroupList(WS_CELL_SETTING_SYORI_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING, False)             ' 処理種別リストを取得
    Dim cellJikkoSyori As Range: Set cellJikkoSyori = Range(searchCell(NAME_JIKO_SYORI))                                                                   ' 実行処理のセルを取得
    Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(NAME_SYORI_SYUBETSU))                                                         ' 【処理種別】セル
    Dim getTotalCase As Integer: getTotalCase = getCountCase(NAME_KOUBAN, 2)                                                                               ' ケース数の定義
    Dim jobList As Object: Set jobList = getVertivalListbyCnt(searchCell(NAME_JOB_ID), getTotalCase, 2)                                                                   ' ジョブリスト
    Dim curCell As Range                                                                                                                                   ' ループ中セルを取得（String）
    Dim curKey As Variant                                                                                                                                       ' Dicのループ中キーを取得（String）
    Dim shText As String                                                                                                                                   ' 【シェル内容】を定義

    ' ジョブリストでループする
    For Each curKey In jobList.keys
        Set curCell = Range(curKey)
        DETAIL_JIKKO_SYORI = Cells(curCell.Row, cellJikkoSyori.Column)
        ' ヘッダの設定
        shText = STR_BIN_BASH + vbLf
        shText = shText + vbLf + setKomokuComment(removeSpecCode(DETAIL_JIKKO_SYORI))
        ' 内容の設定
        shText = setShText(curCell.Row, shText, ListSyoriSyubetsu, ListHulftSyubetsu)

        ' Dicに代入
        If Len(shText) > 0 Then
            If DictShFile.Exists(jobList.item(curKey)) Then
                DictShFile.item(jobList.item(curKey)) = shText                                                    ' すでに存在する場合、内容を更新する
            Else
                DictShFile.Add jobList.item(curKey), shText                                                       ' 存在しない場合、追加する
            End If

            Dim cellJikkoFileName As Range: Set cellJikkoFileName = Range(searchCell(CELL_JIKOU_FILE_NAME))       ' 実行ファイル名称のセルを取得
            Dim jikoKbn As String: jikoKbn = Cells(curCell.Row, cellSyoriSyubetsu.Column)                         ' 対象Rowの実行種別を取得
            Dim prcPath As String: prcPath = ListSyoriSyubetsu.item(jikoKbn)(2)                                   ' 対象実行種別の定義パスを取得
            Cells(curCell.Row, cellJikkoFileName.Column) = prcPath                                                ' 実行ファイル名称を設定
            Cells(curCell.Row, cellJikkoFileName.Column + 1) = curCell.value                                      ' 実行パラメータ名称を設定
        End If
    Next

    Set getShText = DictShFile

End Function

' パラメタ : 現在ロー、現在出力内容、処理種別リスト、HULFT種別リスト
' 戻り値   : 出力内容
Private Function setShText(ByVal curRow As Integer, shText As String, ListSyoriSyubetsu As Object, ListHulftSyubetsu As Object) As String

    Dim cellNextCtgl As Range                                                                                                        ' 次ジョブカタログ
    Dim count As Integer                                                                                                             ' 項目ループ回数
    Dim countCtgy As Integer                                                                                                         ' カタログループ回数
    Dim startCol As Integer                                                                                                          ' 開始カラム
    Dim getJobCtgyContent() As String: getJobCtgyContent() = getTitleList(NAME_SYORI_SYUBETSU, getCountKomoku(NAME_SYORI_SYUBETSU))  ' ジョブカタログの定義
    Dim sizeJobCatagoryList As Integer: sizeJobCatagoryList = getArrayLength(getJobCtgyContent())                                    ' ジョブカタログリストの内容を長さ

    ' ジョブカタログリストの内容の長さでループする
    For countCtgy = 0 To sizeJobCatagoryList - 1
        ' 項目ループ回数を初期化
        count = 0
        ' 次ジョブカタログの取得
        Set cellNextCtgl = Nothing
        ' 次のカタログが存在する場合、取得する
        If countCtgy < sizeJobCatagoryList - 1 Then
            Set cellNextCtgl = Range(searchCell(getJobCtgyContent(countCtgy + 1)))
        End If

        ' 現ジョブカタログを取得
        Dim curJobCatagory As Range: Set curJobCatagory = Range(searchCell(getJobCtgyContent(countCtgy)))

        ' 開始セルの設定
        Dim startCell As Range: Set startCell = Cells(curRow, curJobCatagory.Column)

        ' 【処理種別】のカラムの、対象外の処理種別が入力された場合、作成しない
        If curJobCatagory.value = NAME_SYORI_SYUBETSU Then
            If ListSyoriSyubetsu.Exists(startCell.value) = False Then
                shText = ""
                Exit For
            ElseIf ListSyoriSyubetsu.item(startCell.value)(1) <> "〇" Then
                shText = ""
                Exit For
            End If
        End If

        ' 開始カラム
        startCol = startCell.Column
        shText = shText + vbLf
        ' 次ジョブカタログが存在する場合
        If Not (cellNextCtgl Is Nothing) Then
            ' 次ジョブカタログのカラムと同じまで、ループする
            Do While startCol + count <> cellNextCtgl.Column
                shText = shText + setShText02(startCol, count, curJobCatagory, startCell, ListSyoriSyubetsu, ListHulftSyubetsu)
                count = count + 1
            Loop
        Else
            ' 次のカラムが存在しないまで、ループする
            Do While Len(Cells(curJobCatagory.Row + 1, startCol + count).value) > 0
                shText = shText + setShText02(startCol, count, curJobCatagory, startCell, ListSyoriSyubetsu, ListHulftSyubetsu)
                count = count + 1
            Loop
        End If

        ' ACMS内容を手動で作成
        shText = addACMSExtendDetail(shText, curRow, curJobCatagory.value)

    Next

    If Len(shText) > 0 Then
        ' 定数内容の追加
        shText = addConstantText(shText, curRow)
    End If

    setShText = shText

End Function


' パラメタ : 現在出力内容、現在ロー
' 戻り値   : 出力内容（固定値）
Private Function addConstantText(ByVal shText As String, curRow As Integer) As String

    ' 連番セル
    Dim rnBCell As Range: Set rnBCell = Range(searchCell(NAME_KINOU_RENBAN)): Set rnBCell = Cells(curRow, rnBCell.Column)
    ' 機能セル
    Dim kinouId As Range: Set kinouId = Range(searchCell(NAME_KINOU_ID)): Set kinouId = Cells(curRow, kinouId.Column)
    ' 値
    Dim value As String: value = ""

    If Len(kinouId.Text) > 0 Then
        value = kinouId.Text
    End If

    ' 定数の内容(漢字)を代入
    shText = shText + vbLf + setKomokuComment("定数内容")
    If Len(value) > 0 Then
        shText = shText + setDetailByOneSet("プロセスID", "PROC_ID", value + padLeftString(rnBCell.value, "0", 3))
        shText = shText + setDetailByOneSet("ジョブID", "JOB_ID", value + padLeftString(rnBCell.value, "0", 4))
    Else
        shText = shText + setDetailByOneSet("プロセスID", "PROC_ID", "")
        shText = shText + setDetailByOneSet("ジョブID", "JOB_ID", "")
    End If

    ' 固定名リストを取得
    Dim listConstant As Object: Set listConstant = getGroupList(WS_CELL_SETTING_CONSTANT, WS_NAME_SHEET_KOMOKU_SETTING, True)
    Dim Constkeys As Variant
    ' Dicのループ中キーを取得（String）
    Dim Constkey As String
    Dim value2 As String
    For Each Constkeys In listConstant
        Constkey = Constkeys
        value = listConstant.item(Constkey)(0)
        value2 = listConstant.item(Constkey)(1)
        shText = shText + setDetailByOneSet(Constkeys, value, value2)
    Next

    addConstantText = shText

End Function

' パラメタ : 現在出力内容、セル内容
' 戻り値   : 出力内容（SFTPの追加内容）
Private Function setSFTPExtraDetail(ByVal shText As String, cellValue As String) As String

    Dim value As String: value = ""

    Dim getContent As Boolean: getContent = False

    'SFTP追加内容
    Dim sftpObject As Object
    If Len(cellValue) > 0 Then
        'Set sftpObject = getGroupListbySelectedValue(WS_CELL_SETTING_SETSUZOKU_SAKI, WS_NAME_SHEET_KOMOKU_SETTING, True, cellValue)
        Set sftpObject = getGroupListbySelectedValue(WS_CELL_SETTING_SETSUZOKU_SAKI, WS_NAME_SHEET_KOMOKU_SETTING, True, False, 0, cellValue)
        Dim key As Variant
        ' 一件しかない予定
        For Each key In sftpObject

            value = sftpObject.item(key)(1)
            shText = shText + setDetailByOneSet("SFTPホスト", "SFTP_HOST", value)
            value = sftpObject.item(key)(2)
            shText = shText + setDetailByOneSet("SFTPユーザー", "SFTP_USER", value)
            value = sftpObject.item(key)(3)
            shText = shText + setDetailByOneSet("SFTP秘密鍵パス", "SFTP_KEY_PATH", value)
            ' 設定済フラグ
            getContent = True
        Next
    End If

    If getContent = False Then
        shText = shText + setDetailByOneSet("SFTPホスト", "SFTP_HOST", "")
        shText = shText + setDetailByOneSet("SFTPユーザー", "SFTP_USER", "")
        shText = shText + setDetailByOneSet("SFTP秘密鍵パス", "SFTP_KEY_PATH", "")
    End If
    setSFTPExtraDetail = shText

End Function

' パラメタ : コメント内容
' 戻り値   : 項目コメントの作成
Private Function setKomokuComment(ByVal comment As String) As String

    setKomokuComment = padRightString("# *----" + comment, "-", 60) + vbLf

End Function

' パラメタ : タイトル内容
' 戻り値   : 項目タイトルの作成
Private Function setKomokuTitle(ByVal title As String) As String

    setKomokuTitle = "# " + title + vbLf

End Function

' パラメタ : コンテンツタイトル、内容
' 戻り値   : コンテンツ内容の作成
Private Function setKomokuDetail(ByVal title As String, value As String) As String

    setKomokuDetail = title + "=" + """" + value + """" + vbLf

End Function

' パラメタ : 項目（漢字）、項目（英字）、内容
' 戻り値   : コンテンツ内容（セット）の作成
Private Function setDetailByOneSet(ByVal komokuKanji As String, komoku As String, value As String) As String

    setDetailByOneSet = ""
    setDetailByOneSet = setDetailByOneSet + setKomokuTitle(komokuKanji)
    setDetailByOneSet = setDetailByOneSet + setKomokuDetail(komoku, value)

End Function

' パラメタ : 開始ロー、現在ループ回数、現在カテゴリ、開始セル、処理種別リスト、HULFT種別リスト
' 戻り値   : 出力内容
Private Function setShText02(ByVal startCol As Integer, _
                            count As Integer, _
                            curJobCatagory As Range, _
                            startCell As Range, _
                            ListSyoriSyubetsu As Object, _
                            ListHulftSyubetsu As Object _
                            ) As String
    setShText02 = ""
    Dim value As String
    ' 現在カラム =現ジョブカタログのカラムの場合
    If startCol + count = curJobCatagory.Column Then
        ' 現ジョブカタログの内容(漢字)を代入
        setShText02 = setShText02 + setKomokuComment(Replace(curJobCatagory.value, vbCrLf, ""))
    End If

    ' 現項目の説明（漢字）を代入
    setShText02 = setShText02 + setKomokuTitle(Replace(Cells(curJobCatagory.Row + 2, startCol + count).value, vbLf, ""))

    ' 現項目を代入
    Dim curTitle As String: curTitle = Cells(curJobCatagory.Row + 1, startCol + count)
    ' 現項目の内容を代入
    value = Replace(Cells(startCell.Row, startCol + count).value, vbCrLf, "")
    ' 【処理種別】のカラムの場合
    If curJobCatagory.value = NAME_SYORI_SYUBETSU Then
        value = ListSyoriSyubetsu.item(value)(0)
    ElseIf curTitle = "HULFT_TYPE" And Len(value) > 0 Then
        value = ListHulftSyubetsu.item(value)(0)
    ElseIf curTitle = STR_ACMS_USER And Len(value) > 0 Then
        ' AcmsユーザIDリストを取得
        value = getValueByKeyFromDictionary(STR_ACMS_USER, value, WS_NAME_SHEET_KOMOKU_SETTING)
    ElseIf curTitle = STR_ACMS_FILE And Len(value) > 0 Then
        ' AcmsファイルIDリストを取得
        value = getValueByKeyFromDictionary(STR_ACMS_FILE, value, WS_NAME_SHEET_KOMOKU_SETTING)
    End If

    setShText02 = setShText02 + setKomokuDetail(Replace(Cells(curJobCatagory.Row + 1, startCol + count).value, vbLf, ""), value)

    If curTitle = "SFTP_DEST" Then
        'SFTP処理区分の追加
        'setShText02 = addSFTPSyoriKbn(setShText02, startCell)
        'SFTP追加情報の追加
        setShText02 = setSFTPExtraDetail(setShText02, value)
    End If

End Function


' パラメタ : 現在内容、開始セル
' 戻り値   : 出力内容
' 未使用 20191007
Private Function addSFTPSyoriKbn(ByVal setShText02 As String, startCell As Range)

        ' 現項目の説明（漢字）を代入
        setShText02 = setShText02 + setKomokuTitle("SFTP処理区分")
        Dim SFTP_KBNObject As Object
        Set SFTP_KBNObject = getGroupList("SFTP処理区分", WS_NAME_SHEET_KOMOKU_SETTING, True)

        Dim colKbn As Range
        Set colKbn = Range(searchCell(NAME_SYORI_SYUBETSU))

        Dim selectedKbn As Range
        Set selectedKbn = Range(Cells(startCell.Row, colKbn.Column).Address)

        Dim valueSFTP As String
        If Len(selectedKbn.value) > 0 And SFTP_KBNObject.Exists(selectedKbn.value) = True Then
            valueSFTP = SFTP_KBNObject.item(selectedKbn.value)(0)
        Else
            valueSFTP = ""
        End If
        addSFTPSyoriKbn = setShText02 + setKomokuDetail("SFTP_KBN", valueSFTP)

End Function

' パラメタ : 現在内容、現在ロー、現在ジョブカテゴリ
' 戻り値   : 出力内容（ACMS 内容）
Private Function addACMSExtendDetail(ByVal shText As String, curRow As Integer, jobTitle As String) As String

    If jobTitle = NAME_ACMS Then

        Dim acmsUserRange As Range: Set acmsUserRange = Range(searchCell(STR_ACMS_USER))
        Dim acmsFileRange As Range: Set acmsFileRange = Range(searchCell(STR_ACMS_FILE))

        Dim strCombine As String
        If Cells(curRow, acmsUserRange.Column) <> "" Then
            strCombine = ""
            strCombine = strCombine + getValueByKeyFromDictionary(STR_ACMS_USER, Cells(curRow, acmsUserRange.Column).value, WS_NAME_SHEET_KOMOKU_SETTING) ' AcmsユーザIDリストを取得
            strCombine = strCombine + "_"
            strCombine = strCombine + getValueByKeyFromDictionary(STR_ACMS_FILE, Cells(curRow, acmsFileRange.Column).value, WS_NAME_SHEET_KOMOKU_SETTING) ' AcmsユーザIDリストを取得
            shText = shText + setDetailByOneSet("ACMSアプリケーション [取引先略称]_[ファイル略称]", "ACMS_APL_ID", strCombine)
        Else
            shText = shText + setDetailByOneSet("ACMSアプリケーション [取引先略称]_[ファイル略称]", "ACMS_APL_ID", "")
        End If

    End If

    addACMSExtendDetail = shText
End Function





