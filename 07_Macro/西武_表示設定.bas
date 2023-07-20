Attribute VB_Name = "西武_表示設定"
' 処理　   : 処理種別のリスト関数を作成
Public Function setListForSyoriKbn()
    Const WS_CELL_SETTING_SYORI_SYUBETSU As String = "処理区分"     ' シート【項目設定】の【処理区分】セル定義
    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "項目設定"       ' シート【項目設定】の 定義
    Const CELL_SYORI_SYUBETSU As String = "処理種別"                ' 【処理種別】定義
    Const KOUBAN As String = "項番"                                 ' 【項番】定義

    If isCellExist(CELL_SYORI_SYUBETSU) = True Then

        ' 処理種別リストを取得
        Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getListContent(WS_CELL_SETTING_SYORI_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING)
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(CELL_SYORI_SYUBETSU))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellSyoriSyubetsu.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellSyoriSyubetsu.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim item As Variant

        For Each item In ListSyoriSyubetsu.Items
            If Len(str) = 0 Then
                str = item
            Else
                str = str + "," + item
            End If
        Next

        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
    End If
End Function

' 処理　   : SFTPのリスト関数を作成
Public Function setListForSFTPKbn()
    Const WS_CELL_SETTING_SFTP_SYUBETSU As String = "SFTP処理区分"      ' シート【項目設定】の【SFTP処理区分】セル定義
    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "項目設定"           ' シート【項目設定】の 定義
    Const CELL_SFTP_SYUBETSU As String = "SFTP処理区分"                 ' 【SFTP処理区分別】定義
    Const KOUBAN As String = "項番"                                     ' 【項番】定義

    If isCellExist(CELL_SFTP_SYUBETSU) = True Then
        ' 処理種別リストを取得
        Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getListDictionary(WS_CELL_SETTING_SFTP_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING)
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(CELL_SFTP_SYUBETSU))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellSyoriSyubetsu.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellSyoriSyubetsu.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim key As Variant

        For Each key In ListSyoriSyubetsu.keys
            If Len(str) = 0 Then
                str = key
            Else
                str = str + "," + key
            End If
        Next

        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
    End If
End Function

' 処理　   : SFTP接続先のドロップリストを作成
Public Function isSFTPDestCell(ByVal Target As Range) As Integer
    ' 戻り値
    isSFTPDestCell = 9

    Const CELL_SFTP_SETSUZOKU_SAKI As String = "SFTP接続先"                 ' 【SFTP接続先】定義
    Const CELL_SFTP_SYUBETSU As String = "処理区分"                         ' 【SFTP処理区分別】定義
    Const WS_CELL_SETTING_SFTP_SYUBETSU As String = "SFTP処理区分"          ' シート【項目設定】の【SFTP処理区分】セル定義
    Const WS_CELL_SFTP_KEY As String = "SFTPキー"                           ' シート【項目設定】の【SFTPキー】セル定義
    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "項目設定"               ' シート【項目設定】の 定義


    If isCellExist(CELL_SFTP_SETSUZOKU_SAKI) = True Then
        ' 接続先のセルを取得
        Dim cellSetzuZokuSaki As Range: Set cellSetzuZokuSaki = Range(searchCell(CELL_SFTP_SETSUZOKU_SAKI))
    
        ' 選択カラム = 接続先のセルのカラムの場合
        If Target.Column = cellSetzuZokuSaki.Column Then
            ' SFTP区分セルの取得
            Dim cellSFTPSyubetsu As Range: Set cellSFTPSyubetsu = Range(searchCell(CELL_SFTP_SYUBETSU))
            ' SFTP区分セルのDictionaryを取得
            Dim getKey As Object: Set getKey = getListDictionary(WS_CELL_SETTING_SFTP_SYUBETSU, WS_NAME_SHEET_KOMOKU_SETTING)
            ' キーの定義
            Dim key As String: key = Cells(Target.Row, cellSFTPSyubetsu.Column)
            ' ドロップリストの取得
            Dim getDropList As Object: Set getDropList = getGroupListbySelectedValue(WS_CELL_SFTP_KEY, WS_NAME_SHEET_KOMOKU_SETTING, True, False)
            ' ドロップリストの取得可能の場合
    
            Dim backFlg As Boolean: backFlg = Application.ScreenUpdating
            Application.ScreenUpdating = False
            If getDropList.count > 0 Then
    
                Dim str As String
                Dim keys As Variant
    
                For Each keys In getDropList.keys
                    key = keys
                    If Len(str) = 0 Then
                        str = getDropList.item(key)(1)
                    Else
                        str = str + "," + getDropList.item(key)(1)
                    End If
                Next
                Target.Validation.Delete
                Target.Validation.Add Type:=xlValidateList, Formula1:=str
            Else
                Target.Validation.Delete
    
            End If
    
            Application.ScreenUpdating = backFlg
        Else
    
        End If
    End If

    isSFTPDestCell = 0
End Function

' 処理　   : 処理種別のリスト関数を作成
Public Function setListForEmptyFileFlg()

    Const CELL_KARA_FILE_SAKUSEI As String = "空ファイル作成"                ' 【空ファイル作成】定義
    Const KOUBAN As String = "項番"                                          ' 【項番】定義

    If isCellExist(CELL_KARA_FILE_SAKUSEI) = True Then
        
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellKaraFileSakusei As Range: Set cellKaraFileSakusei = Range(searchCell(CELL_KARA_FILE_SAKUSEI))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellKaraFileSakusei.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellKaraFileSakusei.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim item As Variant
        str = "YES,NO"
        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
    End If
End Function

' 処理　   : HULFT種別リスト関数を作成
Public Function setListForHulftType()

    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "項目設定"           ' シート【項目設定】の 定義
    Const CELL_HULFT_TYPE As String = "HULFT種別"                ' 【HULFT種別】定義
    Const KOUBAN As String = "項番"                                     ' 【項番】定義

    If isCellExist(CELL_HULFT_TYPE) = True Then
        
        ' 処理種別リストを取得
        Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getListDictionary(CELL_HULFT_TYPE, WS_NAME_SHEET_KOMOKU_SETTING)
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(CELL_HULFT_TYPE))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellSyoriSyubetsu.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellSyoriSyubetsu.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim key As Variant

        For Each key In ListSyoriSyubetsu.keys
            If Len(str) = 0 Then
                str = key
            Else
                str = str + "," + key
            End If
        Next

        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
    End If
End Function

' 処理　   : HULFT種別リスト関数を作成
Public Function setListForAcmsType(ByVal strInput As String)

    Const WS_NAME_SHEET_KOMOKU_SETTING As String = "項目設定"           ' シート【項目設定】の 定義
    Const KOUBAN As String = "項番"                                     ' 【項番】定義

    If isCellExist(strInput) = True Then
        
        ' 処理種別リストを取得
        Dim ListSyoriSyubetsu As Object: Set ListSyoriSyubetsu = getListDictionary(strInput, WS_NAME_SHEET_KOMOKU_SETTING)
        Dim cntTotal As Integer: cntTotal = getCountCase(KOUBAN, 2)
        Dim cellKoban As Range: Set cellKoban = Range(searchCell(KOUBAN))
        Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = Range(searchCell(strInput))

        Dim cell As Range
        Dim backFlg As Boolean

        Dim strRangeStart As String: strRangeStart = Cells(cellKoban.Row + 2, cellSyoriSyubetsu.Column).Address
        Dim strRangeEnd As String: strRangeEnd = Cells(cellKoban.Row + 2 + cntTotal - 1, cellSyoriSyubetsu.Column).Address
        Dim rangeArea As String: rangeArea = strRangeStart + " : " + strRangeEnd

        backFlg = Application.ScreenUpdating
        Application.ScreenUpdating = False

        Dim str As String: str = ""
        Set cell = Range(rangeArea)
        Dim key As Variant

        For Each key In ListSyoriSyubetsu.keys
            If Len(str) = 0 Then
                str = key
            Else
                str = str + "," + key
            End If
        Next

        cell.Validation.Delete
        cell.Validation.Add Type:=xlValidateList, Formula1:=str
        Application.ScreenUpdating = backFlg
        
    End If
End Function




