Attribute VB_Name = "ComonMethod"

Public Sub CallMacro()
    ' 定義作成処理
    CreateShFile_Seibu

End Sub

' 改行・スペースコードの削除
' パラメタ : 修正前文字列
' 戻り値   : 修正後文字列
Public Function removeSpecCode(ByVal str As String) As String

    Dim retStr As String
    retStr = str
    retStr = Replace(retStr, " ", "")
    retStr = Replace(retStr, "　", "")
    retStr = Replace(retStr, vbCrLf, "")
    retStr = Replace(retStr, vbCr, "")
    retStr = Replace(retStr, vbLf, "")

    removeSpecCode = retStr

End Function

' キーでDictionaryのインデックスを取得
' 未使用20190910
' パラメタ : 対象キー、対象Dict、(Optional)加算index
' 戻り値   : インデックス
Public Function getIndexFromDicByKey(ByVal key As Variant, dic As Object, Optional plusAfterIndex As Integer = 0) As Integer

    Dim keys As Variant
    Dim index As Integer
    index = 0
    For Each keys In dic.keys
        If keys = key Then
            Exit For
        End If
        index = index + 1
    Next
    getIndexFromDicByKey = index + plusAfterIndex
End Function

' ワークシートの取得
' パラメタ : シート名称
' 戻り値   : ワークシート
Public Function getWorSheet(ByVal selectSheet As String) As Worksheet
        ' ワークシートの定義
    Dim s_workSheet As Worksheet
    ' ワークシートが入力されない場合、現在のワークシートとする
    If selectSheet <> "" Then
        Set s_workSheet = Worksheets(selectSheet)
    Else
        Set s_workSheet = ActiveSheet
    End If
    Set getWorSheet = s_workSheet
End Function

' 横列の数を取得(Ｘ線の指定セルから、X線の空白セルまで計算)
' パラメタ : 開始セルアドレス、(Optional)対象シート名
' 戻り値   : カウント数
Public Function getCountHorizon(ByVal cellAddressStart As String, Optional nameWorkSheet As String = "") As Integer

    ' ワークシートの定義
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(nameWorkSheet)
    ' 開始セルの取得
    Dim startCell As Range: Set startCell = s_workSheet.Range(cellAddressStart)
    ' ループカウント
    Dim cntLoop As Integer: cntLoop = 0
    Do While s_workSheet.Cells(startCell.Row, startCell.Column + cntLoop) <> ""
        cntLoop = cntLoop + 1
    Loop

    getCountHorizon = cntLoop

End Function

' Y線リストの取得
' 開始セルが含めいない
' 未使用20190910
' パラメタ : 開始セルアドレス、(Optional)対象シート名
' 戻り値   : Dictionary(キー : セルアドレス、Item : セル内容)
Public Function getListContent(ByVal cellContent As String, Optional nameWorkSheet As String = "", Optional offset As Integer = 0) As Object
    ' Dictionaryの定義
    Dim retOut   As Object: Set retOut = createDictionary
    ' ワークシートの取得
    Dim sworkSheet As Worksheet: Set sworkSheet = getWorSheet(nameWorkSheet)
    ' 開始セルの取得
    Dim cellStart As Range: Set cellStart = sworkSheet.Range(searchCell(cellContent, nameWorkSheet))
    ' ' 開始セル含めないため、オフセット【1】から開始する
    Dim cntLoop As Integer: cntLoop = offset
    ' ループセルの取得
    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).value <> ""
        ' ループセルの設定
        Set Cellcur = sworkSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        ' 主キー重複チェックの設定
        If retOut.Exists(Cellcur.value) = False Then
            retOut.Add Cellcur.Address, Cellcur.value
        End If
        ' ループカウントを足す
        cntLoop = cntLoop + 1
    Loop
    ' 戻り値
    Set getListContent = retOut

End Function

' 配列の長さを算出
' パラメタ : 対象配列
' 戻り値   : カウント数
Public Function getArrayLength(ByRef arry() As String) As Integer
    ' 最後のインデック - 最初のインデック + 1
    getArrayLength = UBound(arry()) - LBound(arry()) + 1

End Function

' セルの検索
' パラメタ : 検索内容、対象ワークシート名
' 戻り値   : セルアドレス
Public Function searchCell(ByVal cellContent As String, Optional worksheetName As String = "") As String
    ' ワークシートの定義
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(worksheetName)
    ' 検索対象を取得
    Dim retRange As Range: Set retRange = s_workSheet.Cells.Find(cellContent, LookAt:=xlWhole)
    ' 検索対象が存在しない場合、処理終了
    If (retRange Is Nothing) Then
        MsgBox cellContent & "を見つかりません。" & vbCrLf & "シェルファイルを作成できません。" _
                , vbYes + vbExclamation, "異常"
        End
    End If
    ' 検索対象のセルを返す
    searchCell = retRange.Address

End Function

' セルの存在確認
' パラメタ : 検索内容、対象ワークシート名
' 戻り値   : セルアドレス
Public Function isCellExist(ByVal cellContent As String, Optional worksheetName As String = "") As Boolean
    isCellExist = False
    ' ワークシートの定義
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(worksheetName)
    ' 検索対象を取得
    Dim retRange As Range: Set retRange = s_workSheet.Cells.Find(cellContent, LookAt:=xlWhole)
    ' 検索対象が存在しない場合、処理終了
    If (retRange Is Nothing = False) Then
        isCellExist = True
    End If
End Function

' ファイル作成
' Linux対応のため、BOMなしのフォマット
' パラメタ : フォルダパス、ファイル絶対名、ファイル内容
' 戻り値   : 成功値(1)のみ
Public Function CreateFileWithoutBom(ByVal folderPath As String, file As String, fileContent As String) As Integer

    ' ファイルパスの宣言
    Dim strFilePath As String
    ' ファイルパス + ファイル名
    strFilePath = ""
    strFilePath = strFilePath + folderPath
    strFilePath = strFilePath + "\"
    strFilePath = strFilePath + file
    ' Bomを削除
    Dim myStream As Object
    Set myStream = CreateObject("ADODB.Stream")
    myStream.Type = 2
    myStream.Charset = "UTF-8"
    myStream.Open
    myStream.WriteText fileContent
    Dim byteData() As Byte
    myStream.Position = 0
    myStream.Type = 1
    myStream.Position = 3
    byteData = myStream.Read
    myStream.Close
    myStream.Open
    myStream.Write byteData
    myStream.SaveToFile strFilePath, 2
    CreateFileWithoutBom = 1
End Function

' ファイルパスの取得
' 存在しない場合、処理終了
' パラメタ : 確認対象パス
' 戻り値   : 入力内容
Public Function getFolderPath(ByVal pathName As String) As String

    ' 出力パス(絶対パス)のセルを検索
    Dim cellPath As Range: Set cellPath = Range(searchCell(pathName))
    ' フォルダパス
    Dim strFilePath As String: strFilePath = Cells(cellPath.Row + 1, cellPath.Column).value

    ' 出力フォルダが存在しない場合
    If Dir(strFilePath, vbDirectory) = "" Then
            MkDir strFilePath
    End If
    
    getFolderPath = strFilePath

End Function

' ケース数の取得
' Y線のセルから、Y線の空白セルまで計算(入力対象セルは対象外)
' パラメタ : セル名
' 戻り値   : カウント数
Public Function getCountCase(ByVal cellName As String, Optional plsuOffset As Integer = 0) As Integer

    Dim getSeiseiCell As String: getSeiseiCell = searchCell(cellName)
    ' 開始カラム
    Dim startCol As Integer: startCol = Range(getSeiseiCell).Column
    ' 開始ROW
    Dim startRow As Integer: startRow = Range(getSeiseiCell).Row + plsuOffset
    ' ケース数
    Dim countCase As Integer: countCase = 0

    ' 次の値が存在しないまで、取得
    Do While Len(Cells(startRow, startCol).value) > 0
        countCase = countCase + 1
        startRow = startRow + 1
    Loop
    ' ケース数を戻す
    getCountCase = countCase

End Function

' グループ内容を取得
' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セル列の内容, 値 : 開始セル列以降の内容)
Public Function getGroupList(ByVal cellContentStart As String, Optional worksheetName As String = "", Optional parameterIncludeSelf As Boolean = False, Optional offset As Integer = 1) As Object
    ' 対象ワークシート
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(worksheetName)
    ' 戻り値
    Dim retOut   As Object: Set retOut = createDictionary
    ' 開始セルを取得
    Dim cellStart As Range: Set cellStart = s_workSheet.Range(searchCell(cellContentStart, worksheetName))
    ' グループの右側までの長さを取得
    Dim lenArray As Integer: lenArray = getCountHorizon(cellStart.Address, worksheetName)
    ' 代入用な配列
    Dim arryStr() As String
    ' 開始セルが含める場合、インデックス数を足す1
    If parameterIncludeSelf = True Then
        ReDim arryStr(lenArray - 1)
    Else
        ReDim arryStr(lenArray - 2)
    End If
    ' ループ用カウント 1 から計算
    Dim cntLoop As Integer: cntLoop = offset

    ' 現在セル
    Dim cellLoop As Range
    ' Y線開始セルからY線空白セルまでループする
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).value <> ""
        ' 現在セルの取得
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        ' 開始セルのX線から右側までの数までループし、配列に代入する
        For lenArray = 0 To getArrayLength(arryStr()) - 1
            arryStr(lenArray) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + 1 + lenArray)
        Next
        ' 主キー重複チェック
        If retOut.Exists(cellLoop.value) = False Then
            retOut.Add cellLoop.value, arryStr()
        End If
        ' ループカウント足す1
        cntLoop = cntLoop + 1
    Loop
    ' 戻り値の設定
    Set getGroupList = retOut

End Function

' ファイル名作成
' パラメタ : ファイル名、ファイル種別、(Optional)追加内容
' 戻り値   : ファイルの絶対名称
Public Function getFileName(ByVal fileName As String, fileType As String, Optional fileOptNm As String = "") As String
    ' ファイルパスの宣言
    Dim strFilePath As String
    strFilePath = ""
    strFilePath = strFilePath + "/"
    strFilePath = strFilePath + fileName
    ' 追加内容が存在する場合
    If fileOptNm <> "" Then
        strFilePath = strFilePath + fileOptNm
    End If
    strFilePath = strFilePath + fileType
    getFileName = strFilePath
End Function

' ループ数による、Y線リスト取得、対象外の場合スキップする
' 対象外 : 空白 / 【-】
' パラメタ : 開始セルアドレス、ループ数、(Optional) 開始オフセット追加
' 戻り値   : Dictionary(キー : セルアドレス, 値 : セル内容)
Public Function getVertivalListbyCnt(ByVal cellAdressStart As String, cntTotalLoop As Integer, Optional rowPlus As Integer) As Object
    ' 戻り値
    Dim retDic As Object: Set retDic = createDictionary
    ' ジョブIDのカラム
    Dim colJobId As Range: Set colJobId = Range(cellAdressStart)
    ' ループ対象のセル
    Dim targetloop As Range
    '追加Rowを初期化
    If rowPlus < 1 Then
        rowPlus = 0
    End If
    ' ループカウント
    Dim cntLoop As Integer
    For cntLoop = 0 To cntTotalLoop - 1
        ' ループ対象のセルをセット
        Set targetloop = Range(Cells(colJobId.Row + rowPlus + cntLoop, colJobId.Column).Address)
        ' ループ対象のセルが値有 かつ 【-】ではない場合
        If Not (IsEmpty(targetloop.value)) And targetloop.value <> "-" Then
            retDic.Add targetloop.Address, targetloop.value
        End If
    Next

    Set getVertivalListbyCnt = retDic

End Function

' ジョブカタログリストの取得(ループ数による、X線のリストを取得、対象外の場合スキップする)
' 対象外 : 空白
' パラメタ : 開始セル内容、ループ数
' 戻り値   : 文字列配列
Public Function getTitleList(ByVal cellContentStart As String, cntLoop As Integer) As String()

    ' 開始セルを検索
    Dim cellTarget As Range: Set cellTarget = Range(searchCell(cellContentStart))
    ' ループ数を初期化
    Dim countLoop As Integer: countLoop = 0
    ' 出力配列の定義
    Dim retList() As String: ReDim retList(0)
    ' 出力配列サイズの定義
    Dim sizeList As Integer
    ' ループで対象を取得
    For countLoop = 0 To cntLoop
        ' 対象セルが空白ではない場合
        If Len(Cells(cellTarget.Row, cellTarget.Column + countLoop).value) > 0 Then
            ' 出力配列の長さを取得
            sizeList = getArrayLength(retList())
            ' 出力配列に最後の値が空白ではない場合
            If Len(retList(sizeList - 1)) > 0 Then
                ' 出力配列を再定義する（以前の値は残る）
                ReDim Preserve retList(sizeList)
                sizeList = getArrayLength(retList())
            End If
                ' 対象を追加
                retList(sizeList - 1) = Cells(cellTarget.Row, cellTarget.Column + countLoop).value
        End If
    Next
    ' 戻り値の設定
    getTitleList = retList()
End Function

' X線リスト数をカウントする（開始セルは含めない）
' 重複チェックあり
' パラメタ : 開始セル内容
' 戻り値   : カウント数
Public Function getCountKomoku(ByVal cellContentStart As String) As Integer
    ' 検索対象セル
    Dim cellTarget As Range: Set cellTarget = Range(searchCell(cellContentStart))
    ' 項目数
    Dim countTotalContent As Integer
    ' ループ数
    Dim countLoop As Integer: countTotalContent = 0
    '  項目名のDic
    Dim dicKomoku As Object: Set dicKomoku = createDictionary
    ' 対象セル
    Dim komokuCell As Range
    ' ループ数を初期化
    countLoop = 0
    ' ループで項目数を取得
    Do While Len(Cells(cellTarget.Row + 1, cellTarget.Column + countLoop).value) > 0
        Set komokuCell = Cells(cellTarget.Row + 1, cellTarget.Column + countLoop)
        If dicKomoku.Exists(komokuCell.value) Then
            MsgBox "項目IDが重複しています。" _
                    , vbYes + vbExclamation, "異常"
            End
        End If
        countTotalContent = countTotalContent + 1
        countLoop = countLoop + 1
        dicKomoku.Add komokuCell.value, komokuCell.Address
    Loop
    ' 項目数を戻す
    getCountKomoku = countTotalContent

End Function

' 文字列のPadLeft
' パラメタ : 修正前文字列
' 戻り値   : 修正後文字列
Public Function padLeftString(ByVal str As String, ByVal char As String, ByVal digit As Long) As String
  Dim tmp As String: tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Right(String(digit, char) & str, digit)
  End If
  padLeftString = tmp
End Function

' 文字列のPadRight
' パラメタ : 修正前文字列
' 戻り値   : 修正後文字列
Public Function padRightString(ByVal str As String, ByVal char As String, ByVal digit As Long) As String
  Dim tmp As String: tmp = str
  If Len(str) < digit And Len(char) > 0 Then
    tmp = Left(str & String(digit, char), digit)
  End If
  padRightString = tmp
End Function

' グループ内容を取得
' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セルアドレス, 値 : 開始セル列以降の内容)
Public Function getGroupListbySelectedValue(ByVal cellContentStart As String, _
                                            Optional worksheetName As String = "", _
                                            Optional parameterIncludeSelf As Boolean = False, _
                                            Optional includeOnly As String = "", _
                                            Optional limitCount As Integer = 0, _
                                            Optional offset As Integer = 1) As Object
    ' ワークシートの取得
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorSheet(worksheetName)
    Dim retOut   As Object: Set retOut = createDictionary
    ' 開始セルを取得
    Dim cellStart As Range: Set cellStart = s_workSheet.Range(searchCell(cellContentStart, worksheetName))
    ' 出力配列の設定
    Dim arryStr() As String
    ' 出力配列長さの設定
    Dim lenOutput As Integer
    ' 配列長さの指定が存在する場合
    If limitCount > 0 Then
        lenOutput = limitCount
        If parameterIncludeSelf = True Then
            ReDim arryStr(lenOutput - 1)
        Else
            ReDim arryStr(lenOutput - 2)
        End If
    Else
        lenOutput = getCountHorizon(cellStart.Address, worksheetName)
        If parameterIncludeSelf = True Then
            ReDim arryStr(lenOutput - 1)
        Else
            ReDim arryStr(lenOutput - 2)
        End If
    End If

    Dim cntLoop As Integer: cntLoop = offset
    Dim cellLoop As Range
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).value <> ""
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        If cellLoop.value = includeOnly Then
            For lenOutput = 0 To getArrayLength(arryStr()) - 1
                arryStr(lenOutput) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + lenOutput)
            Next
            retOut.Add cellLoop.Address, arryStr()
        End If
            cntLoop = cntLoop + 1
    Loop

    Set getGroupListbySelectedValue = retOut

End Function

' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セルアドレス, 値 : 開始セル列以降の内容)
Public Function getListDictionary(ByVal setworksheet As String, listTitle As String, Optional offset As Integer = 1) As Object

    Dim retOut   As Object: Set retOut = createDictionary
    Dim sworkSheet As Worksheet: Set sworkSheet = getWorSheet(setworksheet)
    Dim cnt As Integer: cnt = offset
    ' 処理種別リストを取得
    Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = sworkSheet.Range(searchCell(listTitle, setworksheet))

    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column).value <> ""
        Set Cellcur = sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column)
        If retOut.Exists(Cellcur.value) = False Then
            retOut.Add Cellcur.value, sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column + 1)
        End If
        cnt = cnt + 1
    Loop

    Set getListDictionary = retOut

End Function

' パラメタ : なし
' 戻り値   : Dictionary
Public Function createDictionary() As Object
    Set createDictionary = CreateObject("Scripting.Dictionary")
End Function

' パラメタ : Grpキー、キー、ワークシート
' 戻り値   : 対象キーのItem
Public Function getValueByKeyFromDictionary(ByVal DictGrpKey As String, key As String, Optional worksheetName As String = "") As String

        Dim dict As Object: Set dict = getGroupList(DictGrpKey, worksheetName, True)             ' AcmsファイルIDリストを取得
        getValueByKeyFromDictionary = dict.item(key)(0)

End Function


' フォルダ内容の削除
Public Function deleteAllFileFromFolder(ByVal folderPath As String)

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fl As Object
    
    Set fl = fso.GetFolder(folderPath) ' フォルダを取得
    
    Dim f As Object
    For Each f In fl.Files ' フォルダ内のファイルを取得
        f.Delete (True)         ' ファイルを削除
    Next
    
    ' 後始末
    Set fso = Nothing

End Function


