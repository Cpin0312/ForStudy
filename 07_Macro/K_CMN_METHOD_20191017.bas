Attribute VB_Name = "K_CMN_METHOD"
Option Explicit

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
Public Function getWorkSheet(ByVal selectSheet As String) As Worksheet
        ' ワークシートの定義
    Dim s_workSheet As Worksheet
    ' ワークシートが入力されない場合、現在のワークシートとする
    If selectSheet <> "" Then
        Set s_workSheet = Worksheets(selectSheet)
    Else
        Set s_workSheet = ActiveSheet
    End If
    Set getWorkSheet = s_workSheet
End Function

' 横列の数を取得(Ｘ線の指定セルから、X線の空白セルまで計算)
' パラメタ : 開始セルアドレス、(Optional)対象シート名
' 戻り値   : カウント数
Public Function getCountHorizon(ByVal cellAddressStart As String, Optional nameWorkSheet As String = "") As Integer

    ' ワークシートの定義
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorkSheet(nameWorkSheet)
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
    Dim sworkSheet As Worksheet: Set sworkSheet = getWorkSheet(nameWorkSheet)
    ' 開始セルの取得
    Dim cellStart As Range: Set cellStart = sworkSheet.Range(searchCell(cellContent, nameWorkSheet))
    ' ' 開始セル含めないため、オフセット【1】から開始する
    Dim cntLoop As Integer: cntLoop = offset
    ' ループセルの取得
    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).Value <> ""
        ' ループセルの設定
        Set Cellcur = sworkSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        ' 主キー重複チェックの設定
        If retOut.Exists(Cellcur.Value) = False Then
            retOut.Add Cellcur.Address, Cellcur.Value
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
Public Function searchCell(ByVal cellContent As String, Optional workSheetName As String = "") As String
    ' ワークシートの定義
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorkSheet(workSheetName)
    ' 検索対象を取得
    Dim retRange As Range: Set retRange = s_workSheet.Cells.Find(cellContent, LookAt:=xlWhole)
    ' 検索対象が存在しない場合、処理終了
    If (retRange Is Nothing) Then
        showMsg cellContent & "を見つかりません。" & vbCrLf & "シェルファイルを作成できません。" _
                , vbYes + vbExclamation, "異常"
        End
    End If
    ' 検索対象のセルを返す
    searchCell = retRange.Address

End Function

' セルの存在確認
' パラメタ : 検索内容、対象ワークシート名
' 戻り値   : セルアドレス
Public Function isCellExist(ByVal cellContent As String, Optional workSheetName As String = "") As Boolean
    isCellExist = False
    ' ワークシートの定義
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorkSheet(workSheetName)
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
    Dim strFilePath As String: strFilePath = Cells(cellPath.Row + 1, cellPath.Column).Value

    ' 出力フォルダが存在しない場合
    If Dir(strFilePath, vbDirectory) = "" Then
        showMsg strFilePath & vbCrLf & "が存在していません。" _
                , vbYes + vbExclamation, "異常"
        End
    End If
    getFolderPath = strFilePath

End Function

' ケース数の取得
' Y線のセルから、Y線の空白セルまで計算(入力対象セルは対象外)
' パラメタ : セル名
' 戻り値   : カウント数
Public Function getCountCase(ByVal cellName As String, Optional workSheetName As String = "", Optional plsuOffset As Integer = 0) As Integer
    ' 開始カラム
    Dim startCol As Integer
    ' 開始ROW
    Dim startRow As Integer
    ' ケース数
    Dim countCase As Integer

    Dim Worksheet As Worksheet
    Set Worksheet = getWorkSheet(workSheetName)

    getSeiseiCell = searchCell(cellName, workSheetName)
    ' 開始カラムを設定
    startCol = Worksheet.Range(getSeiseiCell).Column
    ' 開始ROWを設定
    startRow = Worksheet.Range(getSeiseiCell).Row + plsuOffset
    ' ケース数を初期化
    countCase = 0
    ' 次の値が存在しないまで、取得
    Do While Len(Worksheet.Cells(startRow, startCol).Value) > 0
        countCase = countCase + 1
        startRow = startRow + 1
    Loop
    ' ケース数を戻す
    getCountCase = countCase

End Function

' グループ内容を取得
' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セル列の内容, 値 : 開始セル列以降の内容)
Public Function getGroupList(ByVal cellContentStart As String, Optional workSheetName As String = "", Optional parameterIncludeSelf As Boolean = False, Optional offset As Integer = 1) As Object
    ' 対象ワークシート
    Dim s_workSheet As Worksheet: Set s_workSheet = getWorkSheet(workSheetName)
    ' 戻り値
    Dim retOut   As Object: Set retOut = createDictionary
    ' 開始セルを取得
    Dim cellStart As Range: Set cellStart = s_workSheet.Range(searchCell(cellContentStart, workSheetName))
    ' グループの右側までの長さを取得
    Dim lenArray As Integer: lenArray = getCountHorizon(cellStart.Address, workSheetName)
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
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).Value <> ""
        ' 現在セルの取得
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
        ' 開始セルのX線から右側までの数までループし、配列に代入する
        For lenArray = 0 To getArrayLength(arryStr()) - 1
            arryStr(lenArray) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + 1 + lenArray)
        Next
        ' 主キー重複チェック
        If retOut.Exists(cellLoop.Value) = False Then
            retOut.Add cellLoop.Value, arryStr()
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

' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セルアドレス, 値 : 開始セル列以降の内容)
Public Function getListDictionary(ByVal listTitle As String, Optional setworksheet As String = "", Optional offset As Integer = 1) As Object

    Dim retOut   As Object: Set retOut = createDictionary
    Dim sworkSheet As Worksheet: Set sworkSheet = getWorkSheet(setworksheet)
    Dim cnt As Integer: cnt = offset
    ' 処理種別リストを取得
    Dim cellSyoriSyubetsu As Range: Set cellSyoriSyubetsu = sworkSheet.Range(searchCell(listTitle, setworksheet))

    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column).Value <> ""
        Set Cellcur = sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column)
        If retOut.Exists(Cellcur.Value) = False Then
            retOut.Add Cellcur.Value, sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column + 1)
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
Public Function getValueByKeyFromDictionary(ByVal DictGrpKey As String, key As String, Optional workSheetName As String = "") As String

        Dim dict As Object: Set dict = getGroupList(DictGrpKey, workSheetName, True)             ' AcmsファイルIDリストを取得
        getValueByKeyFromDictionary = dict.Item(key)(0)

End Function

' グループ内容を取得
' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セルアドレス, 値 : 開始セル列以降の内容)
Public Function getGroupListbySelectedValue(ByVal cellContentStart As String, _
                                                    Optional workSheetName As String = "", _
                                                    Optional includeStartCol As Boolean = True, _
                                                    Optional includeStartRow As Boolean = True, _
                                                    Optional offsetRow As Integer = 0, _
                                                    Optional nmFilter As String = "") As Object
    ' ワークシートの取得
    Dim s_workSheet As Worksheet
    Set s_workSheet = getWorkSheet(workSheetName)
    Dim retOut   As Object
    Set retOut = createDictionary
    ' 開始セルを取得
    Dim cellStart As Range
    Set cellStart = s_workSheet.Range(searchCell(cellContentStart, workSheetName))
    ' 出力配列の設定
    Dim arryStr() As String
    ' 出力配列長さの設定
    Dim lenOutput As Integer

    lenOutput = getCountHorizon(cellStart.Address, workSheetName)
    If includeStartCol = True Then
        ReDim arryStr(lenOutput - 1)
    Else
        ReDim arryStr(lenOutput - 2)
    End If

    Dim cntLoop As Integer
    cntLoop = 0

    If (includeStartRow = False) Then
        cntLoop = 1
    End If

    If (offsetRow > 0) Then
        cntLoop = cntLoop + offsetRow
    End If

    Dim cellLoop As Range
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).Value <> ""
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
            For lenOutput = 0 To getArrayLength(arryStr()) - 1
                arryStr(lenOutput) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + lenOutput)
            Next
            If (nmFilter <> "") Then
                If (arryStr(0) = nmFilter) Then
                    retOut.Add cellLoop.Address, arryStr()
                End If
            Else
                retOut.Add cellLoop.Address, arryStr()
            End If
            cntLoop = cntLoop + 1
    Loop

    Set getGroupListbySelectedValue = retOut

End Function

Public Function checkExistArray(ByRef ary() As Variant, checkStr As Variant) As Boolean

    Dim retBool As Boolean: retBool = False
    Dim cnt As Integer
    Dim src As String
    Dim target As String
    If (checkStr <> "") Then
        target = convert2Unicode(checkStr)
    Else
        target = convert2Unicode(checkVar)
    End If
    
    For cnt = LBound(ary()) To UBound(ary())
    
        src = convert2Unicode(ary(cnt))
        If src = target Then
            retBool = True
            Exit For
        End If

    Next

    checkExistArray = retBool

End Function

Public Function checkExistArrayStr(ByRef ary() As String, checkStr As String) As Boolean

    Dim retBool As Boolean: retBool = False
    Dim cnt As Integer
    Dim src As String
    Dim target As String
    If (checkStr <> "") Then
        target = convert2Unicode(checkStr)
    Else
        target = convert2Unicode(checkVar)
    End If
    
    For cnt = LBound(ary()) To UBound(ary())
    
        src = convert2Unicode(ary(cnt))
        If src = target Then
            retBool = True
            Exit For
        End If

    Next

    checkExistArray = retBool

End Function


Public Function checkStringEqual(ByVal input1 As String, input2 As String) As Boolean

    Dim retBool As Boolean
    retBool = False

    src = convert2Unicode(input1)
    target = convert2Unicode(input2)

    If src = target Then
        retBool = True
    End If

    checkStringEqual = retBool

End Function

Public Function convert2Unicode(ByVal inputStr As String) As String

    convert2Unicode = StrConv(inputStr, vbFromUnicode)

End Function

' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セルアドレス, 値 : 開始セル列以降の内容)
Public Function getListDictionaryAsAddress(ByVal listTitle As String, Optional setworksheet As String) As Object

    Dim retOut   As Object
    Set retOut = createDictionary
    Dim sworkSheet As Worksheet
    Set sworkSheet = getWorkSheet(setworksheet)
    Dim cnt As Integer
    ' 処理種別リストを取得
    Dim cellSyoriSyubetsu As Range
    Set cellSyoriSyubetsu = sworkSheet.Range(searchCell(listTitle, setworksheet))

    cnt = 1
    Dim Cellcur As Range
    Do While sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column).Value <> ""
        Set Cellcur = sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column)
        If retOut.Exists(Cellcur.Value) = False Then
            retOut.Add Cellcur.Address, sworkSheet.Cells(cellSyoriSyubetsu.Row + cnt, cellSyoriSyubetsu.Column)
        End If
        cnt = cnt + 1
    Loop

    Set getListDictionaryAsAddress = retOut

End Function

' 配列の長さを算出
' パラメタ : 対象配列
' 戻り値   : カウント数
Public Function getArrayLengthVariant(ByRef arry() As Variant) As Integer
    ' 最後のインデック - 最初のインデック + 1
    getArrayLengthVariant = UBound(arry()) - LBound(arry()) + 1

End Function

Public Function getLeftString(ByVal str As String, count As Integer) As String

    getLeftString = Left(str, count)

End Function

Public Function getRightString(ByVal str As String, count As Integer) As String

    getRightString = Right(str, count)

End Function

Public Function isQuery(ByVal str As String, pathAry() As String) As String
    Dim retBol As Boolean: retBol = False

    Dim skipBol As Boolean: skipBol = False
    Dim cnt As Integer: cnt = 0

    For cnt = LBound(pathAry()) To UBound(pathAry())
        Dim xxx As Integer
        xxx = InStr(str, pathAry(cnt))

        If (InStr(str, pathAry(cnt)) > 0) Then
            skipBol = True
            Exit For
        End If

    Next

    If (skipBol = False) Then

        Dim strStart As String: strStart = getLeftString(str, 1)
        Dim strEnd As String: strEnd = getRightString(str, 1)

        If (strStart = "@" And strEnd <> "@") Then
            retBol = True
        End If

    End If

    isQuery = retBol

End Function

Public Function removeLeftStr(ByVal str As String, cnt As Integer) As String

    removeLeftStr = Mid(str, 1 + cnt)

End Function

' グループ内容を取得
' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セルアドレス, 値 : 開始セル列以降の内容)
Public Function getGroupListbyCellAddress(ByVal cellAddress As String, _
                                                    Optional workSheetName As String = "", _
                                                    Optional includeStartCol As Boolean = True, _
                                                    Optional includeStartRow As Boolean = True, _
                                                    Optional offsetRow As Integer = 0) As Object
    ' ワークシートの取得
    Dim s_workSheet As Worksheet
    Set s_workSheet = getWorkSheet(workSheetName)
    Dim retOut   As Object
    Set retOut = createDictionary
    ' 開始セルを取得
    Dim cellStart As Range
    Set cellStart = s_workSheet.Range(cellAddress)
    ' 出力配列の設定
    Dim arryStr() As String
    ' 出力配列長さの設定
    Dim lenOutput As Integer

    lenOutput = getCountHorizon(cellStart.Address, workSheetName)
    If includeStartCol = True Then
        ReDim arryStr(lenOutput - 1)
    Else
        ReDim arryStr(lenOutput - 2)
    End If

    Dim cntLoop As Integer
    cntLoop = 0

    If (includeStartRow = False) Then
        cntLoop = 1
    End If

    If (offsetRow > 0) Then
        cntLoop = cntLoop + offsetRow
    End If

    Dim cellLoop As Range
    Do While s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column).Value <> ""
        Set cellLoop = s_workSheet.Cells(cellStart.Row + cntLoop, cellStart.Column)
            For lenOutput = 0 To getArrayLength(arryStr()) - 1
                arryStr(lenOutput) = s_workSheet.Cells(cellLoop.Row, cellStart.Column + lenOutput)
            Next
            retOut.Add cellLoop.Address, arryStr()
            cntLoop = cntLoop + 1
    Loop

    Set getGroupListbyCellAddress = retOut

End Function

Public Function addItemToArray(ByRef targetArry() As Variant, addItem As Variant) As Variant()

    Dim outAry() As Variant: outAry() = targetArry()

    Dim newCnt As Integer: newCnt = getArrayLengthVariant(targetArry())

    If newCnt = 1 And outAry(0) = Empty Then

        outAry(0) = addItem

    Else

        ReDim Preserve outAry(newCnt)

        outAry(newCnt) = addItem

    End If

    addItemToArray = outAry()

End Function

Public Function removeItemFromArray(ByRef targetArray() As Variant, deleteItem As String) As Variant()

    Dim outAry() As Variant: ReDim outAry(0)

    Dim loopStr As Variant

    For Each loopStr In targetArray

        If loopStr <> deleteItem Then

            outAry() = addItemToArray(outAry(), loopStr)

        End If

    Next

    removeItemFromArray = outAry()

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
        If Not (IsEmpty(targetloop.Value)) And targetloop.Value <> "-" Then
            retDic.Add targetloop.Address, targetloop.Value
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
        If Len(Cells(cellTarget.Row, cellTarget.Column + countLoop).Value) > 0 Then
            ' 出力配列の長さを取得
            sizeList = getArrayLength(retList())
            ' 出力配列に最後の値が空白ではない場合
            If Len(retList(sizeList - 1)) > 0 Then
                ' 出力配列を再定義する（以前の値は残る）
                ReDim Preserve retList(sizeList)
                sizeList = getArrayLength(retList())
            End If
                ' 対象を追加
                retList(sizeList - 1) = Cells(cellTarget.Row, cellTarget.Column + countLoop).Value
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
    Do While Len(Cells(cellTarget.Row + 1, cellTarget.Column + countLoop).Value) > 0
        Set komokuCell = Cells(cellTarget.Row + 1, cellTarget.Column + countLoop)
        If dicKomoku.Exists(komokuCell.Value) Then
            showMsg "項目IDが重複しています。" _
                    , vbYes + vbExclamation, "異常"
            End
        End If
        countTotalContent = countTotalContent + 1
        countLoop = countLoop + 1
        dicKomoku.Add komokuCell.Value, komokuCell.Address
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

Public Function saveWorkBookMacro() As String

    saveWorkBook = ""
    'ダイアログで保存先・ファイル名を指定
    Dim strFilePath As String
    strFilePath = Application.GetSaveAsFilename( _
           title:="保存先を選択してください！" _
         , InitialFileName:="initialTBL" _
         , FileFilter:="Excelマクロ有効ブック,*.xlsm")
        
        '指定したパスにファイルが作成済でないかを確認。
    If strFilePath <> "False" And Dir(strFilePath) = "" Then
        '新しいファイルを作成
        Set newBook = Workbooks.Add
        '新しいファイルをVBAを実行したファイルと同じフォルダ保存
        newBook.SaveAs strFilePath, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        saveWorkBook = 0
        
    Else
        If strFilePath = "False" Then
            showMsg "ファイル名が入力されていません。"
        ElseIf Dir(strFilePath) <> "" Then
            
            saveWorkBook = strFilePath
        
        Else
            '既に同名のファイルが存在する場合はメッセージを表示
            showMsg "既に" & newBookName & "というファイルは存在します。"
            
        End If
    End If
    
End Function

Public Function saveWorkBook() As String

    saveWorkBook = ""
    'ダイアログで保存先・ファイル名を指定
    Dim strFilePath As String
    strFilePath = Application.GetSaveAsFilename( _
           title:="保存先を選択してください！" _
         , InitialFileName:="initialTBL" _
         , FileFilter:="Excelブック,*.xlsx")
        
        '指定したパスにファイルが作成済でないかを確認。
    If strFilePath <> "False" And Dir(strFilePath) = "" Then
        '新しいファイルを作成
        Set newBook = Workbooks.Add
        '新しいファイルをVBAを実行したファイルと同じフォルダ保存
        newBook.SaveAs strFilePath, FileFormat:=xlOpenXMLWorkbook
        saveWorkBook = "SUCCESS"
        
    Else
        If strFilePath = "False" Then
            showMsg "ファイル名が入力されていません。"
        ElseIf Dir(strFilePath) <> "" Then
            
            saveWorkBook = strFilePath
        
        Else
            '既に同名のファイルが存在する場合はメッセージを表示
            showMsg "既に" & newBookName & "というファイルは存在します。"
            
        End If
    End If
    
End Function

Public Function delSheet(ByVal sheetName As String)
    Dim loopNm As Worksheet
    For Each loopNm In ActiveWorkbook.Worksheets
        
        If loopNm.name = sheetName Then
            Application.DisplayAlerts = False
            ActiveWorkbook.Worksheets(sheetName).Delete
            Application.DisplayAlerts = True
        End If
    Next
    
End Function

Public Function checkBooksExist(ByVal path As String) As Boolean

    checkBooksExist = False

    If Dir(strFilePath) <> "" Then
        checkBooksExist = True
    End If

End Function


Public Function selectBook() As String
    selectBook = ""
    Dim strFilePath As String
    strFilePath = Application.GetOpenFilename( _
           title:="保存先を選択してください！" _
         , FileFilter:="Excelマクロ有効ブック,*.xlsm")
         
    selectBook = strFilePath

End Function


Public Function showMsg(ByVal msg As String, Optional btn As VbMsgBoxStyle = vbOK, Optional title As String = "") As Integer
    'vbOK 1 [OK]ボタンが押された
    'vbCancel 2 [キャンセル]ボタンが押された
    'vbAbort 3 [中止]ボタンが押された
    'vbRetry 4 [再試行]ボタンが押された
    'vbIgnore 5 [無視]ボタンが押された
    'vbYes 6 [はい]ボタンが押された
    'vbNo 7 [いいえ]ボタンが押された

    showMsg = MsgBox(msg, btn, title)

End Function



