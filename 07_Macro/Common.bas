Attribute VB_Name = "Common"

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

' セルの検索
' パラメタ : 検索内容、対象ワークシート名
' 戻り値   : セルアドレス
Public Function searchCell(ByVal cellContent As String, Optional workSheetName As String = "") As String
    ' ワークシートの定義
    Dim s_workSheet As Worksheet
    Set s_workSheet = getWorSheet(workSheetName)
    ' 出力対象を宣言
    Dim retRange As Range
    ' 検索対象を取得
    Set retRange = s_workSheet.Cells.Find(cellContent, LookAt:=xlWhole)
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
Public Function isCellExist(ByVal cellContent As String, Optional workSheetName As String = "") As Boolean
    isCellExist = False
    ' ワークシートの定義
    Dim s_workSheet As Worksheet
    Set s_workSheet = getWorSheet(workSheetName)
    ' 出力対象を宣言
    Dim retRange As Range
    ' 検索対象を取得
    Set retRange = s_workSheet.Cells.Find(cellContent, LookAt:=xlWhole)
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

On Error GoTo Catch

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

    Exit Function

Catch:
    MsgBox "SQLファイルの作成が失敗しました。"
End Function

' ファイルパスの取得
' 存在しない場合、処理終了
' パラメタ : 確認対象パス
' 戻り値   : 入力内容
Public Function getFolderPath(ByVal pathName As String) As String

    ' ファイルパスの宣言
    Dim strFilePath As String
    ' 出力パス(絶対パス)のセルを検索
    Dim cellPath As Range
    Set cellPath = Range(searchCell(pathName))
    ' フォルダパス
    strFilePath = Cells(cellPath.Row + 1, cellPath.Column).Value
    ' 出力フォルダが存在しない場合
    If Dir(strFilePath, vbDirectory) = "" Then
        MsgBox strFilePath & vbCrLf & "が存在していません。" _
                , vbYes + vbExclamation, "異常"
        End
    End If
    getFolderPath = strFilePath

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
Public Function getListDictionary(ByVal listTitle As String, Optional setworksheet As String = "") As Object

    Dim retOut   As Object
    Set retOut = createDictionary
    Dim sworkSheet As Worksheet
    Set sworkSheet = getWorSheet(setworksheet)
    Dim cnt As Integer
    ' 処理種別リストを取得
    Dim cellSyoriSyubetsu As Range
    Set cellSyoriSyubetsu = sworkSheet.Range(searchCell(listTitle, setworksheet))

    cnt = 1
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

' Dictionaryの作成
Public Function createDictionary() As Object
    Set createDictionary = CreateObject("Scripting.Dictionary")
End Function

Public Function getValueByKeyFromDictionary(ByVal DictGrpKey As String, key As String, Optional workSheetName As String = "") As String

        Dim dict As Object              ' AcmsファイルIDリストを取得
        Set dict = getGroupList(DictGrpKey, workSheetName, True)
        getValueByKeyFromDictionary = dict.Item(key)(0)

End Function

' グループ内容を取得
' パラメタ : 開始セル、ワークシート名、開始セル含めフラグ
' 戻り値   : Dictionary(キー : 開始セルアドレス, 値 : 開始セル列以降の内容)
Public Function getGroupListbySelectedValue(ByVal cellContentStart As String, _
                                                    Optional workSheetName As String = "", _
                                                    Optional includeStartCol As Boolean = True, _
                                                    Optional includeStartRow As Boolean = True, _
                                                    Optional offsetRow As Integer = 0) As Object
    ' ワークシートの取得
    Dim s_workSheet As Worksheet
    Set s_workSheet = getWorSheet(workSheetName)
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
            retOut.Add cellLoop.Address, arryStr()
            cntLoop = cntLoop + 1
    Loop

    Set getGroupListbySelectedValue = retOut

End Function

' 横列の数を取得(Ｘ線の指定セルから、X線の空白セルまで計算)
' パラメタ : 開始セルアドレス、(Optional)対象シート名
' 戻り値   : カウント数
Public Function getCountHorizon(ByVal cellAddressStart As String, Optional nameWorkSheet As String = "") As Integer

    ' ワークシートの定義
    Dim s_workSheet As Worksheet
    Set s_workSheet = getWorSheet(nameWorkSheet)
    ' 開始セルの取得
    Dim startCell As Range
    Set startCell = s_workSheet.Range(cellAddressStart)
    ' ループカウント
    Dim cntLoop As Integer
    cntLoop = 0
    Do While s_workSheet.Cells(startCell.Row, startCell.Column + cntLoop) <> ""
        cntLoop = cntLoop + 1
    Loop

    getCountHorizon = cntLoop

End Function

' 配列の長さを算出
' パラメタ : 対象配列
' 戻り値   : カウント数
Public Function getArrayLength(ByRef arry() As String) As Integer
    ' 最後のインデック - 最初のインデック + 1
    getArrayLength = UBound(arry()) - LBound(arry()) + 1

End Function

' ケース数の取得
' Y線のセルから、Y線の空白セルまで計算(入力対象セルは対象外)
' パラメタ : セル名
' 戻り値   : カウント数
Public Function getCountCase(ByVal cellName As String, Optional workSheetName As String, Optional plsuOffset As Integer = 0) As Integer
    ' 開始カラム
    Dim startCol As Integer
    ' 開始ROW
    Dim startRow As Integer
    ' ケース数
    Dim countCase As Integer

    Dim Worksheet As Worksheet
    Set Worksheet = getWorSheet(workSheetName)

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

Public Function checkExistArray(ByRef ary() As Variant, checkStr As String) As Boolean

    Dim retBool As Boolean
    retBool = False
    Dim cnt As Integer
    For cnt = LBound(ary()) To UBound(ary())

        Dim src As String
        Dim Target As String

        src = convert2Unicode(ary(cnt))
        Target = convert2Unicode(checkStr)

        If src = Target Then
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
    Target = convert2Unicode(input2)

    If src = Target Then
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
    Set sworkSheet = getWorSheet(setworksheet)
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
    Set s_workSheet = getWorSheet(workSheetName)
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




