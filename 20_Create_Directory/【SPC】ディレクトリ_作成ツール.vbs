Option Explicit

' 共通系
Dim WORKBOOK                             ' ワークブック
Dim ACTIVE_SHEET                         ' シート
Dim NAME_WORKBOOK                        ' ワークブック
Dim PATH_WORKBOOK                        ' ワークブック
Dim PATH_OUTPUT                          ' 出力パス
Dim CNT_ROW                              ' 実行列数
Dim OBJ_EXCEL                            ' Excelオブジェクト
Dim objProgressMsg                       ' Makes the object a Public object (Critical!)

' ======================処理開始======================

showProcessBar (0)
' 初期パス設定
Call SetDetail
showProcessBar (5)
' 入力内容が空白の場合、初期内容を代入する
if (PATH_OUTPUT <> "") then

    ' 異常の場合、続行する
    'On Error Resume Next
    ' パスを作成する
    ' 処理前確認メッセージ ,【OK:1】の場合のみ実行
    ' if Msgbox ("処理開始します。よろしいですか？",vbOKCancel,"確認") = 1 then
    if showMsgOKCancel ("処理開始します。よろしいですか？","確認") = 1 then

        showProcessBar (10)
        ' ワークブックのオープン
        OpenWorkBook(PATH_WORKBOOK)
        showProcessBar (15)
        if NOT (ACTIVE_SHEET is Nothing) then
            CNT_ROW = 0
            ' 本番
            'PI APサーバ
            CreateShell ACTIVE_SHEET , "pspcpap", CNT_ROW , 1
            showProcessBar (30)
            'バッチ/IFサーバ
            CreateShell ACTIVE_SHEET , "pspcbif", CNT_ROW , 1
            showProcessBar (35)
            'ジョブ管理サーバ
            CreateShell ACTIVE_SHEET , "pspcjob", CNT_ROW , 1
            showProcessBar (40)
            'PI/MI管理サーバ
            CreateShell ACTIVE_SHEET , "pspcpmm", CNT_ROW , 1
            showProcessBar (45)
            'CMSサーバ
            CreateShell ACTIVE_SHEET , "pspccap", CNT_ROW , 1
            showProcessBar (50)
            'SPC国内/emi Webサーバ
            CreateShell ACTIVE_SHEET , "pspcfap", CNT_ROW , 1
            showProcessBar (55)
            'CAFIS IFサーバ
            CreateShell ACTIVE_SHEET , "pspccif", CNT_ROW , 1
            showProcessBar (60)
            ' ST
            'PI APサーバ
            CreateShell ACTIVE_SHEET , "pspcpap", CNT_ROW , 2
            showProcessBar (65)
            'バッチ/IFサーバ
            CreateShell ACTIVE_SHEET , "pspcbif", CNT_ROW , 2
            showProcessBar (70)
            'ジョブ管理サーバ
            CreateShell ACTIVE_SHEET , "pspcjob", CNT_ROW , 2
            showProcessBar (75)
            'PI/MI管理サーバ
            CreateShell ACTIVE_SHEET , "pspcpmm", CNT_ROW , 2
            showProcessBar (80)
            'CMSサーバ
            CreateShell ACTIVE_SHEET , "pspccap", CNT_ROW , 2
            showProcessBar (90)
            'SPC国内/emi Webサーバ
            CreateShell ACTIVE_SHEET , "pspcfap", CNT_ROW , 2
            showProcessBar (95)
            'CAFIS IFサーバ
            CreateShell ACTIVE_SHEET , "pspccif", CNT_ROW , 2

            ' ワークブックのクローズ
            CloseWorkBook
            showProcessBar (100)
        end if
    End if
else

    showMsg "出力パスが設定されていません。処理続行できません"

end if

showMsg "処理終了します！！！"

WScript.Quit 0

' ======================処理終了======================

' 初期パスの設定
Sub SetDetail()

    Dim OBJECT_FOR_ALL      : Set OBJECT_FOR_ALL    = CreateObject("WScript.Shell")
    ' 現在パス
    Dim CUR_PATH            : CUR_PATH              = OBJECT_FOR_ALL.CurrentDirectory & "\"
    ' 対象ファイル
    NAME_WORKBOOK                                   = "【SPC】ディレクトリ一覧_ポイント（アプリ）.xlsx"
    ' 対象ファイルのフルパス
    PATH_WORKBOOK                                   = CUR_PATH & NAME_WORKBOOK
    ' InputBoxのメッセージ
    Dim SHOW_MSG
    SHOW_MSG = ""
    SHOW_MSG = SHOW_MSG & "読み込みファイル : " & PATH_WORKBOOK & vbCrLf
    SHOW_MSG = SHOW_MSG & vbCrLf & vbCrLf
    SHOW_MSG = SHOW_MSG & "出力パスを入力してください。"
    ' 初期出力パス
    Dim PATH_OUTPUT_DEFAULT : PATH_OUTPUT_DEFAULT   = "C:\SpcPoint\GitLocal\git_hisol\sql\buildTmp\shell"
    ' 入力したパス
    PATH_OUTPUT                                     = showInputBox (SHOW_MSG,"出力パスの入力",PATH_OUTPUT_DEFAULT)
    if ( checkWord(getStr(PATH_OUTPUT,1)) = true ) then
        createPath (PATH_OUTPUT)
    else
        PATH_OUTPUT = ""
    end if

    ' 解放する
    Set OBJECT_FOR_ALL = Nothing
End Sub

' ワークブックのオープン
' 引数1  : ワークブックパス
Function OpenWorkBook(PathBook)
On Error Resume Next
    ' ワークブックを読み取り
    Set OBJ_EXCEL = CreateObject("Excel.Application")
    ' ワークブックの取得
    Set ACTIVE_SHEET = OBJ_EXCEL.Workbooks.Open(PathBook).Worksheets("ディレクトリ一覧")

    if ACTIVE_SHEET is Nothing then
        showMsg "対象ファイルが存在していません"
    end if
    ACTIVE_SHEET.Application.ScreenUpdating = False

End Function

' ワークブックのクローズ
Sub CloseWorkBook()
    ACTIVE_SHEET.Application.ScreenUpdating = true
    ' ワークブックを解放
    OBJ_EXCEL.Quit

End Sub

' 初期ファイルを読み込み
' 引数1  : ワークブック
Function CreateShell(activeSheet, targetServer, countCase, envType)

    Dim envRow : Set envRow = activeSheet.Cells.Find("環境",,,1)
    Dim env_startCol
    Dim env_startRow
    if (envType = 1) then
        ' 開始カラムを設定
        env_startCol = envRow.Column
        ' 開始ROWを設定
        env_startRow = envRow.Row + 3
    else
        ' 開始カラムを設定
        env_startCol = envRow.Column + 1
        ' 開始ROWを設定
        env_startRow = envRow.Row + 3
    end if

    if countCase = 0 then
        ' 検索対象を取得
        Dim retRange : Set retRange = activeSheet.Cells.Find("#",,,1)
        ' 検索対象を取得できた場合
        if Not ( retRange is Nothing) then
            ' 開始カラムを設定
            Dim startCol : startCol = retRange.Column
            ' 開始ROWを設定
            Dim startRow : startRow = retRange.Row + 3
            ' 次の値が存在しないまで、取得
            Do While Len(activeSheet.Cells(startRow, startCol).Value) > 0
                countCase = countCase + 1
                startRow = startRow + 1
            Loop
            ' 総件数
            CNT_ROW = countCase
        end if
    end if

    Dim retServer
    Dim createTarget : createTarget = ""
    if countCase > 0 then
        ' 検索サーバを取得
        Set retServer = activeSheet.Cells.Find(targetServer,,,1)
        if Not ( retServer is Nothing) then
            createTarget = activeSheet.Cells(retServer.Row - 2, retServer.Column).Value
        end if
    end if

    ' 検索対象を取得できた場合
    if countCase > 0 and Not ( retServer is Nothing) and createTarget <> "" then
        ' 開始カラムを設定
        Dim serverCol : serverCol = retServer.Column
        ' 開始ROWを設定
        Dim serverRow : serverRow = retServer.Row + 2
        Dim cnt : cnt = 0
        Dim Output_String : Output_String = ""
        Dim flgContinue, Output_Comment, Output_Path, Output_Permission, Output_User, Output_Group

        For cnt = 0 To countCase - 1
            flgContinue = false
            Output_Comment=""
            Output_Path=""
            Output_Permission=""
            Output_User=""
            Output_Group=""

            if activeSheet.Cells(serverRow + cnt, 3).Value <> "" and activeSheet.Cells(env_startRow + cnt, env_startCol).Value <> "" then
                flgContinue = true
            end if

            if (flgContinue =true) then

                ' 処理可能の内容を取得
                if activeSheet.Cells(serverRow + cnt, serverCol).Value <> "" then
                    ' パス
                    Output_Path=activeSheet.Cells(serverRow + cnt, 5).Value
                    ' 権限
                    Output_Permission=activeSheet.Cells(serverRow + cnt, 6).Value
                    ' ユーザ
                    Output_User=activeSheet.Cells(serverRow + cnt, 7).Value
                    ' グループ
                    Output_Group=activeSheet.Cells(serverRow + cnt, 8).Value
                    ' 用途
                    Output_Comment=activeSheet.Cells(serverRow + cnt, 9).Value
                    ' コマンドの作成
                    ' 説明
                    Output_String = Output_String & "echo " & removeSpecCode(Output_Comment) & vbLf
                    ' 内容出力開始
                    Output_String = Output_String & "set -x " & vbLf
                    ' パス作成
                    Output_String = Output_String & "sudo mkdir -p " & Output_Path & vbLf
                    ' 権限変更
                    Output_String = Output_String & "sudo chmod " & Output_Permission & " " & Output_Path & vbLf
                    ' ユーザ：グル変更
                    Output_String = Output_String & "sudo chown " & Output_User & ":" & Output_Group & " " & Output_Path & vbLf
                    ' 内容出力終了
                    Output_String = Output_String & "set +x " & vbLf & vbLf
                end if
            end if
        Next

        ' 内容ありの場合のみ、作成する
        if Output_String <> "" then
            Output_String = "#!/bin/bash " & vbLf & vbLf & Output_String
            Output_String = Output_String & "exit 0" & vbLf

            Dim fileName : fileName = targetServer
            if (envType = 1) then
                fileName = "p" & fileName
            else
                fileName = "s" & fileName
            end if

            CreateFileWithoutBom PATH_OUTPUT, "mkdir_" & fileName & ".sh", Output_String
        end if
    end if

End Function

' ファイル作成
' Linux対応のため、BOMなしのフォマット
' パラメタ : フォルダパス、ファイル絶対名、ファイル内容
' 戻り値   : 成功値(1)のみ
Function CreateFileWithoutBom( folderPath, file , fileContent )
    ' ファイルパスの宣言
    Dim strFilePath
    ' ファイルパス + ファイル名
    strFilePath = ""
    strFilePath = strFilePath + folderPath
    strFilePath = strFilePath + "\"
    strFilePath = strFilePath + file
    ' Bomを削除
    Dim myStream
    Set myStream = CreateObject("ADODB.Stream")
    myStream.Type = 2
    myStream.Charset = "UTF-8"
    myStream.Open
    myStream.WriteText fileContent
    Dim byteData
    myStream.Position = 0
    myStream.Type = 1
    myStream.Position = 3
    byteData = myStream.Read
    myStream.Close
    myStream.Open
    myStream.Write byteData
    myStream.SaveToFile strFilePath, 2
    CreateFileWithoutBom = true

End Function

' 改行・スペースコードの削除
' パラメタ : 修正前文字列
' 戻り値   : 修正後文字列
Function removeSpecCode( str)

    Dim retStr : retStr = """" & str
    retStr = Replace(retStr, vbCrLf, """" & vbLf & "echo """)
    retStr = Replace(retStr, vbCr, """" & vbLf & "echo """)
    retStr = Replace(retStr, vbLf, """" & vbLf & "echo """)
    removeSpecCode = retStr & """"

End Function

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
    else
        Msgbox "ファイルパスを作成できません。処理を続行できません。"
        PATH_OUTPUT = ""
    end if

end function

' 指定文字の取得
Function getStr(str , cnt)
    getStr = Left(str, cnt)
End Function

' 文字のチェック
Function checkWord(intStr)
    checkWord = false

    Dim objRegEx : Set objRegEx = CreateObject("VBScript.RegExp")
    objRegEx.Global = True
    objRegEx.Pattern = "[^a-zA-Z0-9]"
    Dim colMatches : Set colMatches = objRegEx.Execute(intStr)
    If colMatches.Count = 0 Then
        checkWord = true
    End If

end function

' OKCANCELメッセージBox
Function showMsgOKCancel( strMsg, strTitle)

    ProgressMsg "", "実行中。。。"
    showMsgOKCancel = MsgBox (strMsg, vbOKCancel , strTitle)

End function

' InputメッセージBox
Function showInputBox( strMsg, strTitle, defaultInput)

    ProgressMsg "", "実行中。。。"
    showInputBox = InputBox (strMsg, strTitle, defaultInput)

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