Option Explicit

' ���ʌn
Dim WORKBOOK                             ' ���[�N�u�b�N
Dim ACTIVE_SHEET                         ' �V�[�g
Dim NAME_WORKBOOK                        ' ���[�N�u�b�N
Dim PATH_WORKBOOK                        ' ���[�N�u�b�N
Dim PATH_OUTPUT                          ' �o�̓p�X
Dim CNT_ROW                              ' ���s��
Dim OBJ_EXCEL                            ' Excel�I�u�W�F�N�g
Dim objProgressMsg                       ' Makes the object a Public object (Critical!)

' ======================�����J�n======================

showProcessBar (0)
' �����p�X�ݒ�
Call SetDetail
showProcessBar (5)
' ���͓��e���󔒂̏ꍇ�A�������e��������
if (PATH_OUTPUT <> "") then

    ' �ُ�̏ꍇ�A���s����
    'On Error Resume Next
    ' �p�X���쐬����
    ' �����O�m�F���b�Z�[�W ,�yOK:1�z�̏ꍇ�̂ݎ��s
    ' if Msgbox ("�����J�n���܂��B��낵���ł����H",vbOKCancel,"�m�F") = 1 then
    if showMsgOKCancel ("�����J�n���܂��B��낵���ł����H","�m�F") = 1 then

        showProcessBar (10)
        ' ���[�N�u�b�N�̃I�[�v��
        OpenWorkBook(PATH_WORKBOOK)
        showProcessBar (15)
        if NOT (ACTIVE_SHEET is Nothing) then
            CNT_ROW = 0
            ' �{��
            'PI AP�T�[�o
            CreateShell ACTIVE_SHEET , "pspcpap", CNT_ROW , 1
            showProcessBar (30)
            '�o�b�`/IF�T�[�o
            CreateShell ACTIVE_SHEET , "pspcbif", CNT_ROW , 1
            showProcessBar (35)
            '�W���u�Ǘ��T�[�o
            CreateShell ACTIVE_SHEET , "pspcjob", CNT_ROW , 1
            showProcessBar (40)
            'PI/MI�Ǘ��T�[�o
            CreateShell ACTIVE_SHEET , "pspcpmm", CNT_ROW , 1
            showProcessBar (45)
            'CMS�T�[�o
            CreateShell ACTIVE_SHEET , "pspccap", CNT_ROW , 1
            showProcessBar (50)
            'SPC����/emi Web�T�[�o
            CreateShell ACTIVE_SHEET , "pspcfap", CNT_ROW , 1
            showProcessBar (55)
            'CAFIS IF�T�[�o
            CreateShell ACTIVE_SHEET , "pspccif", CNT_ROW , 1
            showProcessBar (60)
            ' ST
            'PI AP�T�[�o
            CreateShell ACTIVE_SHEET , "pspcpap", CNT_ROW , 2
            showProcessBar (65)
            '�o�b�`/IF�T�[�o
            CreateShell ACTIVE_SHEET , "pspcbif", CNT_ROW , 2
            showProcessBar (70)
            '�W���u�Ǘ��T�[�o
            CreateShell ACTIVE_SHEET , "pspcjob", CNT_ROW , 2
            showProcessBar (75)
            'PI/MI�Ǘ��T�[�o
            CreateShell ACTIVE_SHEET , "pspcpmm", CNT_ROW , 2
            showProcessBar (80)
            'CMS�T�[�o
            CreateShell ACTIVE_SHEET , "pspccap", CNT_ROW , 2
            showProcessBar (90)
            'SPC����/emi Web�T�[�o
            CreateShell ACTIVE_SHEET , "pspcfap", CNT_ROW , 2
            showProcessBar (95)
            'CAFIS IF�T�[�o
            CreateShell ACTIVE_SHEET , "pspccif", CNT_ROW , 2

            ' ���[�N�u�b�N�̃N���[�Y
            CloseWorkBook
            showProcessBar (100)
        end if
    End if
else

    showMsg "�o�̓p�X���ݒ肳��Ă��܂���B�������s�ł��܂���"

end if

showMsg "�����I�����܂��I�I�I"

WScript.Quit 0

' ======================�����I��======================

' �����p�X�̐ݒ�
Sub SetDetail()

    Dim OBJECT_FOR_ALL      : Set OBJECT_FOR_ALL    = CreateObject("WScript.Shell")
    ' ���݃p�X
    Dim CUR_PATH            : CUR_PATH              = OBJECT_FOR_ALL.CurrentDirectory & "\"
    ' �Ώۃt�@�C��
    NAME_WORKBOOK                                   = "�ySPC�z�f�B���N�g���ꗗ_�|�C���g�i�A�v���j.xlsx"
    ' �Ώۃt�@�C���̃t���p�X
    PATH_WORKBOOK                                   = CUR_PATH & NAME_WORKBOOK
    ' InputBox�̃��b�Z�[�W
    Dim SHOW_MSG
    SHOW_MSG = ""
    SHOW_MSG = SHOW_MSG & "�ǂݍ��݃t�@�C�� : " & PATH_WORKBOOK & vbCrLf
    SHOW_MSG = SHOW_MSG & vbCrLf & vbCrLf
    SHOW_MSG = SHOW_MSG & "�o�̓p�X����͂��Ă��������B"
    ' �����o�̓p�X
    Dim PATH_OUTPUT_DEFAULT : PATH_OUTPUT_DEFAULT   = "C:\SpcPoint\GitLocal\git_hisol\sql\buildTmp\shell"
    ' ���͂����p�X
    PATH_OUTPUT                                     = showInputBox (SHOW_MSG,"�o�̓p�X�̓���",PATH_OUTPUT_DEFAULT)
    if ( checkWord(getStr(PATH_OUTPUT,1)) = true ) then
        createPath (PATH_OUTPUT)
    else
        PATH_OUTPUT = ""
    end if

    ' �������
    Set OBJECT_FOR_ALL = Nothing
End Sub

' ���[�N�u�b�N�̃I�[�v��
' ����1  : ���[�N�u�b�N�p�X
Function OpenWorkBook(PathBook)
On Error Resume Next
    ' ���[�N�u�b�N��ǂݎ��
    Set OBJ_EXCEL = CreateObject("Excel.Application")
    ' ���[�N�u�b�N�̎擾
    Set ACTIVE_SHEET = OBJ_EXCEL.Workbooks.Open(PathBook).Worksheets("�f�B���N�g���ꗗ")

    if ACTIVE_SHEET is Nothing then
        showMsg "�Ώۃt�@�C�������݂��Ă��܂���"
    end if
    ACTIVE_SHEET.Application.ScreenUpdating = False

End Function

' ���[�N�u�b�N�̃N���[�Y
Sub CloseWorkBook()
    ACTIVE_SHEET.Application.ScreenUpdating = true
    ' ���[�N�u�b�N�����
    OBJ_EXCEL.Quit

End Sub

' �����t�@�C����ǂݍ���
' ����1  : ���[�N�u�b�N
Function CreateShell(activeSheet, targetServer, countCase, envType)

    Dim envRow : Set envRow = activeSheet.Cells.Find("��",,,1)
    Dim env_startCol
    Dim env_startRow
    if (envType = 1) then
        ' �J�n�J������ݒ�
        env_startCol = envRow.Column
        ' �J�nROW��ݒ�
        env_startRow = envRow.Row + 3
    else
        ' �J�n�J������ݒ�
        env_startCol = envRow.Column + 1
        ' �J�nROW��ݒ�
        env_startRow = envRow.Row + 3
    end if

    if countCase = 0 then
        ' �����Ώۂ��擾
        Dim retRange : Set retRange = activeSheet.Cells.Find("#",,,1)
        ' �����Ώۂ��擾�ł����ꍇ
        if Not ( retRange is Nothing) then
            ' �J�n�J������ݒ�
            Dim startCol : startCol = retRange.Column
            ' �J�nROW��ݒ�
            Dim startRow : startRow = retRange.Row + 3
            ' ���̒l�����݂��Ȃ��܂ŁA�擾
            Do While Len(activeSheet.Cells(startRow, startCol).Value) > 0
                countCase = countCase + 1
                startRow = startRow + 1
            Loop
            ' ������
            CNT_ROW = countCase
        end if
    end if

    Dim retServer
    Dim createTarget : createTarget = ""
    if countCase > 0 then
        ' �����T�[�o���擾
        Set retServer = activeSheet.Cells.Find(targetServer,,,1)
        if Not ( retServer is Nothing) then
            createTarget = activeSheet.Cells(retServer.Row - 2, retServer.Column).Value
        end if
    end if

    ' �����Ώۂ��擾�ł����ꍇ
    if countCase > 0 and Not ( retServer is Nothing) and createTarget <> "" then
        ' �J�n�J������ݒ�
        Dim serverCol : serverCol = retServer.Column
        ' �J�nROW��ݒ�
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

                ' �����\�̓��e���擾
                if activeSheet.Cells(serverRow + cnt, serverCol).Value <> "" then
                    ' �p�X
                    Output_Path=activeSheet.Cells(serverRow + cnt, 5).Value
                    ' ����
                    Output_Permission=activeSheet.Cells(serverRow + cnt, 6).Value
                    ' ���[�U
                    Output_User=activeSheet.Cells(serverRow + cnt, 7).Value
                    ' �O���[�v
                    Output_Group=activeSheet.Cells(serverRow + cnt, 8).Value
                    ' �p�r
                    Output_Comment=activeSheet.Cells(serverRow + cnt, 9).Value
                    ' �R�}���h�̍쐬
                    ' ����
                    Output_String = Output_String & "echo " & removeSpecCode(Output_Comment) & vbLf
                    ' ���e�o�͊J�n
                    Output_String = Output_String & "set -x " & vbLf
                    ' �p�X�쐬
                    Output_String = Output_String & "sudo mkdir -p " & Output_Path & vbLf
                    ' �����ύX
                    Output_String = Output_String & "sudo chmod " & Output_Permission & " " & Output_Path & vbLf
                    ' ���[�U�F�O���ύX
                    Output_String = Output_String & "sudo chown " & Output_User & ":" & Output_Group & " " & Output_Path & vbLf
                    ' ���e�o�͏I��
                    Output_String = Output_String & "set +x " & vbLf & vbLf
                end if
            end if
        Next

        ' ���e����̏ꍇ�̂݁A�쐬����
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

' �t�@�C���쐬
' Linux�Ή��̂��߁ABOM�Ȃ��̃t�H�}�b�g
' �p�����^ : �t�H���_�p�X�A�t�@�C����Ζ��A�t�@�C�����e
' �߂�l   : �����l(1)�̂�
Function CreateFileWithoutBom( folderPath, file , fileContent )
    ' �t�@�C���p�X�̐錾
    Dim strFilePath
    ' �t�@�C���p�X + �t�@�C����
    strFilePath = ""
    strFilePath = strFilePath + folderPath
    strFilePath = strFilePath + "\"
    strFilePath = strFilePath + file
    ' Bom���폜
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

' ���s�E�X�y�[�X�R�[�h�̍폜
' �p�����^ : �C���O������
' �߂�l   : �C���㕶����
Function removeSpecCode( str)

    Dim retStr : retStr = """" & str
    retStr = Replace(retStr, vbCrLf, """" & vbLf & "echo """)
    retStr = Replace(retStr, vbCr, """" & vbLf & "echo """)
    retStr = Replace(retStr, vbLf, """" & vbLf & "echo """)
    removeSpecCode = retStr & """"

End Function

' �t�H���_�̍쐬�i�e�t�H���h���쐬�Ώہj
' �p�����^ : �쐬����p�X
' �߂�l   : �Ȃ�
Function createPath(intPath)

    if(intPath <> "") then

        Dim objFso : Set objFso = CreateObject("Scripting.FileSystemObject")
        ' �e�t�H���_�̎擾
        Dim parentPath : parentPath = objFso.GetParentFolderName(intPath)
        ' �Ώېe�t�H���_�̊m�F
        if parentPath <> "" and objFso.FolderExists(parentPath) = false then
            ' �e�t�H���_�̍쐬(�������[�v�I�Ȋ���)
            createPath(parentPath)
        end if

        ' �Ώۃt�H���_�̊m�F
        if objFso.FolderExists(intPath) = false then
            ' �Ώۃt�H���_�̍쐬
            objFso.CreateFolder(intPath)
        end if
        ' ��n��
        Set objFso = Nothing
    else
        Msgbox "�t�@�C���p�X���쐬�ł��܂���B�����𑱍s�ł��܂���B"
        PATH_OUTPUT = ""
    end if

end function

' �w�蕶���̎擾
Function getStr(str , cnt)
    getStr = Left(str, cnt)
End Function

' �����̃`�F�b�N
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

' OKCANCEL���b�Z�[�WBox
Function showMsgOKCancel( strMsg, strTitle)

    ProgressMsg "", "���s���B�B�B"
    showMsgOKCancel = MsgBox (strMsg, vbOKCancel , strTitle)

End function

' Input���b�Z�[�WBox
Function showInputBox( strMsg, strTitle, defaultInput)

    ProgressMsg "", "���s���B�B�B"
    showInputBox = InputBox (strMsg, strTitle, defaultInput)

End function

' ���b�Z�[�WBox
Function showMsg( strMsg)

    ProgressMsg "", "���s���B�B�B"
    MsgBox strMsg

End function

' �i�����b�Z�[�WBox
Function showProcessBar(intPercentage)

    ProgressMsg "", "���s���B�B�B"
    Const SOLID_BLOCK_CHARACTER = "��"
    Const EMPTY_BLOCK_CHARACTER = "��"
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
    ProgressMsg msg, "���s���B�B�B" & intPercentage & "%"

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