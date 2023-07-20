Option Explicit

' DB�n�̒�`
Dim HOST                                            ' DB�z�X�g
Dim PORT                                            ' DB�|�[�g
Dim DATABASE                                        ' DB����
Dim USER                                            ' DB���[�U��
Dim PASSWORD                                        ' DB�p�X���[�h
Dim DB_JAR                                          ' DB���sjar
Dim ADM_USER                                        ' DB���sjar
Dim ADM_PASSWORD                                    ' DB���sjar

' ���e�̒�`
Dim ABATOR_CONFIG_BATCH                             ' Abator�n��`(�o�b�`)
Dim ABATOR_CONFIG_API                               ' Abator�n��`(API)
Dim ABATOR_CONFIG_TEXT                              ' Abator�n��`(���e)
Dim SQL_MAP_BASE_CONFIG_TBL_API                     ' SQLMAP�n��`(API)
Dim SQL_MAP_BASE_CONFIG_TBL_BATCH                   ' SQLMAP�n��`(BATCH)
Dim SQL_MAP_BASE_CONFIG_TXT                         ' SQLMAP�n��`(���e)
Dim APPLICATION_CONTEXT                             ' APPLICATION�n��`(�^�C�g��)
Dim APPLICATION_CONTEXT_CONTENT                     ' APPLICATION�n��`(���e)

' Seq�̒�`���e
Dim SEQ_BATCH_BASEDAO                               ' Seq�o�b�`BaseDao��`
Dim SEQ_BATCH_SQLMAP                                ' Seq�o�b�`SqlMap��`
Dim SEQ_ONLINE_BASEDAO                              ' Seq���A��BaseDao��`
Dim SEQ_ONLINE_BASEDAOIMP                           ' Seq���A��BaseDaoImp��`
Dim SEQ_ONLINE_SQLMAP                               ' Seq���A��SqlMap��`

Dim LIST_SEQ_NAME_CAMEL                             ' SEQ_�L�������P�[�X������
Dim LIST_SEQ_NAME_CLASS_NAME                        ' SEQ_�N���X��
Dim LIST_SEQ_NAME_ID_ORACLE                         ' SEQ_SQLID(Oracle�p)
Dim LIST_SEQ_NAME_ID_POSTGRES                       ' SEQ_SQLID(PostgreSQL�p)
Dim LIST_SEQ_NAME_ID_UCASE                          ' SEQ_�V�[�P���XID�i�啶���j
Dim LIST_SEQ_NAME_ID_LCASE                          ' SEQ_�V�[�P���XID�i�������j
Dim LIST_SEQ_NAME_ID_LCASE_SQLMAP                   ' SEQ_�V�[�P���XID�i�������j
Dim LIST_SEQ_NAME_ID_LCASE_SQLMAP_BASESQLMAP        ' SEQ_�V�[�P���XID�i�������j
Dim LIST_SEQ_NAME_COMMENT                           ' SEQ_�V�[�P���X�R�����g
Dim ARRAY_SEQ_DETAIL()                              ' SEQ_�z��
Dim LIST_ARRAY_SEQ_DETAIL                           ' SEQ�z����i�[���郊�X�g

' �p�X�̒�`
Dim CUR_PATH                                        ' ���݃p�X

' INPUT�n
Dim INPUT_PATH                                      ' INPUT�p�X
Dim INPUT_TBL_DLL_PATH                              ' INPUT�p�X(�e�[�u��DLL)
Dim INPUT_INDEX_DLL_PATH                            ' INPUT�p�X(�C���f�b�N�XDLL)
Dim INPUT_SEQ_DLL_PATH                              ' INPUT�p�X(SEQDLL)

' WORK�n
Dim WORK_PATH                                       ' WORK�p�X
Dim WORK_SQL_PATH                                   ' WORK�p�X(SQL)
Dim WORK_SQL_OUTPUT_PATH                            ' WORK�p�X(OUPUT)
Dim WORK_BAT_PATH                                   ' WORK�p�X(bat)
Dim WORK_VBS_PATH                                   ' WORK�p�X(vbs)
Dim WORK_OUTPUT_XML_PATH                            ' WORK�p�X(xml)
Dim WORK_OUTPUT_TXT_PATH                            ' WORK�p�X(txt)
Dim WORK_SQL_GET_TABLE_LIST                         ' WORK�p�X(�e�[�u�����擾SQL)
Dim WORK_SQL_GET_SEQ_LIST                           ' WORK�p�X(Seq���擾SQL)
Dim WORK_TEMP_REAL_BASEDAO_PATH                     ' RealCommonBaseDao
Dim WORK_TEMP_REAL_SQLMAP_PATH                      ' RealCommonSqlMap
Dim WORK_TEMP_BATCH_BASEDAO_PATH                    ' BatchCommonBaseDao
Dim WORK_TEMP_BATCH_SQLMAP_PATH                     ' BatchCommonSqlMap
Dim WORK_TEMP_BUILD_PATH                            ' Build�t�@�C��
Dim WORK_TEMP_LIB_PATH                              ' lib�t�H���_

' OUTPUT�n
Dim OUTPUT_PATH                                     ' OUTPUT�p�X
Dim OUTPUT_SQLMAP_PATH_API                          ' OUTPUT�p�X(�쐬��API_SQLMAP)
Dim OUTPUT_SQLMAP_PATH_BATCH                        ' OUTPUT�p�X(�쐬��BAT_SQLMAP)
Dim OUTPUT_APP_CONT_PATH                            ' OUTPUT�p�X(�쐬��APPLICATION)
Dim OUTPUT_DAO_PATH                                 ' OUTPUT�p�X(�쐬��DAO)
Dim OUTPUT_SQLMAP_PATH                              ' OUTPUT�p�X(Batch_SQLMAP)
Dim OUTPUT_REAL_DAO_PATH                            ' OUTPUT�p�X(�쐬��RealDAO)
Dim OUTPUT_REAL_SQLMAP_PATH                         ' OUTPUT�p�X(Online_SQLMAP)
Dim OUTPUT_PISCOMMON_PATH                           ' OUTPUT�p�X(PisCommon)
Dim OUTPUT_PISCOMMON_CORE_PATH                      ' OUTPUT�p�X(PisCommon_Src_Core)

' ���ʌn
Dim OBJECT_FOR_ALL                                  ' ���ʂ̃I�u�W�F�N�g
Dim WORKBOOK                                        ' ���ʂ̃��[�N�u�b�N
Dim MESSAGE                                         ' ���ʂ̃��b�Z�[�W�ϐ�
Dim objProgressMsg                                  ' Makes the object a Public object (Critical!)
' ======================�����J�n======================

showProcessBar (0)

' �����p�X�ݒ�
Call SetPath

showProcessBar (5)

' �����O�m�F���b�Z�[�W ,�yOK:1�z�̏ꍇ�̂ݎ��s
MESSAGE = ""
MESSAGE = MESSAGE & "�����J�n���܂��B��낵���ł����H" & vbCrLf
MESSAGE = MESSAGE & vbCrLf
MESSAGE = MESSAGE & "20191201 �X�V���e : �����쐬" & vbCrLf
MESSAGE = MESSAGE & "20191205 �X�V���e : Seq�̑Ή��\" & vbCrLf
MESSAGE = MESSAGE & "20191211 �X�V���e : Abator�̎������s" & vbCrLf
MESSAGE = MESSAGE & "20191211 �X�V���e : IE���ˑ����Ȃ�" & vbCrLf
if showMsgOKCancel (MESSAGE,"�m�F") = 1 then

' �ُ�̏ꍇ�A���s����
On Error Resume Next

    '�e�[�u��DLL�����݂��Ȃ��ꍇ
    if execGetFileCountBatch(INPUT_TBL_DLL_PATH) = 0 and execGetFileCountBatch(INPUT_SEQ_DLL_PATH) = 0 then
        showMsg "�o�^�\TBL�����݂��܂���!!!" & vbCrLf & "����̏����͏I�����܂��B"

    else

        ' �����p�X���쐬
        Call execCreatePathBatch (CUR_PATH)
        ' �i����Ԃ̐ݒ�(10%)
        showProcessBar(10)

        ' �y�ݒ���e.xlsx�z���珉�����e���擾
        Call ReadInitialFile
        ' �i����Ԃ̐ݒ�(20%)
        showProcessBar(20)

        ' Db�ڑ��m�F
        Call execCheckDb

        ' DB�̍폜
        Call execRemoveDb
        ' �i����Ԃ̐ݒ�(30%)
        showProcessBar(30)

        ' DB�̓o�^
        ' �e�[�u��
        Call execRegistDb (INPUT_TBL_DLL_PATH)
        ' �C���f�b�N�X
        Call execRegistDb (INPUT_INDEX_DLL_PATH)
        ' SEQ (�����������܂����A��肠��܂���)
        Call execRegistDb (INPUT_SEQ_DLL_PATH)

        ' �i����Ԃ̐ݒ�(40%)
        showProcessBar(40)

        ' �o�b�`���s�iDB�̑STBL�����擾�j
        Call execGetTableNameListBatch
        ' �i����Ԃ̐ݒ�(50%)
        showProcessBar(50)

        ' �o�b�`���s�iDB�̑SSeq�����擾�j
        Call execGetSeqNameListBatch
        Call setSeqListName
        showProcessBar(55)

        ' Abator�����ݒ���e���擾���āA�u����������
        ' API
        ABATOR_CONFIG_BATCH = ReplaceAstarConfWithUcase(ABATOR_CONFIG_BATCH,ABATOR_CONFIG_TEXT, "@REWRITEHERE_BATCH@")
        ' Batch
        ABATOR_CONFIG_API   = ReplaceAstarConfWithUcase(ABATOR_CONFIG_API, ABATOR_CONFIG_TEXT, "@REWRITEHERE_REAL@")
        ' �i����Ԃ̐ݒ�(60%)
        showProcessBar(60)

        ' AbatorConfig�n���쐬
        Call createConfigFile
        ' �i����Ԃ̐ݒ�(70%)
        showProcessBar(70)

        ' Abator.jar�̎��s
        Call execAbator4JFK
        ' Abator.jar�̎��s���ʂ��ړ�����
        Call execMoveFolder(WORK_BAT_PATH & "java ",OUTPUT_PATH & "java\")
        ' �i����Ԃ̐ݒ�(80%)
        showProcessBar(80)

        ' sqlMap�̍쐬(�u����������)
        ' Online�̑Ή�
        Dim SQLMAP_FOR_API: SQLMAP_FOR_API = SQL_MAP_BASE_CONFIG_TXT
        Call ReplaceAstarConf(WORKBOOK, SQLMAP_FOR_API, SQL_MAP_BASE_CONFIG_TBL_API, "@SQLMAP@")
        Call createFile(OUTPUT_SQLMAP_PATH_API, SQLMAP_FOR_API)

        ' Batch�̑Ή�
        Dim SQLMAP_FOR_BATCH: SQLMAP_FOR_BATCH = SQL_MAP_BASE_CONFIG_TXT
        Call ReplaceAstarConf(WORKBOOK, SQLMAP_FOR_BATCH, SQL_MAP_BASE_CONFIG_TBL_BATCH, "@SQLMAP@")
        Call createFile(OUTPUT_SQLMAP_PATH_BATCH, SQLMAP_FOR_BATCH)
        ' �i����Ԃ̐ݒ�(90%)
        showProcessBar(90)

        ' Dao���擾
        ' applicationContext���쐬(�u����������)
        ' Seq�̑Ή�
        Dim APPLICATION_CONTEXT_CONTENT_SEQ : APPLICATION_CONTEXT_CONTENT_SEQ = APPLICATION_CONTEXT_CONTENT
        APPLICATION_CONTEXT_CONTENT_SEQ = replaceAstarBySeqDao(APPLICATION_CONTEXT_CONTENT_SEQ)
        APPLICATION_CONTEXT = Replace(APPLICATION_CONTEXT, "@APPCONTENT_SEQ@", APPLICATION_CONTEXT_CONTENT_SEQ)
        ' Tbl�̑Ή�
        APPLICATION_CONTEXT_CONTENT = replaceAstarByBaseDao(APPLICATION_CONTEXT_CONTENT)
        APPLICATION_CONTEXT = Replace(APPLICATION_CONTEXT, "@APPCONTENT@", APPLICATION_CONTEXT_CONTENT)
        ' applicationContext�t�@�C���̍쐬
        Call createFile(OUTPUT_APP_CONT_PATH, APPLICATION_CONTEXT)
        ' Seq�t�@�C���̍쐬
        Call createSeqAllFile

        ' CommonBaseDao���R�s�[
        Call execCopyFolder (WORK_TEMP_REAL_BASEDAO_PATH, OUTPUT_REAL_DAO_PATH)
        Call execCopyFolder (WORK_TEMP_BATCH_BASEDAO_PATH, OUTPUT_DAO_PATH)
        Call execCopyFolder (WORK_TEMP_REAL_SQLMAP_PATH, OUTPUT_REAL_SQLMAP_PATH)
        Call execCopyFolder (WORK_TEMP_BATCH_SQLMAP_PATH, OUTPUT_SQLMAP_PATH)

        ' PisCommon�փR�s�[
        Call execCopyFolder (OUTPUT_PATH & "xml\", OUTPUT_PISCOMMON_PATH)
        Call execCopyFolder (OUTPUT_PATH & "java\", OUTPUT_PISCOMMON_CORE_PATH)
        Call execCopyFolder (WORK_TEMP_BUILD_PATH, OUTPUT_PISCOMMON_PATH)
        Call execCopyFolder (WORK_TEMP_LIB_PATH, OUTPUT_PISCOMMON_PATH)

        ' �i����Ԃ̐ݒ�(100%)
        showProcessBar(100)
    End if

    ' �ُ픭���̏ꍇ
    if err <> 0 then
        MESSAGE = ""
        MESSAGE = MESSAGE + "�����r���ɍ쐬���s���܂����B"
        MESSAGE = MESSAGE + vbCrLf
        MESSAGE = MESSAGE + "�Ď��s���Ă�������!!!"
        showMsg MESSAGE
        ' ���₵�����Ȃ��ꍇ�A�R�����g�A�E�g�\
        if showMsg ("�쐬�ς̕����폜���܂����H",vbOKCancel,"�m�F") = 1 then
            err = 0
            Call execDelFileorFolder (OUTPUT_PATH)
            Call execDelFileorFolder (WORK_BAT_PATH & "java")
        End if
    else
        ' �����̏ꍇ
        MESSAGE = ""
        MESSAGE = MESSAGE + "���߂łƂ��I�I�I"
        MESSAGE = MESSAGE + vbCrLf
        MESSAGE = MESSAGE + "�쐬�������܂����I�I�I"
        showMsg MESSAGE
    End if
else
    showMsg "�����I�����܂��I�I�I"
End if

' �����I������
WScript.Quit 0

' ======================�����I��======================

' �����p�X�̐ݒ�
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

    ' �������
    Set OBJECT_FOR_ALL = Nothing

End Sub

' �����t�@�C����ǂݍ���
Sub ReadInitialFile()

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    ' ���[�N�u�b�N��ǂݎ��
    Dim initialXlsxPath : initialXlsxPath = CUR_PATH & "�ݒ���e.xlsx"
    Dim objExcel : Set objExcel = CreateObject("Excel.Application")

    ' ���[�N�u�b�N�̎擾
    Set WORKBOOK = objExcel.Workbooks.Open(initialXlsxPath)
    ' �����f�[�^�̐ݒ�
    Call ReadInitialData(WORKBOOK, "�����ݒ�", 3, 2)
    ' ABATOR�n�̐ݒ�
    ABATOR_CONFIG_BATCH = ReadInitialDataWithReplace(WORKBOOK, "abatorConfigBatch", 3, 2)
    ABATOR_CONFIG_API = ReadInitialDataWithReplace(WORKBOOK, "abatorConfigReal", 2, 2)
    ABATOR_CONFIG_TEXT = ReadInitialDataWithReplace(WORKBOOK, "ConfigText", 2, 2)
    ' SQLMAP�n�̐ݒ�
    SQL_MAP_BASE_CONFIG_TXT = ReadInitialDataWithLoopCnt(WORKBOOK, "sqlMapBaseConfig_001", 18)
    SQL_MAP_BASE_CONFIG_TBL_API = ReadSheetOneCellOnly(WORKBOOK, "sqlMapBaseConfig_API", 2, 2)
    SQL_MAP_BASE_CONFIG_TBL_BATCH = ReadSheetOneCellOnly(WORKBOOK, "sqlMapBaseConfig_Batch", 2, 2)
    ' APPLICATION�n�̐ݒ�
    APPLICATION_CONTEXT = ReadInitialDataWithLoopCnt(WORKBOOK, "applicationContext", 54)
    APPLICATION_CONTEXT_CONTENT = ReadInitialDataWithLoopCnt(WORKBOOK, "applicationContext_Content", 22)

    ' Seq���e�̎擾
    ' Seq�o�b�`BaseDao��`
    SEQ_BATCH_BASEDAO = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_BatchBaseDAO", 1)
    ' Seq�o�b�`SqlMap��`
    SEQ_BATCH_SQLMAP = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_BatchSqlMap", 1)
    ' Seq���A��BaseDao��`
    SEQ_ONLINE_BASEDAO = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_OnlineBaseDao", 1)
    ' Seq���A��BaseDaoImp��`
    SEQ_ONLINE_BASEDAOIMP = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_OnlineBaseDaoImp", 1)
    ' Seq���A��SqlMap��`
    SEQ_ONLINE_SQLMAP = ReadInitialDataWithLoopCntReplaceVbLf(WORKBOOK, "seq_OnlineSqlMap", 1)

    ' ���[�N�u�b�N�����
    objExcel.Quit
End if

End Sub

' �����f�[�^��ǂݍ���
' ����1  : ���[�N�u�b�N�i�I�u�W�F�N�g�j
' ����2  : �V�[�g����   (������)
' ����3  : �J�n��
' ����4  : �J�n�s
' �߂�l : �Ȃ�
Sub ReadInitialData(objWorkbook, sheetName, offsetRow, offsetCol)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    ' �V�[�g�̎擾
    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    ' �J�n��̐ݒ�
    Dim intRow: intRow = offsetRow
    ' �s���󔒂܂Ń��[�v���āA�擾����
    Do Until objWorkSheet.Cells(intRow, offsetCol).Value = ""
        Call setInitialData (objWorkSheet.Cells(intRow, offsetCol), objWorkSheet.Cells(intRow, offsetCol + 1))
        intRow = intRow + 1
    Loop

End if

End Sub

' �����f�[�^��������
' ����1  : �L�[�l
' ����2  : Value�l
' �߂�l : �Ȃ�
Sub setInitialData(iKey, iValue)

    If (iKey = "�z�X�g") Then

        HOST = iValue

    ElseIf (iKey = "�|�[�g") Then

        PORT = iValue

    ElseIf (iKey = "DB��") Then

        DATABASE = iValue

    ElseIf (iKey = "���[�U�[") Then

        USER = iValue

    ElseIf (iKey = "�p�X���[�h") Then

        PASSWORD = iValue

    ElseIf (iKey = "PostgresJar") Then

        DB_JAR = iValue

    ElseIf (iKey = "��ʃ��[�U�[") Then

        ADM_USER = iValue

    ElseIf (iKey = "��ʃp�X���[�h") Then

        ADM_PASSWORD = iValue

    End If

End Sub

' �������e��ǂݍ���ŁA�u����������
' ����1  : ���[�N�u�b�N�i�I�u�W�F�N�g�j
' ����2  : �V�[�g����   (������)
' ����3  : �J�n��
' ����4  : �J�n�s
' �߂�l : �u�������㕶����
Function ReadInitialDataWithReplace(objWorkbook, sheetName, offsetRow, offsetCol)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    ' �V�[�g�̎擾
    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    ' �J�n��̐ݒ�
    Dim intRow: intRow = offsetRow
    ' �߂�l�̐錾
    Dim ConfStr : ConfStr = ""
    ' �s���󔒂܂Ń��[�v���āA�擾����
    Do Until objWorkSheet.Cells(intRow, offsetCol).Value = ""
        ConfStr = ConfStr & objWorkSheet.Cells(intRow, offsetCol) & vbCrLf
        intRow = intRow + 1
    Loop

    ' �擾�������e���㏑������
    ReadInitialDataWithReplace = replaceStr(ConfStr)

End if

End Function

' ������u������
' ����1  : ���͓��e
' �߂�l : �u�������㕶����
Function replaceStr(inStr)

' �ُ�Ȃ��̏ꍇ
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

' X�񃋁[�v�œ��e���擾
' ����1  : ���[�N�u�b�N�i�I�u�W�F�N�g�j
' ����2  : �V�[�g����   (������)
' ����3  : ���[�v��
' �߂�l : �擾����������
Function ReadInitialDataWithLoopCnt(objWorkbook, sheetName, loopCnt)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    Dim ConfStr: ConfStr = ""
    Dim cnt
    ' ���e��2�s�ڂ���B�y�ݒ���e.xlsx�z�ɂĊm�F
    For cnt = 2 To loopCnt+1
        ConfStr = ConfStr & objWorkSheet.Cells(cnt, 2) & vbCrLf
    Next
    ' �ǂݍ��񂾓��e
    ReadInitialDataWithLoopCnt = ConfStr
End if

End Function

' X�񃋁[�v�œ��e���擾
' ����1  : ���[�N�u�b�N�i�I�u�W�F�N�g�j
' ����2  : �V�[�g����   (������)
' ����3  : ���[�v��
' �߂�l : �擾����������
Function ReadInitialDataWithLoopCntReplaceVbLf(objWorkbook, sheetName, loopCnt)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    Dim ConfStr: ConfStr = ""
    Dim cnt
    ' ���e��2�s�ڂ���B�y�ݒ���e.xlsx�z�ɂĊm�F
    For cnt = 2 To loopCnt+1
        ConfStr = ConfStr & objWorkSheet.Cells(cnt, 2) & vbCrLf
    Next

    ConfStr = Replace(ConfStr,vbCrLf,"@XXXX@")
    ConfStr = Replace(ConfStr,vbCr,"@XXXX@")
    ConfStr = Replace(ConfStr,vbLf,"@XXXX@")
    ConfStr = Replace(ConfStr,"@XXXX@",vbCrLf)
    ' �ǂݍ��񂾓��e
    ReadInitialDataWithLoopCntReplaceVbLf = ConfStr
End if

End Function

' �Z����ǂݍ���
' ����1  : ���[�N�u�b�N�i�I�u�W�F�N�g�j
' ����2  : �V�[�g����   (������)
' ����3  : ��ԍ�
' ����4  : �s�ԍ�
' �߂�l : �擾����������
Function ReadSheetOneCellOnly(objWorkbook, sheetName, xRow, yCol)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Dim objWorkSheet: Set objWorkSheet = objWorkbook.Worksheets(sheetName)
    Dim ConfStr: ConfStr = ConfStr & objWorkSheet.Cells(xRow, yCol) & vbCrLf
    ' �ǂݍ��񂾓��e
    ReadSheetOneCellOnly = ConfStr

End if

End Function

' BaseDao�̃��X�g���擾���A�Ώە����y****�z��u������
' ����1  : ���͕�����
' �߂�l : �u����������������
Function replaceAstarByBaseDao(inpurStr)

' �ُ�Ȃ��̏ꍇ
if err = 0 then
    ' BaseDao�̃��X�g���擾
    Dim arryFileName: Set arryFileName = getBaseDaoList

    Dim oStr : oStr = ""
    if Not (arryFileName Is Nothing ) then
        Dim filename
        For Each filename In arryFileName
            ' �u����������
            oStr = oStr & Replace(inpurStr, "****", filename)
        Next
        ' �߂�l
        replaceAstarByBaseDao = oStr
    End if
End if

End Function

' BaseDao�̃��X�g���擾���A�Ώە����y****�z��u������
' ����1  : ���͕�����
' �߂�l : �u����������������
Function replaceAstarBySeqDao(inpurStr)

' �ُ�Ȃ��̏ꍇ
if err = 0 then
    ' BaseDao�̃��X�g���擾
    Dim arryFileName: Set arryFileName = LIST_SEQ_NAME_CLASS_NAME

    Dim oStr : oStr = ""
    Dim filename
    For Each filename In arryFileName
        ' �u����������
        oStr = oStr & Replace(inpurStr, "****", filename)
    Next
    ' �߂�l
    replaceAstarBySeqDao = oStr
End if

End Function

' �e�[�u�����X�g���擾���A�Ώە�����u������
' ����1  : ���͕�����
' ����2  : �u���������������e
' ����3  : �u�������L�[
' �߂�l : �u����������������i�啶���j
Function ReplaceAstarConfWithUcase(outputStr, repSrc, repKey)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Set OBJECT_FOR_ALL = WScript.CreateObject("Scripting.FileSystemObject")
    Dim contentStr: contentStr = ""

    Dim lineStr

    ' �ǂݍ��݃t�@�C���̎w��
    Dim inputFile_db: Set inputFile_db = OBJECT_FOR_ALL.OpenTextFile(WORK_SQL_GET_TABLE_LIST, 1, False, 0)
    ' �ǂݍ��݃t�@�C������1�s���ǂݍ��݁A�����o���t�@�C���ɏ����o���̂��ŏI�s�܂ŌJ��Ԃ�
    Do Until inputFile_db.AtEndOfStream
        lineStr = Trim(inputFile_db.ReadLine)
        If (Len(lineStr) > 0) Then
            lineStr = Replace(repSrc, "****", UCase(lineStr))
        End If
        contentStr = contentStr & "    " & lineStr
    Loop

    ' �u����������
    outputStr = Replace(outputStr, repKey, contentStr)

    Set OBJECT_FOR_ALL = Nothing
    ReplaceAstarConfWithUcase = outputStr

End if

End Function

' �e�[�u�����X�g���擾���A�Ώە�����u������
' ����1  : ���͕�����
' ����2  : �u���������������e
' ����3  : �u�������L�[
' �߂�l : �u����������������
Sub ReplaceAstarConf(objWorkbook, outputStr, repSrc, repKey)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Set OBJECT_FOR_ALL = WScript.CreateObject("Scripting.FileSystemObject")
    Dim contentStrSql: contentStrSql = ""
    Dim contentStr: contentStr = ""
    ' Seq�̓��e��������
    Dim lineStr
    Dim item
    For Each item In LIST_SEQ_NAME_ID_LCASE_SQLMAP
        lineStr = item
        lineStr = Replace(repSrc, "****", lineStr)
        contentStrSql = contentStrSql & "    " & lineStr
    Next
    outputStr = Replace(outputStr, "@SQLMAPSEQ@", contentStrSql)

    if (OBJECT_FOR_ALL.FileExists(WORK_SQL_GET_TABLE_LIST)) then
        ' Tbl�̓��e��������
        Dim inputFile_db: Set inputFile_db = OBJECT_FOR_ALL.OpenTextFile(WORK_SQL_GET_TABLE_LIST, 1, False, 0)

        ' �ǂݍ��݃t�@�C������1�s���ǂݍ��݁A�����o���t�@�C���ɏ����o���̂��ŏI�s�܂ŌJ��Ԃ�
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

' Abator�t�@�C���̍쐬
Sub createConfigFile()

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Dim outputFileName(1)
    outputFileName(0) = WORK_OUTPUT_XML_PATH & "abatorConfigBatch.xml"
    outputFileName(1) = WORK_OUTPUT_XML_PATH & "abatorConfigReal.xml"

    Dim outputStr:  outputStr = ""

    Set OBJECT_FOR_ALL = WScript.CreateObject("Scripting.FileSystemObject")
    Dim cnt

    For cnt = LBound(outputFileName) To UBound(outputFileName)

        ' �����o���t�@�C���̎w�� (����͐V�K�쐬����)
        Dim outputFile: Set outputFile = OBJECT_FOR_ALL.OpenTextFile(outputFileName(cnt), 2, True)
        If (cnt = 0) Then
            outputFile.WriteLine ABATOR_CONFIG_BATCH
        Else
            outputFile.WriteLine ABATOR_CONFIG_API
        End If
        ' �o�b�t�@�� Flush ���ăt�@�C�������
        outputFile.Close

    Next
    Set OBJECT_FOR_ALL = Nothing

End if

End Sub

' �t�@�C���̍쐬
Sub createFile(path, contents)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Dim outputStr:  outputStr = ""
    Set OBJECT_FOR_ALL = WScript.CreateObject("Scripting.FileSystemObject")

    createPath(OBJECT_FOR_ALL.GetParentFolderName(path))
    Dim outputFile: Set outputFile = OBJECT_FOR_ALL.OpenTextFile(path, 2, True)
    outputFile.WriteLine contents
    ' �o�b�t�@�� Flush ���ăt�@�C�������
    outputFile.Close
    Set OBJECT_FOR_ALL = Nothing
End if

End Sub

' �t�@�C���̍쐬
Sub createFile_sjis(path, contents)

' �ُ�Ȃ��̏ꍇ
if err = 0 then
    ' �����o���t�@�C���̎w�� (����͐V�K�쐬����)
    Set OBJECT_FOR_ALL = WScript.CreateObject("ADODB.Stream")
    OBJECT_FOR_ALL.Type = 2
    OBJECT_FOR_ALL.Charset = "Shift-JIS"
    OBJECT_FOR_ALL.Open
    OBJECT_FOR_ALL.WriteText contents
    OBJECT_FOR_ALL.SaveToFile path, 1
    ' �o�b�t�@�� Flush ���ăt�@�C�������
    OBJECT_FOR_ALL.Close
    Set OBJECT_FOR_ALL = Nothing
End if

End Sub

' �o�b�`���s(�t�@�C���E�t�H���_�폜)
' ����1  : �폜�Ώۃp�X
Sub execDelFileorFolder(path)

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "path_delete.bat "
    cmd = cmd & path
    execBatch cmd

End Sub

' �o�b�`���s(�p�X�쐬)
' ����1  : ���݃p�X
Sub execCreatePathBatch(path)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "path_createPath.bat "
    cmd = cmd & path

    execBatch cmd
end if

End Sub

' �o�b�`���s(�t�@�C�����̎擾)
' ����1  : ���݃p�X
Function execGetFileCountBatch(path)

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "file_getCount.bat "
    cmd = cmd & path

    execGetFileCountBatch = execBatchWithResponce (cmd)

End Function

' �o�b�`���s(DB�m�F)
' ����1  : DLL�p�X
Sub execCheckDb

' �ُ�Ȃ��̏ꍇ
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
        showMsg "�f�[�^�x�[�X�ɐڑ��ł��܂���I�I�I" & vbCrLf & "����̏����͏I�����܂��B"
    End if
End if

End Sub

' �o�b�`���s(DB�o�^)
' ����1  : DLL�p�X
Sub execRegistDb(path)

' �ُ�Ȃ��̏ꍇ
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

' �o�b�`���s(DB�o�^)_SJIS
' ����1  : DLL�p�X
Sub execRegistDb_Sjis(path)

' �ُ�Ȃ��̏ꍇ
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

' �o�b�`���s(DB�o�^)
Sub execRemoveDb()

' �ُ�Ȃ��̏ꍇ
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

' �o�b�`���s(DB�̑STBL�����擾)
Sub execGetTableNameListBatch()

' �ُ�Ȃ��̏ꍇ
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

' �o�b�`���s(DB�̑SSeq�����擾)
Sub execGetSeqNameListBatch()

' �ُ�Ȃ��̏ꍇ
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

' �o�b�`���s(����Abator4JFK)
Sub execAbator4JFK()

' �ُ�Ȃ��̏ꍇ
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

' �o�b�`���s(�t�H���_�ړ�)
' ����1  : �ړ���
' ����2  : �ړ���
Sub execMoveFolder(pathSrc, pathDest)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "path_move.bat "
    cmd = cmd & pathSrc & " "
    cmd = cmd & pathDest

    execBatch cmd

End if

End Sub

' �o�b�`���s(�t�H���_�R�s�[)
' ����1  : �ړ���
' ����2  : �ړ���
Sub execCopyFolder(pathSrc, pathDest)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Dim cmd
    cmd = WORK_BAT_PATH
    cmd = cmd & "path_copy.bat "
    cmd = cmd & pathSrc & " "
    cmd = cmd & pathDest

    execBatch cmd

End if

End Sub

' �o�b�`���s
' ����1  : �R�}���h
Function execBatch(cmd)

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    ' WshShell�I�u�W�F�N�g���쐬����
    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")

    ' bat�t�@�C�������s����
    err = WshShell.Run (cmd, 1, True)
    ' �I�u�W�F�N�g���J������
    Set WshShell = Nothing
End if

End Function

' �o�b�`���s�i�B���^�C�v�j
' ����1  : �R�}���h
Function execBatchWithResponce(cmd)

    ' WshShell�I�u�W�F�N�g���쐬����
    Dim WshShell
    Set WshShell = WScript.CreateObject("WScript.Shell")
    ' bat�t�@�C�������s����
    execBatchWithResponce = WshShell.Run (cmd, 1, True)
    ' �I�u�W�F�N�g���J������
    Set WshShell = Nothing
End Function

' BaseDao�̃��X�g���擾
' �߂�l  : BaseDao�̖��̃��X�g
Function getBaseDaoList()

' �ُ�Ȃ��̏ꍇ
if err = 0 then

    Set OBJECT_FOR_ALL = CreateObject("Scripting.FileSystemObject")
    If (OBJECT_FOR_ALL.FolderExists(OUTPUT_DAO_PATH)) then
        Dim folder:Set folder = OBJECT_FOR_ALL.getFolder(OUTPUT_DAO_PATH)
        Dim ArrayList: Set ArrayList = CreateObject("System.Collections.ArrayList")
        ' �t�@�C���ꗗ
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

' SEQ�e���̂̐ݒ�
Sub setSeqListName()
' �ُ�Ȃ��̏ꍇ
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
    inputFile_db.Type = 2    ' 1�F�o�C�i���E2�F�e�L�X�g
    inputFile_db.Charset = "UTF-8"    ' �����R�[�h�w��
    inputFile_db.Open    ' Stream �I�u�W�F�N�g���J��
    inputFile_db.LoadFromFile WORK_SQL_GET_SEQ_LIST    ' �t�@�C����ǂݍ���

    ' �ǂݍ��݃t�@�C������1�s���ǂݍ��݁A�����o���t�@�C���ɏ����o���̂��ŏI�s�܂ŌJ��Ԃ�
    Do Until inputFile_db.EOS

        Dim lineStr : lineStr = Trim(inputFile_db.ReadText(-2))
        Dim str_EN
        Dim str_JP
        If (Len(lineStr) > 0) Then
            Dim nameAry : nameAry = Split(lineStr,",")
            str_EN = nameAry(0)
            str_JP = nameAry(1)
            ' SEQ_�L�������P�[�X������
            LIST_SEQ_NAME_CAMEL.Add (Replace(ProperCase(str_EN),"_",""))
            ' SEQ_�N���X�� = �y@CamelBaseDao@�z
            LIST_SEQ_NAME_CLASS_NAME.Add (Replace(ProperCase(str_EN),"_","") & "BaseDao")
            ' SEQ_SQLID(Oracle�p) = �y@SeqOracle@)�z
            LIST_SEQ_NAME_ID_ORACLE.Add (UCase(str_EN) & ".nextvalue1")
            ' SEQ_SQLID(PostgreSQL�p) = �y@SeqPostgres@�z
            LIST_SEQ_NAME_ID_POSTGRES.Add (UCase(str_EN) & ".nextvalue")
            ' SEQ_�V�[�P���XID�i�啶���j�y@SeqUpperId@�z
            LIST_SEQ_NAME_ID_UCASE.Add (UCase(str_EN))
            ' SEQ_�V�[�P���XID�i�������j
            LIST_SEQ_NAME_ID_LCASE.Add (LCase(str_EN))
            ' SEQ_�V�[�P���XID�i�������j
            LIST_SEQ_NAME_ID_LCASE_SQLMAP_BASESQLMAP.Add ( LCase(str_EN) &  "_BaseSqlMap")
            ' SEQ_�V�[�P���XID�i�������j
            LIST_SEQ_NAME_ID_LCASE_SQLMAP.Add ( LCase(str_EN) )
            ' SEQ_�V�[�P���X�R�����g�y@SeqComment@�z
            LIST_SEQ_NAME_COMMENT.Add (str_JP)

            LIST_SEQ_DETAIL(0) = (Replace(ProperCase(str_EN),"_",""))               ' SEQ_�L�������P�[�X������
            LIST_SEQ_DETAIL(1) = (Replace(ProperCase(str_EN),"_","") & "BaseDao")   ' SEQ_�N���X��
            LIST_SEQ_DETAIL(2) = (str_JP)                                           ' SEQ_�V�[�P���X�R�����g
            LIST_SEQ_DETAIL(3) = (UCase(str_EN) & ".nextvalue1")                    ' SEQ_SQLID(Oracle�p)
            LIST_SEQ_DETAIL(4) = (UCase(str_EN) & ".nextvalue")                     ' SEQ_SQLID(PostgreSQL�p)
            LIST_SEQ_DETAIL(5) = (UCase(str_EN))                                    ' SEQ_�V�[�P���XID�i�啶���j
            LIST_SEQ_DETAIL(6) = (LCase(str_EN))                                    ' SEQ_�V�[�P���XID�i�������j

            LIST_ARRAY_SEQ_DETAIL.Add Join(LIST_SEQ_DETAIL,",")

        End If

    Loop
End if

End Sub

' Seq�̃t�@�C�����쐬
Sub createSeqAllFile()
' �ُ�Ȃ��̏ꍇ
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
    end if

end function

' OKCANCEL���b�Z�[�WBox
Function showMsgOKCancel( strMsg, strTitle)

    ProgressMsg "", "���s���B�B�B"
    showMsgOKCancel = MsgBox (strMsg, vbOKCancel , strTitle)

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

' ���񃁃b�Z�[�W
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