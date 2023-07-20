#!/bin/bash
#========================================================================================================================
#
# システム名 :  SPC西武
# ジョブ名   :  AppJobMain.sh
# ジョブ名称 :  共通変数
# ファイルタイプ  :  ?B-shell
# 実行形式        :  AppJobMain.sh ジョブID
#
# リターンコード : 0(成功),失敗(要確認)
#
# 作成者     :  日立ソリューションズ
# 作成日     :  2019/08/27
# 変更履歴   :
#    日付      更新者    内容
# 2019/08/27  日立ソリューションズ 初版
#========================================================================================================================

# 戻り値（初期値は9:異常）
EXIT_CODE=9
# 入力パラメタ
PARAMETER_JOB_ID=${1}
# ACMSシェルパス
ACMS_SHELL_PATH=@ACMS_SHELL_PATH@

# 共通定義の取得
source `dirname ${0}`/common/CommonConf.sh
source `dirname ${0}`/common/LogMethod.sh
source `dirname ${0}`/common/CommonMethod.sh

# *----実行開始----------------------------------------------------------------------------------------------------------
echo ${PARAMETER_JOB_ID} ジョブ実行開始...


if [ $# -eq 0 ]; then
        echo 入力ジョブがありません
        # ログの初期化
        LOG_INIT JB00000000 XXXXXX
        PARAMETER_JOB_ID="JB00000000"
        LOG_JOB_START ${PARAMETER_JOB_ID}
        LOG_PRINT "入力ジョブがありません"
        EXIT_CODE=9
else
    # 実行対象定義の取得
    source `dirname ${0}`/../conf/${PARAMETER_JOB_ID}Config.sh

    # ログの初期化
    LOG_INIT ${PARAMETER_JOB_ID} ${PROC_KBN}
    LOG_JOB_START ${PARAMETER_JOB_ID}

    # *---------------------------------Java処理-------------------------------------------------------------------------
    # Javaバッチの実行
    if [ ${PROC_KBN} = "JAVA" ]; then
        source `dirname ${0}`/java/PisJava.sh ${PARAMETER_JOB_ID} ${@:2:($#-1)}
        EXIT_CODE=$?


    # *---------------------------------ファイル処理---------------------------------------------------------------------
    # ファイル移動処理
    # *------------------------
    elif [ ${PROC_KBN} = "FILE_PROCESS" ]; then
        source `dirname ${0}`/file/FileAction.sh ${PARAMETER_JOB_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi

    # *---------------------------------DB処理---------------------------------------------------------------------------
    # DB削除
    # *------------------------
    elif [ ${PROC_KBN} = "DATA_DELETE" ]; then
        source `dirname ${0}`/database/DataDelete.sh ${PARAMETER_JOB_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi


    # *---------------------------------SFTP処理-------------------------------------------------------------------------
    # SFTPファイル取得
    # *------------------------
    elif [ ${PROC_KBN} = "SFTP_GET" ]; then
        source `dirname ${0}`/sftp/SftpGetFile.sh ${PARAMETER_JOB_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi


    # *------------------------
    # SFTPファイル配置
    # *------------------------
    elif [ ${PROC_KBN} = "SFTP_PUT" ]; then
        source `dirname ${0}`/sftp/SftpPutFile.sh ${PARAMETER_JOB_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi


    # *-----------------------------------HULFT処理----------------------------------------------------------------------
    # HULFT集信ファイルの移動
    # *------------------------
    elif [ ${PROC_KBN} = "HULFT_R_MOVEFILE" ]; then
        bash `dirname ${0}`/hulft/hulft_r_movefile.sh ${HULFT_DEST_FILE_PATH} ${HULFT_TARGET_FILE_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi


    # *------------------------
    # HULFT配信要求(ファイル送信)および集信要求(ファイル受信)
    # *------------------------
    elif [ ${PROC_KBN} = "HULFT_SR_FILE" ]; then
        bash `dirname ${0}`/hulft/hulft_sr_file.sh ${HULFT_TYPE} ${HULFT_TARGET_FILE_ID} ${HULFT_TARGET_HOST_NAME}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi


    # *------------------------
    # HULFT集配信処理後の連携ファイル移動
    # *------------------------
    elif [ ${PROC_KBN} = "HULFT_SR_MOVEFILE" ]; then
        bash `dirname ${0}`/hulft/hulft_sr_movefile.sh ${HULFT_SOURCE_FILE_PATH} ${HULFT_TARGET_FILE_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi


    # *------------------------
    # HULFT配信処理前の連携ファイル移動とバックアップ
    # *------------------------
    elif [ ${PROC_KBN} = "HULFT_S_COPYMOVEFILE" ]; then
        bash `dirname ${0}`/hulft/hulft_s_copymovefile.sh ${HULFT_SOURCE_FILE_PATH} ${HULFT_TARGET_FILE_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi


    # *------------------------
    # HULFT配信処理前の連携ファイル移動
    # *------------------------
    elif [ ${PROC_KBN} = "HULFT_S_MOVEFILE" ]; then
        bash `dirname ${0}`/hulft/hulft_s_movefile.sh ${HULFT_SOURCE_FILE_PATH} ${HULFT_TARGET_FILE_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi


    # *-----------------------------------ACMS処理----------------------------------------------------------------------
    # ACMSアプリケーション起動（JP1）
    # *------------------------
    elif [ ${PROC_KBN} = "ACMS_AP" ]; then
        bash ${ACMS_SHELL_PATH}/JP1_ApLoad.sh ${ACMS_APL_GROUP_ID} ${ACMS_APL_ID} ${ACMS_GROUP_ID} ${ACMS_USER_ID} ${ACMS_FILE_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi

    # *------------------------
    # ACMS発信受信起動（JP1）
    # *------------------------
    elif [ ${PROC_KBN} = "ACMS_REV" ]; then
        bash ${ACMS_SHELL_PATH}/JP1_RcvLoad.sh ${ACMS_GROUP_ID} ${ACMS_USER_ID} ${ACMS_FILE_ID}
        EXIT_CODE=$?
        if [ ${EXIT_CODE} -ne 0 ]; then
        	EXIT_CODE=9
        fi

    # *-----------------------------------異常分岐処理-------------------------------------------------------------------
    else
        echo 異常終了
        EXIT_CODE=9
    fi
fi

#- 結果の出力 ---+---------+---------+---------+
if [ ${EXIT_CODE} -eq 0 ]; then
    echo ${PARAMETER_JOB_ID} ジョブ正常終了...
    # 終了ログ
    LOG_JOB_STOP ${PARAMETER_JOB_ID}
else
    if [ ${EXIT_CODE} -eq 1 ]; then
        echo ${PARAMETER_JOB_ID} ジョブ警告終了...
        # 終了ログ
        LOG_WARNING ${PARAMETER_JOB_ID}
    else
        echo ${PARAMETER_JOB_ID} ジョブ異常終了...
        # 終了ログ
        LOG_JOB_ERROR ${PARAMETER_JOB_ID}
        EXIT_CODE=9
    fi
fi

echo 終了コード:${EXIT_CODE}
exit ${EXIT_CODE}
