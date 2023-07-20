#!/bin/bash
#===============================================================
#
# システム名 :  SPC西武
# ジョブ名   :  LogMethod.sh
# ジョブ名称 :  共通変数
# ファイルタイプ  :  ?B-shell
# 実行形式        :  LogMethod.sh
#
# リターンコード : なし
#
# 作成者     :  日立ソリューションズ
# 作成日     :  2019/10/08
# 変更履歴   :
#    日付      更新者    内容
# 2019/10/08  日立ソリューションズ 初版
#===============================================================

#===============================================================
#
# 共通処理 : ログ系
#
#===============================================================
# 日付YMD
DATEYMD=$(date +"%Y%m%d")
# 日付YMD hms
DATE_FULL=$(date +"%Y-%m-%d_%H-%M-%S.%3N")
#日付YMD hms
# ユーザ名
LOG_PROCESS_USER=$(whoami)
# 実行ID
LOG_PROCESS_ID=
# 追加ID
LOG_PROCESS_ID_ADD=
# ログパス(/var/app/logs/job/日付フォルダ/jobid.log)
LOG_PATH="@APP_JOB_LOG_DIR@"
# ログファイルパス
LOG_FILE_PATH=

#
# ログ処理の初期化
#
#パラメタ1 : ジョブID・バッチID
#
function LOG_INIT(){
    LOG_PROCESS_ID=${1}
    LOG_PROCESS_ID_ADD=${2}
	LOG_FILE_PATH=
    LOG_FILE_PATH=${LOG_FILE_PATH}${LOG_PATH}/
    LOG_FILE_PATH=${LOG_FILE_PATH}${DATEYMD}/
    LOG_FILE_PATH=${LOG_FILE_PATH}${LOG_PROCESS_ID}.log
}
#
# ログ出力
#
#パラメタ1 : 出力メッセージ
#
function LOG_PRINT(){
    # ディレクトリが存在しない場合
    if [ ! -e ${LOG_FILE_PATH} ]; then
        mkdir -p "$(dirname ${LOG_FILE_PATH})"
    fi
    # ファイルの作成・更新
    touch ${LOG_FILE_PATH}
    # メッセージの設定
    MSG="${DATE_FULL}:${LOG_PROCESS_USER}:${LOG_PROCESS_ID}:${LOG_PROCESS_ID_ADD}:${1}"
    echo ${MSG} >> ${LOG_FILE_PATH}
}

#
# 開始ログ出力
#
function LOG_START(){

    if [  $# -ne 0 ]; then
	    LOG_PRINT "処理を開始しました。(機能ID : ${1})"
    else
	    LOG_PRINT "処理を開始しました。"
    fi
}

#
# 終了ログ出力
#
function LOG_STOP(){

    if [  $# -ne 0 ]; then
	    LOG_PRINT "処理を終了しました。(機能ID : ${1})"
    else
	    LOG_PRINT "処理を終了しました。"
    fi
}

#
# 警告ログ出力
#
function LOG_WARNING(){

    if [  $# -ne 0 ]; then
	    LOG_PRINT "処理を警告しました。(機能ID : ${1})"
    else
	    LOG_PRINT "処理を警告しました。"
    fi
}

#
# タイムアウトログ出力
#
function LOG_TIME_OUT(){
    LOG_PRINT "処理がタイムアウト終了しました。"
}

#
# 異常ログ出力
#
function LOG_ERROR(){
    LOG_PRINT "処理が異常終了しました。"
}

#
# ジョブ開始ログ出力
#
function LOG_JOB_START(){
    LOG_PRINT "ジョブ処理を開始しました。(ジョブID : ${1})"
}

#
# ジョブ終了ログ出力
#
function LOG_JOB_STOP(){
    LOG_PRINT "ジョブ処理を終了しました。(ジョブID : ${1})"
}

#
# ジョブ警告終了ログ出力
#
function LOG_JOB_WARNING(){
    LOG_PRINT "ジョブ警告が異常終了しました。(ジョブID : ${1})"
}

#
# ジョブ異常終了ログ出力
#
function LOG_JOB_ERROR(){
    LOG_PRINT "ジョブ処理が異常終了しました。(ジョブID : ${1})"
}
