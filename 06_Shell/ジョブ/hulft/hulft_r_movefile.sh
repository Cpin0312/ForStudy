#!/bin/bash

###########################################################################
#
# システム名       : SPC
# ファイル名       : hulft_r_movefile.sh
# ファイルタイプ   : B-shell
# 処理名           : HULFT集信ファイルの移動
# 処理概要         : HULFT集信ファイルを取り込む前に
#                    集信ファイルを処理フォルダへ移動する
#                  : HULFTからローカルに移動
#
# 実行形式         : hulft_r_movefile.sh△取り込む処理ファイル名△ファイルID
#
# 引数             : HULFT集信ファイル名（フルパス）
#                    ファイルID（8文字以内の英数字)
#
# リターンコード
# 正常終了         : 0
# 異常終了         : 100
#
#-------------------------------------------------------------------------
# 日付         担当者           変更理由
# -----------  ---------------  ------------------------------------------
# 2019/07/31   HISOL            新規作成
# YYYY/MM/DD   bbbbbbb          XXXXの修正
#-------------------------------------------------------------------------
#
###########################################################################
LANG=ja_JP.UTF-8

#- ユーザ変数設定 --+---------+---------+---------+
LOGDIR="@HULFT_JOB_LOG_DIR@"
TMPDIR="@HULFT_JOB_TMPLOG_DIR@"
RECEIVE="@HULFT_RECIEVE_DIR@"
SELFNAME=hulft_r_movefile.sh
SELFPID=$$
SELFPREFIX=`echo $SELFNAME | sed 's/\..*//'`

#---------+---------+---------+---------+---------+

#- 引数チェック ----+---------+---------+---------+
if [ $# != 2 ];
then
    echo "Usage: ${SELFNAME} <FULLPATCH_FILE_NAME> <FILE ID>"
    logger -p notice -t "${SELFPREFIX}-999-E [${SELFPID}]" "Start Up Parameter Error"
    exit 9
fi

TARGET_FILE=$1
TARGET_FILE_ID=$2
TARGET_FOLDER=${RECEIVE}$TARGET_FILE_ID

#---------+---------+---------+---------+---------+

#- 対象ファイル存在チェック --+---------+---------+
if [ -f $TARGET_FILE ];
then
    logger -p notice -t "${SELFPREFIX}-999-I [${SELFPID}]" "ファイルがすでに存在。(ファイルパス：${TARGET_FILE})"
    mv ${TARGET_FILE} ${TARGET_FILE}_`date '+%Y%m%d%H%M%S'`
    RET=$?
    if [ ${RET} -ne 0 ]; then
       logger -p notice -t "${SELFPREFIX}-999-E [${SELFPID}]" "取り込む処理ファイルのリーネムエラー。(ファイル：${TARGET_FILE})"
       exit 9
    fi
fi

#---------+---------+---------+---------+---------+

#- 集信フォルダ存在チェック --+---------+---------+
if [ ! -d $RECEIVE ];
then
    logger -p notice -t "${SELFPREFIX}-999-I [${SELFPID}]" "集信フォルダが存在しません。(フォルダパス：${RECEIVE})"
    mkdir -p ${RECEIVE}
fi

#- 連携ファイル存在チェック --+---------+---------+
if [ ! -f $TARGET_FOLDER ];
then
    logger -p notice -t "${SELFPREFIX}-999-E [${SELFPID}]" "連携ファイルが存在しません。(ファイル：${TARGET_FOLDER})"
    exit 9
fi

#---------+---------+---------+---------+---------+


#- mvコマンド，ログメッセージ編集 --------+---------+
TARGET_FILE_NAME=`echo $TARGET_FILE | awk -F/ '{ print $NF }'`
MSGPREFIX="[${SELFPID}][${TARGET_FILE_NAME}][${TARGET_FILE_ID}]"
CMDEXEC="mv ${TARGET_FOLDER} ${TARGET_FILE}"

LOGMSG001="${SELFPREFIX}-001-I ${MSGPREFIX} ファイル移動処理を開始しました。"
LOGMSG002="${SELFPREFIX}-002-I ${MSGPREFIX} ファイル移動処理を正常終了しました。"
LOGMSG003="${SELFPREFIX}-003-E ${MSGPREFIX} ファイル移動処理を異常終了しました。"
LOGMSG004="${SELFPREFIX}-007-I ${MSGPREFIX} 実行コマンド :"
LOGMSG005="${SELFPREFIX}-008-I ${MSGPREFIX} 戻り値 :"
LOGMSG006="${SELFPREFIX}-009-I ${MSGPREFIX}"
LOGMSG007="${SELFPREFIX}-010-E ${MSGPREFIX}"

#---------+---------+---------+---------+---------+

#- ログファイル生成 +---------+---------+---------+
LOGPREFIX=${SELFPREFIX}

LOG_CURR=${LOGDIR}${LOGPREFIX}.log
TMPLOG=${TMPDIR}${LOGPREFIX}_`date '+%Y%m%d%H%M%S'`.tmp
#---------+---------+---------+---------+---------+

#- mvコマンド実行 ----+---------+---------+---------+
echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG001}" >> ${LOG_CURR}
echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG004} ${CMDEXEC}" >> ${LOG_CURR}

TMPSHEAD="`date '+%Y\/%m\/%d %H:%M:%S'` ${LOGMSG006}"
TMPEHEAD="`date '+%Y\/%m\/%d %H:%M:%S'` ${LOGMSG007}"

${CMDEXEC} >> ${TMPLOG} 2>&1
RET=$?

echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG005} ${RET}" >> ${LOG_CURR}

if [ ${RET} -ne 0 ]; then
   sed "s/^/${TMPEHEAD} /" ${TMPLOG} >> ${LOG_CURR}
   rm -f ${TMPLOG}

   echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG003}" >> ${LOG_CURR}
   echo >> ${LOG_CURR}
   exit 9
fi

sed "s/^/${TMPSHEAD} /" ${TMPLOG} >> ${LOG_CURR}
rm -f ${TMPLOG}
echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG002} ${RET}" >> ${LOG_CURR}
echo >> ${LOG_CURR}
exit 0

#---------+---------+---------+---------+---------+

