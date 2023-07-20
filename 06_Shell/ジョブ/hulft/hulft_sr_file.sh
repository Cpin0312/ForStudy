#!/bin/bash

###########################################################################
#
# システム名       : SPC
# ファイル名       : hulft_sr_file.sh
# ファイルタイプ   : B-shell
# 処理名           : HULFT配信要求(ファイル送信)および集信要求(ファイル受信)
# 処理概要         : バッチサーバから配信側として、ファイル伝送を
#                    実行する（HULFTが提供する（/opt/HULFT/bin/）utlsend
#                    コマンドを使用し、ファイル伝送を行う）
#                    バッチサーバから集信側として、ファイル伝送を
#                    実行する（HULFTが提供する（/opt/HULFT/bin/）utlrecv
#                    コマンドを使用し、ファイル伝送を行う）
#                    伝送要求を登録してからファイル伝送の完了をもって
#                    終了とする同期型で実行する
#
# 実行形式         : hulft_sr_file.sh△s|r△ファイルID△相手ホスト名
# 引数             : s[S]、またはr[R] s（配信要求），r（集信要求）
#                    ファイルID（8文字以内の英数字)
#                    相手ホスト名（68文字以内の英数字）
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
CMDDIR="@HULFT_CMD_DIR@"
CMDSEND=${CMDDIR}utlsend
CMDRECV=${CMDDIR}utlrecv
LOGDIR="@HULFT_JOB_LOG_DIR@"
TMPDIR="@HULFT_JOB_TMPLOG_DIR@"
SELFNAME=hulft_sr_file.sh
SELFPID=$$
SELFPREFIX=`echo $SELFNAME | sed 's/\..*//'`

#---------+---------+---------+---------+---------+

#- 引数チェック ----+---------+---------+---------+
if [ $# != 3 ];
then
    echo "Usage: ${SELFNAME} <s/r> <FILE ID> <HOST>"
    logger -p notice -t "${SELFPREFIX}-999-E [${SELFPID}]" "Start Up Parameter Error"
    exit 9
fi

TYPE=$1
TARGET_FILEID=$2
TARGET_HOST=$3

#- コマンド，ログメッセージ編集 --------+---------+---------+
MSGPREFIX="[${SELFPID}][${TYPE}][${TARGET_HOST}][${TARGET_FILEID}]"
CMDEXEC=

if [ "$1" = "s" -o "$1" = "S" ];
then
    CMDEXEC="${CMDSEND} -f ${TARGET_FILEID} -sync"

    LOGMSG001="${SELFPREFIX}-001-I ${MSGPREFIX} ファイル送信処理を開始しました。"
    LOGMSG002="${SELFPREFIX}-002-I ${MSGPREFIX} ファイル送信処理を正常終了しました。"
    LOGMSG003="${SELFPREFIX}-003-E ${MSGPREFIX} ファイル送信処理を異常終了しました。"
fi
if [ "$1" = "r" -o "$1" = "R" ];
then
    CMDEXEC="${CMDRECV} -f ${TARGET_FILEID} -h ${TARGET_HOST} -sync"

    LOGMSG001="${SELFPREFIX}-004-I ${MSGPREFIX} ファイル受信処理を開始しました。"
    LOGMSG002="${SELFPREFIX}-005-I ${MSGPREFIX} ファイル受信処理を正常終了しました。"
    LOGMSG003="${SELFPREFIX}-006-E ${MSGPREFIX} ファイル受信処理を異常終了しました。"
fi

if [ "${CMDEXEC}" = "" ];
then
    echo "Usage: ${SELFNAME} <s/r> <FILE ID> <HOST>"
    logger -p notice -t "${SELFPREFIX}-999-E [${SELFPID}]" "Start Up Parameter Error. Hulft command is null."
    exit 9
fi

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

#- コマンド実行 ----+---------+---------+---------+

echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG001}" >> ${LOG_CURR}
echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG004} ${CMDEXEC}" >> ${LOG_CURR}

TMPSHEAD="`date '+%Y\/%m\/%d %H:%M:%S'` ${LOGMSG006}"
TMPEHEAD="`date '+%Y\/%m\/%d %H:%M:%S'` ${LOGMSG007}"

${CMDEXEC} >> ${TMPLOG} 2>&1
RET=$?

echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG005} ${RET}" >> ${LOG_CURR}

if [ ${RET} = 0 ];
then
    sed "s/^/${TMPSHEAD} /" ${TMPLOG} >> ${LOG_CURR}
    rm -f ${TMPLOG}

    echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG002}" >> ${LOG_CURR}
    echo >> ${LOG_CURR}
    exit 0
else
    sed "s/^/${TMPEHEAD} /" ${TMPLOG} >> ${LOG_CURR}
    rm -f ${TMPLOG}

    echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG003}" >> ${LOG_CURR}
    echo >> ${LOG_CURR}
    exit 9
fi
#---------+---------+---------+---------+---------+
