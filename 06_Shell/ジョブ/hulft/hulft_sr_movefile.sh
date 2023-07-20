#!/bin/bash

###########################################################################
#
# システム名       : SPC
# ファイル名       : hulft_sr_movefile.sh
# ファイルタイプ   : B-shell
# 処理名           : HULFT集配信処理後の連携ファイル移動
# 処理概要         : HULFT配信及び集信要求後に連携ファイルを
#                    HULFT集配信フォルダからバックアップフォルダへ移動する
#                    正常終了時は正常終了バックアップフォルダ
#                    異常終了時は異常終了バックアップフォルダへ移動する
#                  : ローカルからバックアップフォルダにコピー
#                  : ローカルを削除
#
# 実行形式         : hulft_sr_movefile.sh△HULFT送受信ファイル名△ファイルID
#                                       △移動先フォルダ
# 引数             : HULFT送受信ファイル名（フルパス）
#                    ファイルID（8文字以内の英数字)
#                    移動先フォルダ（バックアップフォルダフルパス）
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
SELFNAME=hulft_sr_movefile.sh
SELFPID=$$
SELFPREFIX=`echo $SELFNAME | sed 's/\..*//'`

#---------+---------+---------+---------+---------+

#- 引数チェック ----+---------+---------+---------+
if [ $# != 2 ];
then
    echo "Usage: ${SELFNAME} <FULLPATCH_FILE_NAME> <FILE ID> <BACK_FOLDER>"
    logger -p notice -t "${SELFPREFIX}-999-E [${SELFPID}]" "Start Up Parameter Error"
    exit 9
fi

TARGET_FILE=$1
TARGET_FILEID=$2
#ディレクトリ最後のスラッシュを削除
TARGET_FOLDER=@HULFT_BACKUP_DIR@
TARGET_FOLDER=${TARGET_FOLDER}${TARGET_FILEID}

#---------+---------+---------+---------+---------+

#- 対象ファイル存在チェック --+---------+---------+
if [ ! -f $TARGET_FILE ];
then
    logger -p notice -t "${SELFPREFIX}-999-E [${SELFPID}]" "ファイルが存在しません。(ファイルパス：${TARGET_FILE})"
    exit 9
fi

#---------+---------+---------+---------+---------+

#- 移動先フォルダ存在チェック --+---------+---------+
if [ ! -d $TARGET_FOLDER ];
then
    logger -p notice -t "${SELFPREFIX}-999-I [${SELFPID}]" "バックアップフォルダが存在しません。(フォルダパス：${TARGET_FOLDER})"
    mkdir -p ${TARGET_FOLDER}
fi

#---------+---------+---------+---------+---------+


#- cpコマンド，ログメッセージ編集 --------+---------+
TARGET_FILE_NAME=`echo $TARGET_FILE | awk -F/ '{ print $NF }'`
MSGPREFIX="[${SELFPID}][${TARGET_FILE_NAME}][${TARGET_FILEID}]"
CMDEXEC="cp -fp ${TARGET_FILE} ${TARGET_FOLDER}/${TARGET_FILEID}"_`date '+%Y%m%d%H%M%S'`


if [ "${CMDEXEC}" = "" ];
then
    echo "Usage: ${SELFNAME} <FULLPATCH_FILE_NAME> <FILE ID> <BACK_FOLDER>"
    logger -p notice -t "${SELFPREFIX}-999-E [${SELFPID}]" "Start Up Parameter Error. OS command is null."
    exit 9
fi

LOGMSG001="${SELFPREFIX}-001-I ${MSGPREFIX} ファイルコピー処理を開始しました。"
LOGMSG002="${SELFPREFIX}-002-I ${MSGPREFIX} ファイルコピー処理を正常終了しました。"
LOGMSG003="${SELFPREFIX}-003-E ${MSGPREFIX} ファイルコピー処理を異常終了しました。"
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

#- cpコマンド実行 ----+---------+---------+---------+
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
echo "`date '+%Y/%m/%d %H:%M:%S'` ${LOGMSG002} ${RET}" >> ${LOG_CURR}

#---------+---------+---------+---------+---------+

#- rmコマンド，ログメッセージ編集 --------+---------+
CMDEXEC="rm -f ${TARGET_FILE}"

LOGMSG001="${SELFPREFIX}-001-I ${MSGPREFIX} ファイル削除処理を開始しました。"
LOGMSG002="${SELFPREFIX}-002-I ${MSGPREFIX} ファイル削除処理を正常終了しました。"
LOGMSG003="${SELFPREFIX}-003-E ${MSGPREFIX} ファイル削除処理を異常終了しました。"
LOGMSG004="${SELFPREFIX}-007-I ${MSGPREFIX} 実行コマンド :"
LOGMSG005="${SELFPREFIX}-008-I ${MSGPREFIX} 戻り値 :"
LOGMSG006="${SELFPREFIX}-009-I ${MSGPREFIX}"
LOGMSG007="${SELFPREFIX}-010-E ${MSGPREFIX}"

#---------+---------+---------+---------+---------+

#- rmコマンド実行 ----+---------+---------+---------+
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

