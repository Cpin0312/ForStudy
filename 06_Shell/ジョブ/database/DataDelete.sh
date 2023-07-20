#!/bin/bash
#===============================================================
#
# システム名 :  SPC西武
# ジョブ名   :  DataUpdate.sh
# ジョブ名称 :  ファイル移動処理
# ファイルタイプ  :  ?B-shell
# 実行形式        :  DataUpdate ジョブID
#
# リターンコード : なし
#
# 作成者     :  日立ソリューションズ
# 作成日     :  2019/08/30
# 変更履歴   :
#    日付      更新者    内容
# 2019/08/30  日立ソリューションズ 初版
#===============================================================


#- returnCode ---------+---------+---------+---------+---------+
errorLevel=9

#- ジョブID
INPUT_JOB_ID=${1}

#- テーブルID ---------+---------+---------+---------+---------+
DB_DELETE_TABLE_ID=${TABLE_ID}

#- 条件
DB_DELETE_CONDITION=${CONDITION}


# *----実行開始-------------------------------------------------
echo ${INPUT_JOB_ID} 処理実行開始...

# 開始ログ
LOG_START

SQLQUERY=
SQLQUERY="DELETE FROM "
SQLQUERY=${SQLQUERY}${DB_DELETE_TABLE_ID}
SQLQUERY=${SQLQUERY}" WHERE "
SQLQUERY=${SQLQUERY}${DB_DELETE_CONDITION}

LOG_PRINT "${SQLQUERY}"

#SQLコマンド実行
PGPASSWORD=${DB_PASSWORD} psql -h ${DB_HOST} -p ${DB_PORT} -U ${DB_USER} -d ${DB_NAME} --set ON_ERROR_STOP=on &>> ${LOG_FILE_PATH} << EOF
BEGIN;${SQLQUERY};COMMIT;
EOF

errorLevel=$?

if [ ${errorLevel} -eq 0 ];then
    echo ${INPUT_JOB_ID} 処理実行終了...
    # 終了ログ
    LOG_PRINT "削除成功しました。"
else
    echo ${INPUT_JOB_ID} 処理異常実行終了...
    # 異常ログ
    LOG_PRINT "削除処理失敗しました。"
fi

# echo 終了コード:${errorLevel}
return ${errorLevel}
