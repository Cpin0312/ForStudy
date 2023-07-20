#!/bin/bash
#================================================================
# Databaseのリストア
# 処理概要             :      指定テーブルのデータをエクスポートする
# パラメタ             :      対象テーブル
#
#===============================================================


# 実行SQLファイル（フルパス）
TBL_NAME="${1^^}"
# 実行SQLファイル（フルパス）の長さ
LEN_FILE=""
# 作成先
TARGET_PATH=`dirname ${0}`/insertSql

if [ ${TBL_NAME} != "" ]; then

    # バックアップファイルの作成
    mkdir -p ${TARGET_PATH}

    # 定義内容
    HOST="localhost"
    PORT="5433"
    USER="itauser"
    DATABASE="itadb"
    PASSWORD="11111111"
    OUTPUT_FILE=${TARGET_PATH}/${TBL_NAME}_partition.sql
    
    # ExportSQLを実行
    # トランザクション開始
	echo "BEGIN;" >> ${OUTPUT_FILE}
    # 削除文クエリ
    SQLQUERY_SELECT="SELECT 'Delete from '||table_name||' ;' as Sql FROM information_schema.tables WHERE table_schema='public' AND table_type='BASE TABLE' AND table_name LIKE '${1,,}%' order by table_name desc"
    # 削除SQL文作成
    PGPASSWORD=${PASSWORD} psql -h ${HOST} -U ${USER} -p ${PORT} -d ${DATABASE} -q -c"${SQLQUERY_SELECT}" -t >> ${OUTPUT_FILE}
    # バックアップSQL文作成
    PGPASSWORD=${PASSWORD} pg_dump -h ${HOST} -p ${PORT} -U ${USER} -d ${DATABASE} --table="${TBL_NAME}*" --data-only --column-inserts >> ${OUTPUT_FILE}
    # トランザクション終了
	echo "COMMIT;" >> ${OUTPUT_FILE}

else

    echo "無効なパラメタ : ${1}"

fi

exit 0