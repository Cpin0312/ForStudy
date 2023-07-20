#!/bin/bash
#===============================================================
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
TARGET_PATH="./insertSql"

if [ ${TBL_NAME} != "" ]; then

	# バックアップファイルの作成
	mkdir -p ${TARGET_PATH}

	# 定義内容
	HOST="@DB_HOST@"
	PORT="@DB_PORT@"
	USER="@DB_UID@"
	DATABASE="@DB_NAME@"
	PASSWORD="@DB_PWD@"

	# ExportSQLを実行
	PGPASSWORD=${PASSWORD} pg_dump -h ${HOST} -p ${PORT} -U ${USER} -d ${DATABASE} --table=${TBL_NAME} --data-only --column-inserts > ${TARGET_PATH}/${TBL_NAME}.sql
else

	echo "無効なパラメタ : ${1}"

fi

exit 0