#!/bin/bash
#===============================================================
# Databaseのリストア
# 処理概要             :      入力のSQL文を実行する
# パラメタ             :      実行対象のSQL文
#
#===============================================================


# 実行SQLファイル（フルパス）
SQL_FILENAME="$1"
# 実行SQLファイル（フルパス）の長さ
LEN_FILE=""

if [ $# != 1 ]; then
    echo 実行ファイル（フルパス）が指定されていません。
    exit 0
elif [ ! -e ${SQL_FILENAME} ]; then
    echo 実行ファイル（フルパス）が存在していません。
    exit 0
else
    LEN_FILE=${#SQL_FILENAME}
    if [ ${LEN_FILE} -lt 5 ]; then
        echo 実行ファイル（フルパス）が.sqlファイルではありません。
        exit 0
    elif [ ${SQL_FILENAME:LEN_FILE-4:4} != ".sql" ]; then
        echo 実行ファイル（フルパス）が.sqlファイルではありません。
        exit 0
    fi
fi

# 定義内容
HOST="localhost"
PORT="5433"
USER="mydatabase"
DATABASE="mydatabase"
PASSWORD="myDatabase"

# 削除SQLを実行
PGPASSWORD=${PASSWORD} psql -h ${HOST} -p ${PORT} -U ${USER} -d ${DATABASE} -a -f ${SQL_FILENAME}

exit 0