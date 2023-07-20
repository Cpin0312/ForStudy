#!/bin/bash
#===============================================================
# Databaseのリストア
# 処理概要             :      指定ダンプをリストアする
# パラメタ             :      対象ダンプファイル
#
#===============================================================

# 実行ダンプファイル（フルパス）
DUMPFILE_FILENAME="$1"

# 実行ダンプファイル（フルパス）の長さ
LEN_FILE=""

if [ $# != 1 ]; then
    echo 実行ファイル（フルパス）が指定されていません。
    exit 0
elif [ ! -e ${DUMPFILE_FILENAME} ]; then
    echo 実行ファイル（フルパス）が存在していません。
    exit 0
else
    LEN_FILE=${#DUMPFILE_FILENAME}
    if [ ${LEN_FILE} -lt 8 ]; then
        echo 実行ファイル（フルパス）が.sql.gzファイルではありません。
        exit 0
    elif [ ${DUMPFILE_FILENAME:LEN_FILE-7:7} != ".sql.gz" ]; then
        echo 実行ファイル（フルパス）が.sql.gzファイルではありません。
        exit 0
    fi
fi

# 定義内容
HOST="10.211.247.104"
PORT="5432"
USER="piadmin"
DATABASE="sspcpostgre_external"
PASSWORD="piadmin"

# リストア
# gzip -cd ${DUMPFILE_FILENAME} | PGPASSWORD=${PASSWORD} psql -h ${HOST} -p ${PORT} -U ${USER} -d ${DATABASE}

# データのみの場合（ユーザ違うなど）
gzip -cd ${DUMPFILE_FILENAME} | PGPASSWORD=${PASSWORD} psql -h ${HOST} -p ${PORT} -U ${USER} -d ${DATABASE} -a

exit 0