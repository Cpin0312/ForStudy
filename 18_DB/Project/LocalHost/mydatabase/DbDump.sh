#!/bin/bash
#===============================================================
# Databaseのダンプ
# 処理概要             :      指定DBをダンプする
# 出力先               :      シエル直下に[/dumpFile]を作成し、中に格納する
# 出力ファイル名       :      dumpFile_データベース名_実行日付(yyyyMMddHHmmss)
#
#===============================================================

# 作成先
TARGET_PATH=`dirname ${0}`/dumpFile

# 定義内容
HOST="localhost"
PORT="5432"
USER="postgres"
DATABASE="postgres"
PASSWORD="11111111"

# バックアップファイルの作成
mkdir -p ${TARGET_PATH}

# 作成名の定義
DUMPFILE_NAME="${TARGET_PATH}/dumpFile_${DATABASE}_`date '+%Y%m%d%H%M%S'`.sql.gz"

# Dumpファイルの作成
PGPASSWORD=${PASSWORD} pg_dump -h ${HOST} -p ${PORT} -U ${USER} -d ${DATABASE} | gzip > ${DUMPFILE_NAME}

# 作成成功の場合のみ実行
if [ -e ${DUMPFILE_NAME} ]; then
    echo "出力ファイルパス : ${DUMPFILE_NAME}"
fi

exit 0