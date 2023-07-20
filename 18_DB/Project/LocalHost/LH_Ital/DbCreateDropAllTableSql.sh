#!/bin/bash
#===============================================================
# テーブル削除クエリ文の作成
# 処理概要             :      指定DBの全部のテーブルを取得し、Drop に作成する
# 出力先               :      シエル直下に[/dropSql]を作成し、中に格納する
# 出力ファイル名       :      dropData_データベース名_実行日付(yyyyMMddHHmmss)
#
#===============================================================

# 作成先
TARGET_PATH=`dirname ${0}`/dropSql

# 定義内容
HOST="localhost"
PORT="5433"
USER="itauser"
DATABASE="itadb"
PASSWORD="11111111"

# 削除文クエリ
SQLQUERY_SELECT="SELECT 'Drop table '||table_name||' cascade;' as Sql FROM information_schema.tables WHERE table_schema='public' AND table_type='BASE TABLE' order by table_name desc"

# バックアップファイルの作成
mkdir -p ${TARGET_PATH}

# 作成名の定義
rm -rf ${TARGET_PATH}/dropData_tbl_${DATABASE}_*
DROPFILE_NAME="${TARGET_PATH}/dropData_tbl_${DATABASE}_`date '+%Y%m%d%H%M%S'`.sql"

# 削除SQL文作成
PGPASSWORD=${PASSWORD} psql -h ${HOST} -U ${USER} -p ${PORT} -d ${DATABASE} -q -c"${SQLQUERY_SELECT}" -t > ${DROPFILE_NAME}

# 作成成功の場合のみ実行
if [ -e ${DROPFILE_NAME} ]; then
    echo "出力ファイルパス : ${DROPFILE_NAME}"
fi

exit 0