#!/bin/bash
#===============================================================
# テーブル削除クエリ文の作成(シーケンスのみ)
# 処理概要             :      指定DBの全部のテーブルを取得し、Drop に作成する
# 出力先               :      シエル直下に[/dropSql]を作成し、中に格納する
# 出力ファイル名       :      dropData_seq_データベース名_実行日付(yyyyMMddHHmmss)
#
#===============================================================

# 作成先
TARGET_PATH=`dirname ${0}`/dropSql

# 定義内容
HOST="localhost"
PORT="5433"
USER="meniconuser"
DATABASE="meniconudb"
PASSWORD="meniconuser"

# 削除文クエリ
SQLQUERY_SELECT="SELECT 'Drop sequence '||c.relname||' cascade;' as Sql FROM pg_class c LEFT join pg_user u ON c.relowner = u.usesysid WHERE c.relkind = 'S';"

# バックアップファイルの作成
mkdir -p ${TARGET_PATH}

# 作成名の定義
DROPFILE_NAME="${TARGET_PATH}/dropData_seq_${DATABASE}_`date '+%Y%m%d%H%M%S'`.sql"

# 削除SQL文作成
PGPASSWORD=${PASSWORD} psql -h ${HOST} -U ${USER} -p ${PORT} -d ${DATABASE} -q -c"${SQLQUERY_SELECT}" -t > ${DROPFILE_NAME}

# 作成成功の場合のみ実行
if [ -e ${DROPFILE_NAME} ]; then
    echo "出力ファイルパス : ${DROPFILE_NAME}"
fi

exit 0