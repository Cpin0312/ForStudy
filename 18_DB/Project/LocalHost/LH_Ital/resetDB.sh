#!/bin/bash
#===============================================================
# Databaseのダンプ
# 処理概要             :      指定DBをダンプする
# 出力先               :      シエル直下に[/dumpFile]を作成し、中に格納する
# 出力ファイル名       :      dumpFile_データベース名_実行日付(yyyyMMddHHmmss)
#
#===============================================================

# 作成先
CUR_PATH=`dirname ${0}`

bash ${CUR_PATH}/DbDump.sh

bash ${CUR_PATH}/DbCreateDropAllTableSql.sh
bash ${CUR_PATH}/DbCreateDropAllSeqSql.sh

bash ${CUR_PATH}/DbRunSql.sh       ${CUR_PATH}/dropSql/dropData_tbl_postgres_*
bash ${CUR_PATH}/DbRunSql_Seq.sh   ${CUR_PATH}/dropSql/dropData_seq_postgres_*

bash ${CUR_PATH}/DbRestore.sh      ${1}

exit 0