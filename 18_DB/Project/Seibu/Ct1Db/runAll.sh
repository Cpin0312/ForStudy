rm -rf ./dropSql
rm -rf ./dumpFile

bash ./DbDump.sh
bash ./DbCreateDropAllSeqSql.sh
bash ./DbCreateDropAllTableSql.sh

