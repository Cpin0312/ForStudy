#!/bin/bash
#===============================================================
# Database�̃_���v
# �����T�v             :      �w��DB���_���v����
# �o�͐�               :      �V�G��������[/dumpFile]���쐬���A���Ɋi�[����
# �o�̓t�@�C����       :      dumpFile_�f�[�^�x�[�X��_���s���t(yyyyMMddHHmmss)
#
#===============================================================

# �쐬��
CUR_PATH=`dirname ${0}`

bash ${CUR_PATH}/DbDump.sh

bash ${CUR_PATH}/DbCreateDropAllTableSql.sh
bash ${CUR_PATH}/DbCreateDropAllSeqSql.sh

bash ${CUR_PATH}/DbRunSql.sh       ${CUR_PATH}/dropSql/dropData_tbl_postgres_*
bash ${CUR_PATH}/DbRunSql_Seq.sh   ${CUR_PATH}/dropSql/dropData_seq_postgres_*

bash ${CUR_PATH}/DbRestore.sh      ${1}

exit 0