#!/bin/sh
#===============================================================

# 作成者     :  日立ソリューションズ
# 作成日     :  2019/08/27
# 変更履歴   :
#    日付      更新者    内容
# 2019/08/27  日立ソリューションズ 初版
#===============================================================


CUR_PATH=`dirname ${0}`
INPUT=${CUR_PATH}/INPUT
OUTPUT=${CUR_PATH}/OUTPUT_${1}
LIST_PROPERTIES=`find ${INPUT} -type f -name "*.${2}"`

# ファイル数でループする
for dir in ${LIST_PROPERTIES};
do
    FILE_NAME=`basename ${dir}`
    OUTPUT_PATH=${dir:${#INPUT}-${#dir}:${#dir}}
    mkdir -p ${OUTPUT}`dirname ${OUTPUT_PATH}`
    iconv -f ${3} -t ${4} ${dir} > ${OUTPUT}/${OUTPUT_PATH}
done

exit 0
