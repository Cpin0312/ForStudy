#!/bin/sh
#===============================================================

# 作成者     :  日立ソリューションズ
# 作成日     :  2019/08/27
# 変更履歴   :
#    日付      更新者    内容
# 2019/08/27  日立ソリューションズ 初版
#===============================================================


CUR_PATH=`dirname ${0}`
#INPUT=${CUR_PATH}/INPUT
INPUT=${2}
OUTPUT=${CUR_PATH}/OUTPUT_${1}
LIST_PROPERTIES=`find ${INPUT} -type f -name "*.properties"`

# ファイル数でループする
for dir in ${LIST_PROPERTIES};
do
    FILE_NAME=`basename ${dir}`
    OUTPUT_PATH=${dir:${#INPUT}-${#dir}:${#dir}}
    mkdir -p ${OUTPUT}`dirname ${OUTPUT_PATH}`
    if [ ${1} = "SJIS" ]; then
        iconv -f UTF-8 -t SJIS ${dir} > ${OUTPUT}/${OUTPUT_PATH}
    elif [ ${1} = "UTF8" ]; then
        iconv -f SJIS -t UTF-8 ${dir} > ${OUTPUT}/${OUTPUT_PATH}
    fi
done

exit 0
