#!/bin/sh
#===============================================================

# �쐬��     :  �����\�����[�V�����Y
# �쐬��     :  2019/08/27
# �ύX����   :
#    ���t      �X�V��    ���e
# 2019/08/27  �����\�����[�V�����Y ����
#===============================================================


CUR_PATH=`dirname ${0}`
INPUT=${CUR_PATH}/INPUT
OUTPUT=${CUR_PATH}/OUTPUT_${1}
LIST_PROPERTIES=`find ${INPUT} -type f -name "*.${2}"`

# �t�@�C�����Ń��[�v����
for dir in ${LIST_PROPERTIES};
do
    FILE_NAME=`basename ${dir}`
    OUTPUT_PATH=${dir:${#INPUT}-${#dir}:${#dir}}
    mkdir -p ${OUTPUT}`dirname ${OUTPUT_PATH}`
    iconv -f ${3} -t ${4} ${dir} > ${OUTPUT}/${OUTPUT_PATH}
done

exit 0
