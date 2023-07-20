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
OUTPUT=${CUR_PATH}/OUTPUT
LIST_PROPERTIES=`find ${INPUT} -type f -name "*.properties"`

    # �t�@�C�����Ń��[�v����
	for dir in ${LIST_PROPERTIES};
	do
		FILE_NAME=`basename ${dir}`
		OUTPUT_PATH=${dir:${#INPUT}-${#dir}:${#dir}}
		mkdir -p ${OUTPUT}`dirname ${OUTPUT_PATH}`
		native2ascii -encoding UTF-8 -reverse ${dir} ${OUTPUT}/${OUTPUT_PATH}
	done

exit 0
