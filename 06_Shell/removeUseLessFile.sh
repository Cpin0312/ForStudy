#!/bin/sh
#===============================================================

# �쐬��     :  �����\�����[�V�����Y
# �쐬��     :  2019/08/27
# �ύX����   :
#    ���t      �X�V��    ���e
# 2019/08/27  �����\�����[�V�����Y ����
#===============================================================

CUR_PATH=`dirname ${0}`
INPUT=${CUR_PATH}
LIST_PROPERTIES=`find ${INPUT} -type f -name "*.${1}"`

# �t�@�C�����Ń��[�v����
for dir in ${LIST_PROPERTIES};
do
	rm -rf ${dir}
done

echo "Finish!!!"
# find "C:/Users/hisol/Desktop/java" -type f -name "*.bak" | xargs -i rm -f {} 

exit 0
