#!/bin/sh
#===============================================================

# 作成者     :  日立ソリューションズ
# 作成日     :  2019/08/27
# 変更履歴   :
#    日付      更新者    内容
# 2019/08/27  日立ソリューションズ 初版
#===============================================================

CUR_PATH=`dirname ${0}`
INPUT=${CUR_PATH}
LIST_PROPERTIES=`find ${INPUT} -type f -name "*.${1}"`

# ファイル数でループする
for dir in ${LIST_PROPERTIES};
do
	rm -rf ${dir}
done

echo "Finish!!!"
# find "C:/Users/hisol/Desktop/java" -type f -name "*.bak" | xargs -i rm -f {} 

exit 0
