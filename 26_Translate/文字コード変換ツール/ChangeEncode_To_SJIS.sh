#!/bin/sh
#===============================================================

# 作成者     :  日立ソリューションズ
# 作成日     :  2019/08/27
# 変更履歴   :
#    日付      更新者    内容
# 2019/08/27  日立ソリューションズ 初版
#===============================================================


CUR_PATH=`dirname ${0}`

bash ${CUR_PATH}/ChangeEncode_Common.sh SJIS ${1}

exit 0
