#!/bin/sh
#===============================================================
#
# システム名 :  SPC西武
# ジョブ名   :  PisJava.sh
# ジョブ名称 :  Java実行sシェル
# ファイルタイプ  :  ?B-shell
# 実行形式        :  PisJava.sh ジョブID
#
# リターンコード : 0(正常),失敗(9),警告(1)
#
# 作成者     :  日立ソリューションズ
# 作成日     :  2019/08/27
# 変更履歴   :
#    日付      更新者    内容
# 2019/08/27  日立ソリューションズ 初版
#===============================================================

#- Lib ---+---------+---------+---------+
batchPath=/var/app/batch
CLASS_PATH=.
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/activation.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/commons-beanutils-1.8.0.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/commons-codec-1.4.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/commons-digester.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/commons-httpclient-3.1.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/commons-io-1.4.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/commons-lang.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/commons-logging.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/commons-validator.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/cool.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/dom4j-1.6.1.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/ibatis-2.3.0.677.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/jakarta-oro-2.0.8.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/jfk.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/log4j-1.2.13.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/mail.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/message.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/ojdbc6.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/pis-common-dao.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/poi-3.2-FINAL-20081019.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/postgresql-42.2.5.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/servlet-api.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spring-aop-5.0.5.RELEASE.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spring-beans-5.0.5.RELEASE.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spring-context-5.0.5.RELEASE.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spring-core-5.0.5.RELEASE.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spring-expression-5.0.5.RELEASE.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spring-jdbc-5.0.5.RELEASE.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spring-tx-5.0.5.RELEASE.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spc-common.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/spc_code.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/im4java-1.4.0.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/lib/ant.jar
CLASS_PATH=${CLASS_PATH}:${batchPath}/bin

# *----入力パラメタを設定---------------------
INPUT_JOB_ID=${1}
declare -a OPTPARAM=(${@:2:($#-1)})

# *----生成したJavaアプリケーションのパッケージを指定する---------------------
GYOUMU_OPT=${GYOMU_PARAMS}

# *----生成したJavaアプリケーションのメインクラスを設定する-------------------
BATCH_PKG=jp.hitachisoft.pis.batch.main


# *----生成したJavaアプリケーションを実行する引数（オプション）を指定する。---
r=${RERUN}
c=${TRANSACTION_COUNT}
m=${PROC_ID}
d=${BATCH_DATE}
t=${TORIHIKI_DATE_CODE}
j=${JOB_ID}

# *----手動実行の場合、入力値を優先にする
for val in ${OPTPARAM[@]};
do
    if [ ${val:0:2} = "r:" ]; then
        r=${val:2}
    elif [ ${val:0:2} = "c:" ]; then
        c=${val:2}
    elif [ ${val:0:2} = "m:" ]; then
        m=${val:2}
    elif [ ${val:0:2} = "d:" ]; then
        d=${val:2}
    elif [ ${val:0:2} = "t:" ]; then
        t=${val:2}
    elif [ ${val:0:2} = "j:" ]; then
        j=${val:2}
    fi
done

BAT_OPTION="r:${r} c:${c} m:${m} d:${d} t:${t} j:${j}"

# *----実行開始---------------------------------------------------------------
echo ${KINOU_ID} 実行開始...

# 開始ログ
LOG_START ${KINOU_ID}

echo 実行パラメータ：${BAT_OPTION}

java -cp ${CLASS_PATH} ${BATCH_PKG}.${KINOU_ID} ${BAT_OPTION} ${GYOUMU_OPT}
errorlevel=$?

#- 結果の出力 ---+---------+---------+---------+
if [ ${errorlevel} -eq 0 ]; then
    echo  ${KINOU_ID} バッチ正常終了...
    # 終了ログ
    LOG_STOP ${KINOU_ID}
else
    if [ ${errorlevel} -eq 1 ]; then
        echo  ${KINOU_ID} バッチ警告終了...
        # 終了ログ
        LOG_WARNING ${KINOU_ID}
    else
        echo  ${KINOU_ID} バッチ異常終了...
        # 終了ログ
        LOG_ERROR ${KINOU_ID}
    fi
fi


# echo 終了コード:${errorlevel}
return ${errorlevel}
