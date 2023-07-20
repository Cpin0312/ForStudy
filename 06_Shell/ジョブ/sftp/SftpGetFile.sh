#!/bin/sh
#===============================================================
#
# システム名 :  SPC西武
# ジョブ名   :  SftpGetFile.sh
# ジョブ名称 :  Java実行sシェル
# ファイルタイプ  :  ?B-shell
# 実行形式        :  SftpGetFile.sh ジョブID
#
# リターンコード : 0(正常),失敗(9),警告(1),タイムアウト(3)
#
# 作成者     :  日立ソリューションズ
# 作成日     :  2019/09/09
# 変更履歴   :
#    日付      更新者    内容
# 2019/09/09  日立ソリューションズ 初版
#===============================================================

# ファイル取得のコマンド
CMD_GET=
CMD_GET="${CMD_GET}get "
CMD_GET="${CMD_GET}${SFTP_MOTO_FILEPATH} "
CMD_GET="${CMD_GET}${SFTP_SAKI_FILEPATH} "
CMD_GET="${CMD_GET}\r "

# ファイル削除
CMD_REMOVE=
CMD_REMOVE="${CMD_REMOVE}rm "
CMD_REMOVE="${CMD_REMOVE}${SFTP_MOTO_FILEPATH} "
CMD_REMOVE="${CMD_REMOVE}\r "

# 終了のコマンド
CMD_END=
CMD_END="${CMD_END}bye"
# 戻り値
errorlevel=9

# *----入力パラメタを設定---------------------
INPUT_JOB_ID=${1}

# *----実行開始---------------------------------------------------------------
echo ${INPUT_JOB_ID} 実行開始...

# 開始ログ
LOG_START

LOG_PRINT "${CMD_GET}"

# 下記のインストールが必要です
# yum install expect
expect -c "
        # タイムアウトの設定
        set timeout ${TIMEOUT_SEC}
        # SFTP接続
        spawn sftp  -i ${SFTP_KEY_PATH} ${SFTP_USER}@${SFTP_HOST}
        # 自動応答とファイル取得コマンドの入力、タイムアウト発生のキャッチ
        expect {
            \"*sftp> \"  { send  \"${CMD_GET}\" }
            default { return 5 }
        }
        # 対象削除
        expect {
            \"*sftp> \"  { send  \"${CMD_REMOVE}\" }
            default { return 5 }
        }
        # 自動応答とExpect終了コマンドの入力、タイムアウト発生のキャッチ
        expect {
            \"*sftp> \"  { send \"${CMD_END}\" }
            default { return 5 }
        }
        expect {
            \"bye\"  { send \"return 0\" }
            default { return 5 }
        }
        # 成功の戻り値を返却
        return 0
"

# 戻り値の設定
errorlevel=$?

#- 結果の出力 ---+---------+---------+---------+
# 改行のため
echo
if [ ${errorlevel} -eq 0 ]; then
    echo  ${INPUT_JOB_ID} 処理正常終了...
    LOG_PRINT "移動元パス : ${SFTP_MOTO_FILEPATH}"
    LOG_PRINT "移動先パス : ${SFTP_SAKI_FILEPATH}"
    # 終了ログ
    LOG_STOP
else
    if [ ${errorlevel} -eq 5 ]; then
        echo  ${INPUT_JOB_ID} 処理タイムアウト終了...
        # 終了ログ
        LOG_TIME_OUT
    else
        echo  ${INPUT_JOB_ID} 処理異常終了...
        # 終了ログ
        LOG_ERROR
    fi
fi

# echo "戻り値 : ${errorlevel}"
return ${errorlevel}
