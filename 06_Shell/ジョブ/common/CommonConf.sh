#!/bin/bash
#===============================================================
#
# システム名 :  SPC西武
# ジョブ名   :  CommonConf.sh
# ジョブ名称 :  共通変数
# ファイルタイプ  :  ?B-shell
# 実行形式        :  CommonConf.sh
#
# リターンコード : なし
#
# 作成者     :  日立ソリューションズ
# 作成日     :  2019/08/27
# 変更履歴   :
#    日付      更新者    内容
# 2019/08/27  日立ソリューションズ 初版
#===============================================================

#===============================================================
#
# 共通処理 : DB系
#
#===============================================================
# DBホスト
DB_HOST="@DB_HOST@"
# DBポート
DB_PORT="@DB_PORT@"
# DB名
DB_NAME="@DB_NAME@"
# DBユーザー
DB_USER="@DB_UID@"
# DBパスワード
DB_PASSWORD="@DB_PWD@"

#===============================================================
#
# 共通処理 : ファイル受信系
#
#===============================================================
#CTR
# SFTPホスト
SFTP_CTR_HOST="@SFTP_CTR_HOST@"
# SFTPユーザー
SFTP_CTR_USER="@SFTP_CTR_USER@"
# 秘密鍵パス
SFTP_CTR_SECRETKEY="@SFTP_CTR_KEY_PATH@"

#CCMP
# SFTPホスト
SFTP_CCMP_HOST="@SFTP_CCMP_HOST@"
# SFTPユーザー
SFTP_CCMP_USER="@SFTP_CCMP_USER@"
# 秘密鍵パス
SFTP_CCMP_SECRETKEY="@SFTP_CCMP_KEY_PATH@"

# SPIRAL
# SFTPホスト
SFTP_SPIRAL_HOST="@SFTP_SPIRAL_HOST@"
# SFTPユーザー
SFTP_SPIRAL_USER="@SFTP_SPIRAL_USER@"
# 秘密鍵パス
SFTP_SPIRAL_SECRETKEY="@SFTP_SPIRAL_KEY_PATH@"
