#!/bin/bash
#===============================================================
#
# システム名 :  SPC西武
# ジョブ名   :  FileAction.sh
# ジョブ名称 :  ファイル移動処理
# ファイルタイプ  :  ?B-shell
# 実行形式        :  FILEACTION_JOB ジョブID
#
# リターンコード : なし
#
# 作成者     :  日立ソリューションズ
# 作成日     :  2019/10/04
# 変更履歴   :
#    日付      更新者    内容
# 2019/10/04  日立ソリューションズ 初版
#===============================================================

#- returnCode ---------+---------+---------+---------+---------+
errorLevel=9


#- ジョブID   ---------+---------+---------+---------+---------+
INPUT_JOB_ID=${1}

#- ファイルアクション             ---------+---------+---------+
F_ACTION=${FILE_ACTION}
#- ファイル名 ---------+---------+---------+---------+---------+
F_NAME=${FA_FILE_NAME}
#- ファイル拡張子                 ---------+---------+---------+
F_EXTENTION=${FA_EXTENTION}
#- ファイル拡張子                 ---------+---------+---------+
F_FILE_WITH_EXTENTION=${F_NAME}${FA_EXTENTION}
#- ディレクトリ元パス             ---------+---------+---------+
F_MOTO_DIR=`SetPath ${FA_MOTO_PATH}`
#- ディレクトリ先ファイル         ---------+---------+---------+
F_SAKI_DIR=`SetPath ${FA_SAKI_PATH}`
#- BKディレクトリファイル         ---------+---------+---------+
F_BACKUP_DIR=`SetPath ${FA_BACKUP_PATH}`

#- ディレクトリ元パスファイル     ---------+---------+---------+
F_MOTO_FILE_PATH=${F_MOTO_DIR}${F_FILE_WITH_EXTENTION}
#- ディレクトリ先パスファイル     ---------+---------+---------+
F_SAKI_FILE_PATH=${F_SAKI_DIR}${F_FILE_WITH_EXTENTION}
#- BKファイルパスファイル     ---------+---------+---------+
F_BACKUP_FILE_PATH="${F_BACKUP_DIR}${F_NAME}/${F_FILE_WITH_EXTENTION}_`date '+%Y%m%d%H%M%S'`"

# *----実行開始-------------------------------------------------
echo ${INPUT_JOB_ID} 実行開始...

# 開始ログ
LOG_START

# 移動元が存在しない場合
if [ ! -e ${F_MOTO_FILE_PATH} ]; then

    LOG_PRINT "移動元が存在しません。パス:${F_MOTO_FILE_PATH}"

else

	# ファイルコピー・移動の場合
	if [ ${F_ACTION} = "COPY" -o ${F_ACTION} = "MOVE" ]; then

		# 移動先ディレクトリが存在しない場合
		if [ ! -e ${F_SAKI_DIR} ]; then

    		LOG_PRINT "移動先のディレクトリが存在しません。パス:${F_SAKI_DIR}"
		else
			if [ ${F_ACTION} = "COPY" ]; then

       			LOG_PRINT "cp -f ${F_MOTO_FILE_PATH}  ${F_SAKI_FILE_PATH}"
				# ファイルコピー(強制上書き)
				cp -f ${F_MOTO_FILE_PATH}  ${F_SAKI_FILE_PATH}
				errorLevel=$?

				if [ ${errorLevel} -eq 0 ]; then
       				LOG_PRINT "ファイル:[${F_MOTO_FILE_PATH}]をファイル:[${F_SAKI_FILE_PATH}]にコピーしました。"
       			fi

			else

       			LOG_PRINT "mv -f ${F_MOTO_FILE_PATH}  ${F_SAKI_FILE_PATH}"
				# ファイル移動(強制上書き)
				mv -f ${F_MOTO_FILE_PATH}  ${F_SAKI_FILE_PATH}
				errorLevel=$?

				if [ ${errorLevel} -eq 0 ]; then
       				LOG_PRINT "ファイル:[${F_MOTO_FILE_PATH}]をファイル:[${F_SAKI_FILE_PATH}]に移動しました。"
       			fi
			fi

		fi

	# ファイルバックアップの場合
	elif [ ${F_ACTION} = "BACKUP" -o ${F_ACTION} = "MOVE_TO_BACKUP" ]; then

		# バックアップ先ディレクトリが存在しない場合
		if [ ! -e ${F_BACKUP_DIR} ]; then

    		LOG_PRINT "バックアップ先のディレクトリが存在しません。パス:${F_BACKUP_DIR}"

		else

			if [ ! -e  ${F_BACKUP_DIR}/${F_NAME} ]; then

				LOG_PRINT "mkdir -p ${F_BACKUP_DIR}/${F_NAME}"
				# バックアップディレクトリを作成
				mkdir -p ${F_BACKUP_DIR}/${F_NAME}
				errorLevel=$?

				if [ ${errorLevel} -eq 0 ]; then
       				LOG_PRINT "バックアップ先のディレクトリを作成しました。"
       			fi

			fi

			if [ ${errorLevel} -eq 0 ]; then
				if [ ${F_ACTION} = "BACKUP" ]; then

					LOG_PRINT "cp -f ${F_MOTO_FILE_PATH}  ${F_BACKUP_FILE_PATH}"
					# ファイルバックアップ
					cp -f ${F_MOTO_FILE_PATH}  ${F_BACKUP_FILE_PATH}
					errorLevel=$?

					if [ ${errorLevel} -eq 0 ]; then
	       				LOG_PRINT "ファイル:[${F_MOTO_FILE_PATH}]をファイル:[${F_BACKUP_FILE_PATH}]にコピーしました。"
	       			fi

				else

					LOG_PRINT "mv -f ${F_MOTO_FILE_PATH}  ${F_BACKUP_FILE_PATH}"
					# ファイルをバックアップフォルダに移動
					mv -f ${F_MOTO_FILE_PATH}  ${F_BACKUP_FILE_PATH}
					errorLevel=$?

					if [ ${errorLevel} -eq 0 ]; then
	       				LOG_PRINT "ファイル:[${F_MOTO_FILE_PATH}]をファイル:[${F_BACKUP_FILE_PATH}]に移動しました。"
	       			fi

				fi
       		fi
		fi

	# ファイル移動とバックアップの場合
	elif [ ${F_ACTION} = "MOVE_AND_BACKUP" ]; then

		# 移動先ディレクトリが存在しない場合
		if [ ! -e ${F_SAKI_DIR} ]; then

    		LOG_PRINT "移動先のディレクトリが存在しません。パス:${F_SAKI_DIR}"

		# バックアップ先ディレクトリが存在しない場合
		elif [ ! -e ${F_BACKUP_DIR} ]; then

    		LOG_PRINT "バックアップ先のディレクトリが存在しません。パス:${F_BACKUP_DIR}"

		else

			if [ ! -e  ${F_BACKUP_DIR}/${F_NAME} ]; then

				LOG_PRINT "mkdir -p ${F_BACKUP_DIR}/${F_NAME}"
				# バックアップディレクトリを作成
				mkdir -p ${F_BACKUP_DIR}/${F_NAME}
				errorLevel=$?

				if [ ${errorLevel} -eq 0 ]; then
       				LOG_PRINT "バックアップ先のディレクトリを作成しました。"
       			fi

			fi

			# ファイルバックアップ
			LOG_PRINT "cp -f ${F_MOTO_FILE_PATH}  ${F_BACKUP_FILE_PATH}"
			cp -f ${F_MOTO_FILE_PATH}  ${F_BACKUP_FILE_PATH}
			errorLevel=$?

			if [ ${errorLevel} -eq 0 ]; then
				LOG_PRINT "ファイル:[${F_MOTO_FILE_PATH}]をファイル:[${F_BACKUP_FILE_PATH}]にコピーしました。"
				# ファイル移動
				LOG_PRINT "mv -f ${F_MOTO_FILE_PATH}  ${F_SAKI_FILE_PATH}"
				mv -f ${F_MOTO_FILE_PATH}  ${F_SAKI_FILE_PATH}
				errorLevel=$?
	       		LOG_PRINT "ファイル:[${F_MOTO_FILE_PATH}]をファイル:[${F_SAKI_FILE_PATH}]に移動しました。"
   			fi
		fi
	# ファイル削除の場合
	elif [ ${F_ACTION} = "REMOVE" ]; then

		# 移動先ディレクトリが存在しない場合
		if [ ! -e ${F_MOTO_FILE_PATH} ]; then

    		LOG_PRINT "削除元のファイルが存在しません。パス:${F_MOTO_FILE_PATH}"
		else

			# ファイル削除
			LOG_PRINT "rm -f ${F_MOTO_FILE_PATH}"
			rm -f ${F_MOTO_FILE_PATH}
			errorLevel=$?
       		LOG_PRINT "ファイル:[${F_MOTO_FILE_PATH}]を削除しました。"

		fi
	else

		LOG_PRINT "処理対象外です。アクション:${F_ACTION}"
	fi

fi

#- 結果の出力 ---+---------+---------+---------+
if [ ${errorLevel} -eq 0 ]; then
    echo ${INPUT_JOB_ID} 処理実行終了...
    # 終了ログ
    LOG_STOP
else
    echo ${INPUT_JOB_ID} 処理異常終了...
    errorLevel=9
    # 異常ログ
    LOG_ERROR
fi

# echo 終了コード:${errorLevel}
return ${errorLevel}
