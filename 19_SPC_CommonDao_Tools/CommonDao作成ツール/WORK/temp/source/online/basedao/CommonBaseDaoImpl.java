/*
 * ファイル名　：CommonBaseDaoImpl.java
 * 作成者　　　：日立ソフト
 * 作成日　　　：2012/05/24
 * 変更履歴　　：
 *   日付        更新者　 　　 内容
 * 2012/05/24    日立ソフト　  初版作成
 */
package jp.hitachisoft.jfk.online.common.db.dao;

// Util
import java.util.Date;

import jp.hitachisoft.jfk.commons.common.LogMessageId;
import jp.hitachisoft.jfk.commons.common.MessageId;
import jp.hitachisoft.jfk.commons.common.db.JFKBaseDao;
import jp.hitachisoft.jfk.commons.common.log.LogInfo;
import jp.hitachisoft.jfk.commons.pattern.exception.BaseException;
import jp.hitachisoft.jfk.commons.pattern.exception.SystemException;
import jp.hitachisoft.jfk.online.common.util.CommonConst;
import jp.hitachisoft.jfk.online.common.util.DataSourceManager;

/**
 * クラス名 : CommonBaseDaoImpl
 * 機能概要 : 共通情報取得用DAO実装クラス
 */
public class CommonBaseDaoImpl extends JFKBaseDao implements
		CommonBaseDao {


	/**
	 * クラス構造方法 <br>
	 */
	public CommonBaseDaoImpl() throws BaseException {
		super();
	}

	/**
	 * 処理名称 : DBサーバー日時取得メソッド
	 * 機能 : DBサーバー日時を取得する。
	 * @return systimestamp String
	 * @throws BaseException - 処理中に予期せぬ例外が発生した場合、通知します。
	 */
	public Date selectTimestamp() throws BaseException {
		// 使用SQLID初期化
		String sqlId = null;
		try {
			// リソース名取得
			String resource = DataSourceManager.getRDBMSName();
			// リソース名判定
			if (resource.equals(CommonConst.RDBMS_ORACLE)){
				// リソース名が"Oracle"の場合
				// Oracle用のSQLIDを設定
				sqlId = CommonConst.SELECT_TIMESTAMP_ORACLE;
			} else if (resource.equals(CommonConst.RDBMS_POSTGRESQL)){
				// リソース名が"PostgreSQL"の場合
				// PostgreSQLのSQLIDを設定
				sqlId = CommonConst.SELECT_TIMESTAMP_POSTGRESQL;
			}
			// SQL文を生成し、DB検索を行う
			Date systimestamp = (Date)getSqlMapClientTemplate().queryForObject(
					sqlId, "");
			// iBatisのメソッド実行後にこの1行を追加します
			super.edit(1, sqlId, "");
			// 結果値返却
			return systimestamp;
		} catch (Exception e) {
			// catchした際にもこの1行を追加します
			super.errorEdit(0, sqlId, "");
			// レコード抽出処理のログIDをパラメタとして、logInfoを生成する
			LogInfo logInfo = new LogInfo(LogMessageId.MECO00006,
					e.getMessage());
			// システムExceptionを生成
			SystemException se = new SystemException(logInfo,
					MessageId.ERROR_E9009);
			// throw
			throw se;
		}
	}
}