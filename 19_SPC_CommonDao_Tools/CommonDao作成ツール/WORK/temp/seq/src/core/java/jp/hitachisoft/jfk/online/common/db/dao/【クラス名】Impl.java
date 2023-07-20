package jp.hitachisoft.jfk.online.common.db.dao;

import jp.hitachisoft.jfk.commons.common.LogMessageId;
import jp.hitachisoft.jfk.commons.common.MessageId;
import jp.hitachisoft.jfk.commons.common.db.JFKBaseDao;
import jp.hitachisoft.jfk.commons.common.log.LogInfo;
import jp.hitachisoft.jfk.commons.pattern.exception.BaseException;
import jp.hitachisoft.jfk.commons.pattern.exception.SystemException;
import jp.hitachisoft.jfk.online.common.util.CommonConst;
import jp.hitachisoft.jfk.online.common.util.DataSourceManager;

/**
 * クラス名 : 【クラス名】Impl
 * 機能概要 : 【シーケンス名】にアクセスするDAOクラス
 */
public class 【クラス名】Impl extends JFKBaseDao implements 【クラス名】 {

	/**
	 * 【シーケンス名】取得
	 * @return long 【シーケンス名】
	 * @throws BaseException - 処理中に予期せぬ例外が発生した場合、通知します。
	 */
	public long getSequenceNo() throws BaseException {
		Long resultLong = 0L;
		// 使用SQLID初期化
		String sqlId = null;
		try {
			// リソース名取得
			String resource = DataSourceManager.getRDBMSName();
			// リソース名判定
			if (resource.equals(CommonConst.RDBMS_ORACLE)){
				// リソース名が"Oracle"の場合
				// Oracle用のSQLIDを設定
				sqlId = "【SQLID(Oracle用)】";
			} else if (resource.equals(CommonConst.RDBMS_POSTGRESQL)){
				// リソース名が"PostgreSQL"の場合
				// PostgreSQLのSQLIDを設定
				sqlId = "【SQLID(PostgreSQL用)】";
			}

			resultLong = (Long)getSqlMapClientTemplate().queryForObject(sqlId);

			super.edit(1, sqlId, null);

		} catch (Exception e) {
			// catchした際にもこの1行を追加します
			super.errorEdit(0, sqlId, null);
			// レコード抽出処理のログIDをパラメタとして、logInfoを生成する
			LogInfo logInfo = new LogInfo(LogMessageId.MECO00006,
					e.getMessage());
			// システムExceptionを生成
			SystemException se = new SystemException(logInfo,
					MessageId.ERROR_E9009);
			// throwします
			throw se;
		}
		return resultLong;
	}

}
