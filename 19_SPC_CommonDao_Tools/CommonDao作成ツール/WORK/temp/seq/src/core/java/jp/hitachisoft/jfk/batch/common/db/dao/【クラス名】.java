package jp.hitachisoft.jfk.batch.common.db.dao;

import java.sql.Connection;
import java.sql.ResultSet;
import java.sql.SQLException;
import jp.hitachisoft.jfk.batch.common.BatchCommonConst;
import jp.hitachisoft.jfk.batch.common.bean.BatchInputInfo;
import jp.hitachisoft.jfk.batch.common.dbaccess.BatchBaseDAO;
import jp.hitachisoft.jfk.batch.common.exception.BatchSystemException;
import jp.hitachisoft.jfk.batch.common.log.BatchLogInfo;
import jp.hitachisoft.jfk.batch.common.util.CommonConst;
import jp.hitachisoft.jfk.commons.common.LogMessageId;

/**
 * クラス名 : 【クラス名】
 * 機能概要 : 【シーケンス名】にアクセスするDAOの実装クラス
 */
public class 【クラス名】 extends BatchBaseDAO {

	private static Connection conn = null;
	/**
	 * クラス構造方法 <br>
	 */
	public 【クラス名】(BatchInputInfo inputInfo, Connection connection) {
		super(inputInfo, connection);
        conn = connection;
	}

	/**
	 * 【シーケンス名】取得 
	 * @return long 通番
	 * @throws BatchSystemException - 処理中に予期せぬ例外が発生した場合、通知します。
	 */
	public long getSequenceNo() throws BatchSystemException {
		long returnResult = 0L;

		ResultSet rs = null;
		// 検索実行し、結果を取得
		try {
			// 使用SQLID初期化
			String sqlId = null;
			// リソース名取得
			String resource = conn.getMetaData().getDatabaseProductName();
			// リソース名判定
			if (resource.equals(CommonConst.RDBMS_ORACLE)){
				// リソース名が"Oracle"の場合
				// Oracle用のSQLIDを設定
				sqlId =  "【SQLID(Oracle用)】";
			} else if (resource.equals(CommonConst.RDBMS_POSTGRESQL)){
				// リソース名が"PostgreSQL"の場合
				// PostgreSQLのSQLIDを設定
				sqlId =  "【SQLID(PostgreSQL用)】";
			}
			if (getPstmt() == null) {
				loadPreparedStatement(BatchCommonConst.SQLMAP_FILE, sqlId, null);
			}
			rs = getPstmt().executeQuery();
			while (rs.next()) {
				returnResult = rs.getLong("【シーケンスID】");
			}
		} catch (SQLException e) {

			BatchLogInfo logInfo = new BatchLogInfo(LogMessageId.MECO00006, e.getMessage(), e);  //ログ出力
            throw new BatchSystemException(logInfo);
		} finally {
			try {
				if (rs != null) {
					rs.close();
					rs = null;
				}
			} catch (SQLException e) {
				BatchLogInfo logInfo = new BatchLogInfo(LogMessageId.MECO00046,
						e.getMessage(), e);
				throw new BatchSystemException(logInfo);
			}
		}

		return returnResult;
	}
}
