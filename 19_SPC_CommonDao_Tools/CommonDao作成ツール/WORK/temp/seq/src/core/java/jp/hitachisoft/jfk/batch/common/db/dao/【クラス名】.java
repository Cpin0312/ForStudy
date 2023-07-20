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
 * �N���X�� : �y�N���X���z
 * �@�\�T�v : �y�V�[�P���X���z�ɃA�N�Z�X����DAO�̎����N���X
 */
public class �y�N���X���z extends BatchBaseDAO {

	private static Connection conn = null;
	/**
	 * �N���X�\�����@ <br>
	 */
	public �y�N���X���z(BatchInputInfo inputInfo, Connection connection) {
		super(inputInfo, connection);
        conn = connection;
	}

	/**
	 * �y�V�[�P���X���z�擾 
	 * @return long �ʔ�
	 * @throws BatchSystemException - �������ɗ\�����ʗ�O�����������ꍇ�A�ʒm���܂��B
	 */
	public long getSequenceNo() throws BatchSystemException {
		long returnResult = 0L;

		ResultSet rs = null;
		// �������s���A���ʂ��擾
		try {
			// �g�pSQLID������
			String sqlId = null;
			// ���\�[�X���擾
			String resource = conn.getMetaData().getDatabaseProductName();
			// ���\�[�X������
			if (resource.equals(CommonConst.RDBMS_ORACLE)){
				// ���\�[�X����"Oracle"�̏ꍇ
				// Oracle�p��SQLID��ݒ�
				sqlId =  "�ySQLID(Oracle�p)�z";
			} else if (resource.equals(CommonConst.RDBMS_POSTGRESQL)){
				// ���\�[�X����"PostgreSQL"�̏ꍇ
				// PostgreSQL��SQLID��ݒ�
				sqlId =  "�ySQLID(PostgreSQL�p)�z";
			}
			if (getPstmt() == null) {
				loadPreparedStatement(BatchCommonConst.SQLMAP_FILE, sqlId, null);
			}
			rs = getPstmt().executeQuery();
			while (rs.next()) {
				returnResult = rs.getLong("�y�V�[�P���XID�z");
			}
		} catch (SQLException e) {

			BatchLogInfo logInfo = new BatchLogInfo(LogMessageId.MECO00006, e.getMessage(), e);  //���O�o��
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
