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
 * �N���X�� : �y�N���X���zImpl
 * �@�\�T�v : �y�V�[�P���X���z�ɃA�N�Z�X����DAO�N���X
 */
public class �y�N���X���zImpl extends JFKBaseDao implements �y�N���X���z {

	/**
	 * �y�V�[�P���X���z�擾
	 * @return long �y�V�[�P���X���z
	 * @throws BaseException - �������ɗ\�����ʗ�O�����������ꍇ�A�ʒm���܂��B
	 */
	public long getSequenceNo() throws BaseException {
		Long resultLong = 0L;
		// �g�pSQLID������
		String sqlId = null;
		try {
			// ���\�[�X���擾
			String resource = DataSourceManager.getRDBMSName();
			// ���\�[�X������
			if (resource.equals(CommonConst.RDBMS_ORACLE)){
				// ���\�[�X����"Oracle"�̏ꍇ
				// Oracle�p��SQLID��ݒ�
				sqlId = "�ySQLID(Oracle�p)�z";
			} else if (resource.equals(CommonConst.RDBMS_POSTGRESQL)){
				// ���\�[�X����"PostgreSQL"�̏ꍇ
				// PostgreSQL��SQLID��ݒ�
				sqlId = "�ySQLID(PostgreSQL�p)�z";
			}

			resultLong = (Long)getSqlMapClientTemplate().queryForObject(sqlId);

			super.edit(1, sqlId, null);

		} catch (Exception e) {
			// catch�����ۂɂ�����1�s��ǉ����܂�
			super.errorEdit(0, sqlId, null);
			// ���R�[�h���o�����̃��OID���p�����^�Ƃ��āAlogInfo�𐶐�����
			LogInfo logInfo = new LogInfo(LogMessageId.MECO00006,
					e.getMessage());
			// �V�X�e��Exception�𐶐�
			SystemException se = new SystemException(logInfo,
					MessageId.ERROR_E9009);
			// throw���܂�
			throw se;
		}
		return resultLong;
	}

}
