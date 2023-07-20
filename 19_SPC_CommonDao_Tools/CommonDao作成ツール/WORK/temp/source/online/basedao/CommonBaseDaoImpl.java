/*
 * �t�@�C�����@�FCommonBaseDaoImpl.java
 * �쐬�ҁ@�@�@�F�����\�t�g
 * �쐬���@�@�@�F2012/05/24
 * �ύX�����@�@�F
 *   ���t        �X�V�ҁ@ �@�@ ���e
 * 2012/05/24    �����\�t�g�@  ���ō쐬
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
 * �N���X�� : CommonBaseDaoImpl
 * �@�\�T�v : ���ʏ��擾�pDAO�����N���X
 */
public class CommonBaseDaoImpl extends JFKBaseDao implements
		CommonBaseDao {


	/**
	 * �N���X�\�����@ <br>
	 */
	public CommonBaseDaoImpl() throws BaseException {
		super();
	}

	/**
	 * �������� : DB�T�[�o�[�����擾���\�b�h
	 * �@�\ : DB�T�[�o�[�������擾����B
	 * @return systimestamp String
	 * @throws BaseException - �������ɗ\�����ʗ�O�����������ꍇ�A�ʒm���܂��B
	 */
	public Date selectTimestamp() throws BaseException {
		// �g�pSQLID������
		String sqlId = null;
		try {
			// ���\�[�X���擾
			String resource = DataSourceManager.getRDBMSName();
			// ���\�[�X������
			if (resource.equals(CommonConst.RDBMS_ORACLE)){
				// ���\�[�X����"Oracle"�̏ꍇ
				// Oracle�p��SQLID��ݒ�
				sqlId = CommonConst.SELECT_TIMESTAMP_ORACLE;
			} else if (resource.equals(CommonConst.RDBMS_POSTGRESQL)){
				// ���\�[�X����"PostgreSQL"�̏ꍇ
				// PostgreSQL��SQLID��ݒ�
				sqlId = CommonConst.SELECT_TIMESTAMP_POSTGRESQL;
			}
			// SQL���𐶐����ADB�������s��
			Date systimestamp = (Date)getSqlMapClientTemplate().queryForObject(
					sqlId, "");
			// iBatis�̃��\�b�h���s��ɂ���1�s��ǉ����܂�
			super.edit(1, sqlId, "");
			// ���ʒl�ԋp
			return systimestamp;
		} catch (Exception e) {
			// catch�����ۂɂ�����1�s��ǉ����܂�
			super.errorEdit(0, sqlId, "");
			// ���R�[�h���o�����̃��OID���p�����^�Ƃ��āAlogInfo�𐶐�����
			LogInfo logInfo = new LogInfo(LogMessageId.MECO00006,
					e.getMessage());
			// �V�X�e��Exception�𐶐�
			SystemException se = new SystemException(logInfo,
					MessageId.ERROR_E9009);
			// throw
			throw se;
		}
	}
}