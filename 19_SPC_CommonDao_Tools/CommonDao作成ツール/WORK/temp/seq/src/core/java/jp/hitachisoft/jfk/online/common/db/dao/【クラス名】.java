package jp.hitachisoft.jfk.online.common.db.dao;

import jp.hitachisoft.jfk.commons.pattern.exception.BaseException;

/**
 * �N���X�� : �y�N���X���z 
 * �@�\�T�v : �y�V�[�P���X���z�ɃA�N�Z�X����DAO�̃C���^�t�F�[�X
 */
public interface �y�N���X���z {

	/**
     * �y�V�[�P���X���z�擾
     * @return long �y�V�[�P���X���z
     * @throws BaseException - �������ɗ\�����ʗ�O�����������ꍇ�A�ʒm���܂��B
     */
	public long getSequenceNo() throws BaseException;

}
