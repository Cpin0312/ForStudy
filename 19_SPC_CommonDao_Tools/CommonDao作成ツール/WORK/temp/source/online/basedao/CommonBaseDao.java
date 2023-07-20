/*
 * �t�@�C�����@�FCommonBaseDao.java
 * �쐬�ҁ@�@�@�F�����\�t�g
 * �쐬���@�@�@�F2012/05/24
 * �ύX�����@�@�F
 *   ���t        �X�V�ҁ@ �@�@ ���e
 * 2012/05/24    �����\�t�g�@  ���ō쐬
 */
package jp.hitachisoft.jfk.online.common.db.dao;

// Pattern
import java.util.Date;

import jp.hitachisoft.jfk.commons.pattern.exception.BaseException;

/**
 * �N���X�� : CommonBaseDao
 * �@�\�T�v : ���ʏ��擾�pDAO�C���^�[�t�F�[�X
 */
public interface CommonBaseDao {

	/**
	 * �������� : DB�T�[�o�[�����擾���\�b�h
	 * �@�\ : DB�T�[�o�[�������擾����B
	 * @return systimestamp String
	 * @throws BaseException - �������ɗ\�����ʗ�O�����������ꍇ�ɒʒm���܂��B
	 */
	public Date selectTimestamp() throws BaseException;

}