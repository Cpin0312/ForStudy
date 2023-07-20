/*
 * ファイル名　：CommonBaseDao.java
 * 作成者　　　：日立ソフト
 * 作成日　　　：2012/05/24
 * 変更履歴　　：
 *   日付        更新者　 　　 内容
 * 2012/05/24    日立ソフト　  初版作成
 */
package jp.hitachisoft.jfk.online.common.db.dao;

// Pattern
import java.util.Date;

import jp.hitachisoft.jfk.commons.pattern.exception.BaseException;

/**
 * クラス名 : CommonBaseDao
 * 機能概要 : 共通情報取得用DAOインターフェース
 */
public interface CommonBaseDao {

	/**
	 * 処理名称 : DBサーバー日時取得メソッド
	 * 機能 : DBサーバー日時を取得する。
	 * @return systimestamp String
	 * @throws BaseException - 処理中に予期せぬ例外が発生した場合に通知します。
	 */
	public Date selectTimestamp() throws BaseException;

}