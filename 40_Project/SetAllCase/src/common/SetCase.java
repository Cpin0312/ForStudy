package common;

import java.util.Arrays;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class SetCase {

	public static void main(String[] args) {
		String[] aList = new String[] { "111", "A@BLOCK@" };
		String[] bList = new String[] { "B@BLOCK@", "222" };
		String[] cList = new String[] { "333", "C@BLOCK@", "444" };
		String[] dList = new String[] { "ZZZ", "QQQ", "SDF" };

		for (String str : setAllCase01(aList, bList, cList,dList)) {
			System.out.println(str);
		}
	}

	private static String[] setAllCase01(String[]... arrays) {
		int totalCase = 1;
		for (String[] array : arrays) {
			totalCase *= array.length;
		}
		String[] returnList = new String[totalCase];
		String[] lastList = null;
		;
		int runCase = 0;
		for (String[] array : arrays) {
			runCase = 0;
			lastList = returnList.clone();
			boolean isSame = false;
			while (runCase < totalCase) {
				for (String value : array) {
					String currentCase = returnList[runCase];
					String lastCloneCase = (runCase == 0) ? null : lastList[runCase - 1];
					if (currentCase == null) {
						// 何もしない
					} else if (isDoubleBlock(currentCase) || isDelete(currentCase)
							|| (isBlock(currentCase) && isBlock(value))) {
						value += "@DELETE@";
					} else if (isBlock(lastCloneCase) && isEquals(lastCloneCase, currentCase)) {
						if (isSame) {
							value += "@DELETE@";
						} else {
							isSame = true;
						}
					} else if (isBlock(value) && !isBlock(currentCase)) {
						value += "@REMAIN@";
					} else if (!isEquals(lastCloneCase, currentCase) && isBlock(currentCase)) {
						isSame = true;
					}
					if (currentCase == null || currentCase.isEmpty()) {
						returnList[runCase] = value;
					} else {
						returnList[runCase] += "@CRLF@" + value;
					}
					runCase++;
				}
			}
			Arrays.sort(returnList);
		}
		returnList = Arrays.stream(returnList).map(xx -> xx.replace("@CRLF@", ", ")).toArray(String[]::new);
		returnList = Arrays.stream(returnList).filter(xx -> !isDoubleBlock(xx) && !isDelete(xx)).toArray(String[]::new);
		returnList = Arrays.stream(returnList).map(xx -> xx.replace("@REMAIN@", "")).toArray(String[]::new);
		returnList = Arrays.stream(returnList).map(xx -> xx.replace("@BLOCK@", "")).toArray(String[]::new);

		return returnList;
	}

	private static boolean isEquals(String... strList) {
		String compareStr = strList[0];
		boolean returnBoolean = false;

		for (String str : strList) {
			if (compareStr == null || str == null) {
				returnBoolean = (compareStr == null && str == null) ? true : false;
			} else {
				returnBoolean = returnBoolean ? compareStr.equals(str) : true;
			}

		}
		return returnBoolean;
	}

	private static boolean isDoubleBlock(String str) {
		return isTarget(str, "@BLOCK@.*@BLOCK@");
	}

	private static boolean isRemain(String str) {
		return isTarget(str, "@REMAIN@");
	}

	private static boolean isBlock(String str) {
		return isTarget(str, "@BLOCK@");
	}

	private static boolean isDelete(String str) {
		return isTarget(str, "@DELETE@");
	}

	private static boolean isTarget(String str, String pattermStr) {
		if (str == null) {
			return false;
		}
		// 判定するパターンを生成
		Pattern p = Pattern.compile(pattermStr);
		Matcher m = p.matcher(str);
		return m.find();
	}

}
