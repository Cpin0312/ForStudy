# ForStudy
Code for Study

I have using here for my StudyKnowledge.Thanks

import java.util.ArrayList;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        String[] aList = new String[]{"a", "@X@"};
        String[] bList = new String[]{"b", "@X@"};
        String[] cList = new String[]{"c", "d", "e"};

        List<String[]> allCombinations = new ArrayList<>();
        generateCombinations(new String[][]{aList, bList, cList}, new String[aList.length], allCombinations);

        // 出力
        for (String[] combination : allCombinations) {
            if (isValidCombination(combination)) {
                System.out.println(String.join(", ", combination));
            }
        }
    }

    public static void generateCombinations(String[][] arrays, String[] current, List<String[]> result) {
        int index = result.size();
        if (index == arrays.length) {
            result.add(current.clone());
            return;
        }

        String[] currentList = arrays[index];
        for (String item : currentList) {
            if (!containsItem(current, item) && (item.equals("@X@") || index == 0 || current[index - 1].equals("@X@"))) {
                current[index] = item;
                generateCombinations(arrays, current, result);
                current[index] = null; // リセットして次の要素を試す
            }
        }
    }

    public static boolean containsItem(String[] array, String item) {
        for (String str : array) {
            if (item.equals(str)) {
                return true;
            }
        }
        return false;
    }

    public static boolean isValidCombination(String[] combination) {
        int xCount = 0;
        for (String item : combination) {
            if (item.equals("@X@")) {
                xCount++;
                if (xCount > 1) {
                    return false;
                }
            }
        }
        return true;
    }
}
