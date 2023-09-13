# ForStudy
Code for Study

I have using here for my StudyKnowledge.Thanks


import java.util.ArrayList;
import java.util.List;

public class StringProcessor {

    public static List<String[]> generateOutput(String[]... inputLists) {
        List<String[]> outputList = new ArrayList<>();

        // リストの組み合わせを生成
        for (int i = 0; i < inputLists.length; i++) {
            for (int j = 0; j < inputLists[i].length; j++) {
                if (!inputLists[i][j].contains("@BLOCK@")) {
                    generateCombinations(inputLists, i, j, outputList);
                }
            }
        }

        return outputList;
    }

    private static void generateCombinations(String[][] inputLists, int index1, int index2, List<String[]> outputList) {
        for (int k = 0; k < inputLists.length; k++) {
            if (k != index1) {
                for (int l = 0; l < inputLists[k].length; l++) {
                    if (!inputLists[k][l].contains("@BLOCK@")) {
                        String[] result = new String[2];
                        result[0] = removeBlock(inputLists[index1][index2]);
                        result[1] = removeBlock(inputLists[k][l]);
                        outputList.add(result);
                    }
                }
            }
        }
    }

    private static String removeBlock(String str) {
        return str.replace("@BLOCK@", "");
    }

    public static void main(String[] args) {
        String[] aList = new String[]{"aa", "cc@BLOCK@"};
        String[] bList = new String[]{"bb", "dd@BLOCK@"};

        List<String[]> outputList = generateOutput(aList, bList);

        for (int i = 0; i < outputList.size(); i++) {
            System.out.print("String[" + i + "] = {");
            for (int j = 0; j < outputList.get(i).length; j++) {
                System.out.print("\"" + outputList.get(i)[j] + "\",");
            }
            System.out.println("}");
        }
    }
}