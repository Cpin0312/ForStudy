# ForStudy
Code for Study

I have using here for my StudyKnowledge.Thanks


2023年9月14日
import java.util.ArrayList;
import java.util.List;

public class StringProcessor {

    public static List<String[]> generateOutput(String[] aList, String[] bList) {
        List<String[]> outputList = new ArrayList<>();
        
        for (String a : aList) {
            for (String b : bList) {
                if (!a.contains("@BLOCK@") && !b.contains("@BLOCK@")) {
                    String[] result = new String[2];
                    result[0] = a;
                    result[1] = b;
                    outputList.add(result);
                }
            }
        }

        return outputList;
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


2023年9月14日
import java.util.ArrayList;
import java.util.List;

public class StringProcessor {

    public static List<String[]> generateOutput(String[]... inputLists) {
        List<String[]> outputList = new ArrayList<>();
        int numLists = inputLists.length;

        for (int i = 0; i < numLists; i++) {
            for (int j = 0; j < inputLists[i].length; j++) {
                if (!inputLists[i][j].contains("@BLOCK@")) {
                    for (int k = 0; k < numLists; k++) {
                        if (k != i) {
                            for (int l = 0; l < inputLists[k].length; l++) {
                                if (!inputLists[k][l].contains("@BLOCK@")) {
                                    String[] result = new String[2];
                                    result[0] = inputLists[i][j];
                                    result[1] = inputLists[k][l];
                                    outputList.add(result);
                                }
                            }
                        }
                    }
                }
            }
        }

        return outputList;
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