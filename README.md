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