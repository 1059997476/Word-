package Doc;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STHighlightColor;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Scanner;

/**
 * @Parmae 第一个输入为根文件夹 \172042 如：C:\Users\胡雍杰\Desktop\172042
 * @Parmae 第二个输入为自己文件名 如：17202120-胡雍杰-调查问卷.docx
 * @Parmae 第三个输入为题目数量
 * ·
 * ·
 * ·
 * @Return 每一题有一个map,里面存着所选个数
 *
 * @Tip 不支持文本格式如下
 *          1）答案在同一行
 *          2）不支持非word文档
 */
public class Main {
    static Map<String,HashMap<String,Integer>> map = new HashMap<>();
    public static void main(String[] args) throws IOException, XmlException {
        Scanner sc = new Scanner(System.in);
        System.out.println("请输入文件目录");
        String directory = sc.next();
        System.out.println("请输入文件名称（包含文件后缀）");//
        String fileName = sc.next();
        System.out.println("请输入题目数");
        int num = sc.nextInt();

        for (int i = 0; i < num; i++) {
            System.out.print("请输入第"+(i+1)+"题选项个数");
            int count = sc.nextInt();
            HashMap<String, Integer> hashMap = new HashMap<>();
            map.put(""+(i+1),hashMap);
            for (int j = 0; j < count; j++) {
                hashMap.put(String.valueOf((char)(65+j)),0);
            }
        }
        getFiles(new File(directory),fileName);
        System.out.println(map);
    }
    private static void getFiles(File root,String name){
        if(root.isDirectory()){
            File[] files = root.listFiles();
            for (File file:files)
                getFiles(file,name);
        }else {
            if(root.toString().endsWith(name))
                try {
                    tj(root);
                } catch (Exception e) {
                    e.printStackTrace();
                }
        }
    }
    private static void tj(File fileName) throws IOException{
        System.out.println("当前文件名:"+fileName.toString().substring(0,fileName.toString().lastIndexOf("\\")));
        FileInputStream file = new FileInputStream(fileName);
        XWPFDocument doc = new XWPFDocument(file);
        List<XWPFParagraph> paragraphs = doc.getParagraphs();
        HashMap<String,Integer> maps = null;
        for (XWPFParagraph xw:paragraphs){
            String text = xw.getText();
            String index = text.substring(0, 1);
            if(map.containsKey(index)){
                maps = map.get(index);
                System.out.print(index+"题选:");
            }else {
                List<XWPFRun> runs = xw.getRuns();
                STHighlightColor.Enum color = STHighlightColor.NONE;
                String textColor = runs.get(0).getColor();
                for (XWPFRun run:runs){
                    if(run.getTextHightlightColor()!=STHighlightColor.NONE)
                        color = run.getTextHightlightColor();
                }
                if("0000FF".equals(textColor)||"FF0000".equals(textColor)||color==STHighlightColor.YELLOW||color==STHighlightColor.RED||color==STHighlightColor.GREEN&&maps!=null) {
                    Integer integer = maps.get(index);
                    System.out.print(index);
                    maps.put(index,integer+1);
                }
            }
        }
        System.out.println("\n-------------------------------");
    }
}
