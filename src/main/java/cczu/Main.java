package cczu;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class Main {
    private static final String PATH = "工程文件/";
    private static final String file_Name = "石工213王钱龙的快速会议-412480991-601cba7c5083.xlsx";
    private static final int start_Row = 9;
    public static void main(String[] args) throws IOException {
        //文件初始化
        ArrayList<String> init = tools.init();

        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(PATH + file_Name);
        //获取一个工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        //获取一个工作表
        XSSFSheet sheetAt = workbook.getSheetAt(0);
        //获取行数
        int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows() +1;
        System.out.println(physicalNumberOfRows);
        for (int i = start_Row; i < physicalNumberOfRows; i++) {
            //获取开始行
            XSSFRow row = sheetAt.getRow(i);
            //获取单元格
            XSSFCell cell = row.getCell(0);
            String s = cell.toString();
            System.out.println(s);
            for (int j = 0; j < init.size(); j++) {
                boolean judge = tools.judge(s, init.get(j));
                if (judge){
                    init.remove(j);
                    j--;
                }
            }
        }

        System.out.print("请输入第几次培训:");
        Scanner sc = new Scanner(System.in);
        int next = sc.nextInt();
        String new_file_Name = "第"+next+"次培训.xlsx";
        //写入缺席名单
        tools.write(init,new_file_Name);
        //结束
        System.out.println("全部结束！");
    }
}