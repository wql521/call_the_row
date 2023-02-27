package cczu;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class tools {
    private tools(){} //私有化构造方法
    public static boolean judge(String s,String rule){
        //1.获取正则表达式对象
        Pattern compile = Pattern.compile(rule);
        //2.获取文本匹配器对象
        Matcher matcher = compile.matcher(s);
        //3.获取匹配结果
        return matcher.find();
    }

    public static ArrayList<String> init() throws IOException {
        String path = "/Users/wangqianlong/Desktop/All/IntelliJ IDEA/call_the_roll/工程文件/";
        String fileName = "武进西太湖干事.xlsx";
        ArrayList<String> all_people = new ArrayList<>();
        //获取文件流
        FileInputStream fileInputStream = new FileInputStream(path + fileName);
        //获取工作簿
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        //获取工作表
        XSSFSheet sheetAt = workbook.getSheetAt(0);
        //获取行
        int physicalNumberOfRows = sheetAt.getPhysicalNumberOfRows();
        for (int i = 0; i < physicalNumberOfRows; i++) {
            //获取行
            XSSFRow row = sheetAt.getRow(i);
            //获取单元格
            XSSFCell cell = row.getCell(0);
            String s = cell.toString();
            all_people.add(s);
        }
        return all_people;
    }

    //创建并且写入文件
    public static void write(ArrayList<String> all_People,String fileName) throws IOException {
        String path = "工程文件/";
        //创建工作簿
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建一个工作表
        XSSFSheet sheet = workbook.createSheet("缺席名单");
        //遍历列表
        for (int i = 0; i < all_People.size(); i++) {
            String s = all_People.get(i).toString();
            //获取行
            XSSFRow row = sheet.createRow(i);
            //创建一个单元格
            XSSFCell cell = row.createCell(0);
            //写入值
            cell.setCellValue(s);
        }

        //输出文件流
        FileOutputStream fileOutputStream = new FileOutputStream(path + fileName);
        //写入文件
        workbook.write(fileOutputStream);
        //关闭文件流
        fileOutputStream.close();
        System.out.println("写入完毕");
    }
}
