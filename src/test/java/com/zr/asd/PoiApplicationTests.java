package com.zr.asd;


import com.zr.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;

@RunWith(SpringRunner.class)
@SpringBootTest
public class PoiApplicationTests {

    @Test
    public void contextLoads() {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //创建工作表
        HSSFSheet sheet = workbook.createSheet("测试");

        //创建第一行
        HSSFRow row = sheet.createRow(0);
        //创建第一个单元格
        HSSFCell cell = row.createCell(0);
        //给单元格填数据
        cell.setCellValue("第一个单元格");

        try {
            workbook.write(new FileOutputStream(new File("C:/Users/acer/Desktop/a.xls")));
            System.out.println("创建成功");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    @Test
    public void contextLoads1() {
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //设置日期格式
        HSSFDataFormat dataFormat = workbook.createDataFormat();
        short format = dataFormat.getFormat("yyyy-MM-dd");
        //将日期格式设置进单元格样式里
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(format);
        //创建一个新单元格样式

        HSSFCellStyle cellStyle1 = workbook.createCellStyle();
        //字体居中
        cellStyle1.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontName("楷体");
        font.setColor(Font.COLOR_RED);
        font.setItalic(true);

        cellStyle1.setFont(font);

        //创建工作表
        HSSFSheet sheet = workbook.createSheet("测试");
        //设置第三列的宽度为15像素
        sheet.setColumnWidth(2, 15 * 256);
        //先创键标题行
        HSSFRow row = sheet.createRow(0);
        String[] str = {"id", "姓名", "生日"};
        for (int i = 0; i < str.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(cellStyle1);
            cell.setCellValue(str[i]);
        }

        ArrayList<User> users = new ArrayList<>();
        User user = new User("1", "htf", new Date());
        User user1 = new User("2", "cpx", new Date());
        User user2 = new User("3", "rxx", new Date());
        users.add(user);
        users.add(user1);
        users.add(user2);

        for (int i = 0; i < users.size(); i++) {
            HSSFRow row1 = sheet.createRow(i + 1);
            HSSFCell cell = row1.createCell(0);
            int i1 = user.hashCode();
            System.out.println(i1);
            cell.setCellValue(users.get(i).getId());
            HSSFCell cell1 = row1.createCell(1);
            cell1.setCellValue(users.get(i).getName());
            HSSFCell cell2 = row1.createCell(2);
            //将单元格样式运用在此单元格
            cell2.setCellStyle(cellStyle);
            cell2.setCellValue(users.get(i).getBir());
        }
        try {
            workbook.write(new FileOutputStream(new File("C:/Users/acer/Desktop/a.xls")));
            System.out.println("创建成功");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    @Test
    public void contextLoads3() throws IOException {
        HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(new File("C:/Users/acer/Desktop/a.xls")));
        HSSFSheet sheet = workbook.getSheet("测试");
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 1; i <= lastRowNum; i++) {
            HSSFRow row = sheet.getRow(i);
            HSSFCell cell = row.getCell(0);
            String id = cell.getStringCellValue();
            HSSFCell cell1 = row.getCell(1);
            String name = cell1.getStringCellValue();
            HSSFCell cell2 = row.getCell(2);
            Date bir = cell2.getDateCellValue();
            User user = new User(id, name, bir);
            System.out.println(user);
        }
    }

}
