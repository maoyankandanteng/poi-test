package cn.itcast.poi.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiTest {
    public static void main(String[] args) throws IOException {
        //1.创建工作簿 HSSFWorkBook 03版本的
        Workbook wb=new XSSFWorkbook();
        //2.创建表单sheet
        Sheet sheet = wb.createSheet("test");
        //参数：行的索引index
        Row row = sheet.createRow(2);
        //参数：列的索引index
        Cell cell = row.createCell(2);
        cell.setCellValue("北京浩坤");
        //样式
        CellStyle cellStyle=wb.createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.DASH_DOT);
        cellStyle.setBorderLeft(BorderStyle.DASH_DOT);
        cellStyle.setBorderRight(BorderStyle.DASH_DOT);
        cellStyle.setBorderTop(BorderStyle.DASH_DOT);



        //字体
        Font font = wb.createFont();
        font.setFontName("华文行楷");
        font.setFontHeightInPoints((short) 28);
        cellStyle.setFont(font);

        cell.setCellStyle(cellStyle);

        //行高
        row.setHeight((short) (50*20));
        //设置列宽
        sheet.setColumnWidth(2,31*256);


        //居中显示
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        //3.文件流
        FileOutputStream fileOutputStream=new FileOutputStream("E:\\test\\test.xls");
        //4.写入文件
        wb.write(fileOutputStream);
        fileOutputStream.close();
    }
}
