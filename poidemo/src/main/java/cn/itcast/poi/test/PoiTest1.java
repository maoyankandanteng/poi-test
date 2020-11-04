package cn.itcast.poi.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiTest1 {
    public static void main(String[] args) throws IOException {
        //1.创建工作簿 HSSFWorkBook 03版本的
        Workbook wb=new XSSFWorkbook();
        //2.创建表单sheet
        Sheet sheet = wb.createSheet("test");
        //读取图片流
        FileInputStream fileInputStream=new FileInputStream("E:\\test\\logo.jpg");

        //转化为二进制数组
        byte[] bytes = IOUtils.toByteArray(fileInputStream);
        fileInputStream.read(bytes);

        //向poi内存中添加一张图片,返回图片在图片集合中的索引
        int i = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
        //绘制图片工具类，
        CreationHelper creationHelper = wb.getCreationHelper();

        //创建绘图对象
        Drawing<?> drawingPatriarch = sheet.createDrawingPatriarch();

        //创建锚点
        ClientAnchor clientAnchor = creationHelper.createClientAnchor();
        clientAnchor.setRow1(0);
        clientAnchor.setCol1(0);
//        clientAnchor.setRow2();

        Picture picture = drawingPatriarch.createPicture(clientAnchor, i);
        picture.resize();//自适应图片大小


        //3.文件流
        FileOutputStream fileOutputStream=new FileOutputStream("E:\\test\\test1.xls");
        //4.写入文件
        wb.write(fileOutputStream);
        fileOutputStream.close();
    }
}
