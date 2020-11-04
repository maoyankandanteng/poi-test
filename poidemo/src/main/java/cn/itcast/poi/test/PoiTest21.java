package cn.itcast.poi.test;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class PoiTest21 {
    public static void main(String[] args) throws IOException {
        //1.创建工作簿 HSSFWorkBook 03版本的
        Workbook wb=new XSSFWorkbook("E:\\test\\demo.xlsx");
        //2.创建表单sheet
        Sheet sheet = wb.getSheetAt(0);
        for(int rowNum=0;rowNum<=sheet.getLastRowNum();rowNum++){
            Row row = sheet.getRow(rowNum);
            StringBuilder sb=new StringBuilder();
            for(int cellNum=2;cellNum<row.getLastCellNum();cellNum++){
                Cell cell = row.getCell(cellNum);
                Object value=getCellValue(cell);
                sb.append(value.toString());
            }
            System.out.println(sb.toString());
        }



        //3.文件流
        FileOutputStream fileOutputStream=new FileOutputStream("E:\\test\\test1.xls");
        //4.写入文件
        wb.write(fileOutputStream);
        fileOutputStream.close();
    }
    public static Object getCellValue(Cell cell){
        CellType cellType = cell.getCellType();
        Object value=null;
        switch (cellType){
            case STRING:
                value=cell.getStringCellValue();
                break;
            case BOOLEAN:
                value=cell.getBooleanCellValue();
                break;
            case NUMERIC:
                if(DateUtil.isCellDateFormatted(cell)){
                    value=cell.getDateCellValue();
                }else{
                    value=cell.getNumericCellValue();
                }
                break;
            case FORMULA:
                value=cell.getCellFormula();
                break;
            default:break;
        }
        return value;
    }
}
