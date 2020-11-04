package cn.itcast.poi.test;

import cn.itcast.poi.handler.SheetHandler;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

/**
 * 使用事件模型解析百万数据
 */
public class PoiTest6 {
    public static void main(String[] args) throws IOException, OpenXML4JException, SAXException {
        //1.创建工作簿 HSSFWorkBook 03版本的
        String path="E:\\test\\demo.xlsx";
        OPCPackage opcPackage = OPCPackage.open(path, PackageAccess.READ);
        XSSFReader reader=new XSSFReader(opcPackage);
        SharedStringsTable table = reader.getSharedStringsTable();
        StylesTable stylesTable = reader.getStylesTable();
        XMLReader xmlReader= XMLReaderFactory.createXMLReader();
        XSSFSheetXMLHandler xssfSheetXMLHandler = new XSSFSheetXMLHandler(stylesTable,table,new SheetHandler(),false);
        xmlReader.setContentHandler(xssfSheetXMLHandler);
         XSSFReader.SheetIterator sheetIterator = (XSSFReader.SheetIterator) reader.getSheetsData();

         while (sheetIterator.hasNext()){
             InputStream s = sheetIterator.next();
             InputSource inputSource = new InputSource(s);
              xmlReader.parse(inputSource);
         }
    }

}
