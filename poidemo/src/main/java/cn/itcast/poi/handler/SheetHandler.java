package cn.itcast.poi.handler;

import cn.itcast.poi.entity.PoiEntity;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

//逐行读取
public class SheetHandler implements XSSFSheetXMLHandler.SheetContentsHandler {
    private PoiEntity entity;

    /**
     * 开始解析某一行
     *
     * @param i
     */
    public void startRow(int i) {
        if (i > 0) {
            entity = new PoiEntity();
        }
    }

    /**
     * 结束解析某一行
     *
     * @param i
     */
    public void endRow(int i) {
        //使用对象进行业务操作
        if(entity!=null){

            System.out.println(entity.toString());
        }
    }

    /**
     * 对行中的单元格处理
     *
     * @param cellReference 单元格名称
     * @param value
     * @param xssfComment   批注
     */
    public void cell(String cellReference, String value, XSSFComment xssfComment) {
        if (entity != null) {
            String pre = cellReference.substring(0, 1);
            switch (pre) {
                case "A":
                    entity.setId(value);
                    break;
                case "B":
                    entity.setBreast(value);
                    break;
                case "C":
                    entity.setAdipocytes(value);
                    break;
                case "D":
                    entity.setNegative(value);
                    break;
                case "E":
                    entity.setStaining(value);
                    break;
                case "F":
                    entity.setSupportive(value);
                    break;
                default:
                    break;
            }
        }
    }
}
