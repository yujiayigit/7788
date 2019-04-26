package poi;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;

public class test {
    public static void main(String[] args) throws IOException {

        new For().xunhuan();
    }
}


class Read {
    public String getCellValue(String title, Cell cell, int cellType, String cellValue) {
        switch (cellType) {
            case Cell.CELL_TYPE_STRING: //字符串类型
                cellValue = cell.getStringCellValue().trim();
                cellValue = cellValue.isEmpty() ? "空" : cellValue;
                break;
            case Cell.CELL_TYPE_BOOLEAN:  //布尔类型
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_NUMERIC: //数值类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {  //判断日期类型
                    Date dateCellValue = cell.getDateCellValue();
                    cellValue = new SimpleDateFormat("yyyy-MM-dd hh:mm").format(dateCellValue);
                } else {  //否
                    cellValue = new DecimalFormat("#.######").format(cell.getNumericCellValue());
                }
                break;
            default: //其它类型，取空串吧
                cellValue = "空";
                break;
        }
        //System.out.println(title + "=" + cellValue);
        return cellValue;
    }

    public void asd() {


    }

}

class For {
    public void xunhuan() throws IOException {


        String yiJi = "D:\\Documents\\Tencent Files\\409095785\\FileRecv\\MobileFile\\一级目录表.xlsx";
        String ziBiao = "D:\\Documents\\Tencent Files\\409095785\\FileRecv\\MobileFile\\二级目录表.xlsx";
        String muLu = "D:\\Documents\\Tencent Files\\409095785\\FileRecv\\MobileFile\\卷组目录表.xlsx";

        Read read = new Read();

        FileInputStream muLuFile = new FileInputStream(muLu);
        Workbook workbook = null;
        //判断excel的两种格式xls,xlsx
        if (muLu.toLowerCase().endsWith("xlsx")) {
            workbook = new XSSFWorkbook(muLuFile);
        } else if (muLu.toLowerCase().endsWith("xls")) {
            workbook = new HSSFWorkbook(muLuFile);
        }

        //對Excel的讀取
        // 創建對Excel工作簿文件的引用

        // 創建對工作表的引用
        Sheet sheet = workbook.getSheet("卷组目录表");
        //HSSFSheet sheet = workbook.getSheetAt("卷组目录表");//讀取第一張工作表 Sheet1
        short lastCellNum = sheet.getRow(0).getLastCellNum();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                for (int j = 0; j < lastCellNum; j++) {
                    Cell cell = row.getCell(j);
                    String cellValue = "";
                    int cellType = 0;
                    if (cell != null) {
                        cellType = cell.getCellType();
                    } else {
                        String title = sheet.getRow(0).getCell(j).getStringCellValue();
                        continue;
                    }
                    String title = sheet.getRow(0).getCell(j).getStringCellValue();
                    //String cellValue1 = read.getCellValue(title, cell, cellType, cellValue);

                }
            } else {
                continue;
            }

            //

            FileInputStream yijiFile = new FileInputStream(yiJi);
            Workbook workbook2 = null;
            //判断excel的两种格式xls,xlsx
            if (yiJi.toLowerCase().endsWith("xlsx")) {
                workbook2 = new XSSFWorkbook(yijiFile);
            } else if (yiJi.toLowerCase().endsWith("xls")) {
                workbook2 = new HSSFWorkbook(yijiFile);
            }

            //對Excel的讀取
            // 創建對Excel工作簿文件的引用

            // 創建對工作表的引用
            Sheet sheet2 = workbook2.getSheet("一级目录表");
            //HSSFSheet sheet = workbook.getSheetAt("卷组目录表");//讀取第一張工作表 Sheet1
            short lastCellNum2 = sheet2.getRow(0).getLastCellNum();
            int index = 1;
            for (int t = 1; t <= sheet2.getLastRowNum(); t++) {
                Row row2 = sheet2.getRow(t);
                if (row2 != null) {
                    for (int j = 0; j < lastCellNum2; j++) {
                        Cell cell = row2.getCell(j);
                        String cellValue = "";
                        int cellType = 0;
                        if (cell != null) {
                            cellType = cell.getCellType();
                        } else {
                            String title = sheet.getRow(0).getCell(j).getStringCellValue();
                            continue;
                        }
                        String title = sheet2.getRow(0).getCell(j).getStringCellValue();
                        String cellValue1 = read.getCellValue(title, cell, cellType, cellValue);
                    }
                } else {
                    continue;
                }
                System.out.println("一级表"+t+"次");
                //
                FileInputStream ziBiaoFile = new FileInputStream(ziBiao);
                Workbook workbook3 = null;
                //判断excel的两种格式xls,xlsx
                if (ziBiao.toLowerCase().endsWith("xlsx")) {
                    workbook3 = new XSSFWorkbook(ziBiaoFile);
                } else if (ziBiao.toLowerCase().endsWith("xls")) {
                    workbook3 = new HSSFWorkbook(ziBiaoFile);
                }

                //對Excel的讀取
                // 創建對Excel工作簿文件的引用

                // 創建對工作表的引用
                Sheet sheet3 = workbook3.getSheet("二级目录表");
                //HSSFSheet sheet = workbook.getSheetAt("卷组目录表");//讀取第一張工作表 Sheet1
                short lastCellNum3 = sheet3.getRow(0).getLastCellNum();


                for (int c = index; c <= sheet3.getLastRowNum(); c++) {
                    Row row3 = sheet3.getRow(c);
                    String cell1 = row3.getCell(1).getStringCellValue();
                    String cell2 = row2.getCell(0).getStringCellValue();
                    if (!cell1.equals(cell2)) {
                        continue;
                    }
                    if (row3 != null && cell1.equals(cell2)) {

                        for (int a = 0; a < lastCellNum3; a++) {
                            Cell cell = row3.getCell(a);
                            String cellValue = "";
                            int cellType = 0;
                            if (cell != null) {
                                cellType = cell.getCellType();
                            } else {
                                String title = sheet3.getRow(0).getCell(a).getStringCellValue();
                                continue;
                            }
                            String title = sheet3.getRow(0).getCell(a).getStringCellValue();
                            String cellValue1 = read.getCellValue(title, cell, cellType, cellValue);
                        }

                    } else {
                        index=c+1;
                        break;
                    }
                    System.out.println("子表"+c+"次");
                }

                //

            }
            //
            System.out.println("目录表"+i+"次");
        }
    }
}