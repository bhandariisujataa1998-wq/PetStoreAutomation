package api.utilities;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.*;

public class XLUtility {

    private XSSFWorkbook workbook;
    private String path;

    public XLUtility(String path) throws IOException {
        this.path = path;
        try (FileInputStream fi = new FileInputStream(path)) {
            this.workbook = new XSSFWorkbook(fi);
        }
    }

    public int getRowCount(String sheetName) {
        XSSFSheet sheet = workbook.getSheet(sheetName);
        return (sheet != null) ? sheet.getLastRowNum() : 0;
    }

    public int getCellCount(String sheetName, int rowNum) {
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) return 0;
        XSSFRow row = sheet.getRow(rowNum);
        return (row != null) ? row.getLastCellNum() : 0;
    }

    public String getCellData(String sheetName, int rowNum, int colNum) {
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) return "";

        XSSFRow row = sheet.getRow(rowNum);
        if (row == null) return "";

        XSSFCell cell = row.getCell(colNum);
        if (cell == null) return "";

        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    public void close() throws IOException {
        if (workbook != null) {
            workbook.close();
        }
    }
}

/** package api.utilities;

import java.io.FileInputStream; import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter; import
org.apache.poi.xssf.usermodel.XSSFCell; import
org.apache.poi.xssf.usermodel.XSSFRow; import
org.apache.poi.xssf.usermodel.XSSFSheet; import
org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLUtility {

public FileInputStream fi;

public XSSFWorkbook workbook; public XSSFSheet sheet; public XSSFRow row;
public XSSFCell cell;

String path;

public XLUtility(String path) { this.path = path; }

public int getRowCount(String sheetname) throws IOException { fi= new
FileInputStream(path); workbook = new XSSFWorkbook(fi); sheet =
workbook.getSheet(sheetname); int rowcount = sheet.getLastRowNum();
workbook.close(); fi.close(); return rowcount;

}

public int getCellCount(String sheetname, int rownum) throws IOException {
fi= new FileInputStream(path); workbook = new XSSFWorkbook(fi); sheet =
workbook.getSheet(sheetname); row = sheet.getRow(rownum); int cellcount =
row.getLastCellNum(); workbook.close(); fi.close(); return cellcount; }

public String getCellData(String sheetname, int rownum, int colnum) throws
IOException

{ fi= new FileInputStream(path); workbook = new XSSFWorkbook(fi); sheet =
workbook.getSheet(sheetname); row = sheet.getRow(rownum); cell =
row.getCell(colnum);

DataFormatter formatter = new DataFormatter(); String data; try { data =
formatter.formatCellValue(cell); } catch(Exception e) { data = ""; }
workbook.close(); fi.close(); return data;

}

}

**/