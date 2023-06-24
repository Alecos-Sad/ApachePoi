package by.sadovnick.exeloperations;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

public class ReadingExel {
    public static void main(String[] args) throws IOException {
        String exelFilePath = ".\\datafiles\\countries.xlsx";
        FileInputStream inputStream = new FileInputStream(exelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);//        workbook.getSheet("Лист1");

        //USING FOR LOOP
        int rows = sheet.getLastRowNum();
        int cols = sheet.getRow(1).getLastCellNum();
        readLoop(sheet, rows, cols);

        //USING ITERATOR
        readIterator(sheet);
    }

    public static void readIterator(XSSFSheet sheet) {
        for (Row cells : sheet) {
            XSSFRow row = (XSSFRow) cells;
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                XSSFCell cell = (XSSFCell) cellIterator.next();
                parseCellType(cell);
            }
            System.out.println();
        }
    }

    public static void readLoop(XSSFSheet sheet, int rows, int cols) {
        for (int r = 0; r <= rows; r++) {
            XSSFRow row = sheet.getRow(r);
            for (int c = 0; c < cols; c++) {
                XSSFCell cell = row.getCell(c);
                parseCellType(cell);
            }
            System.out.println();
        }
    }

    private static void parseCellType(@NotNull XSSFCell cell) {
        switch (cell.getCellType()) {
            case STRING:
                System.out.print(cell.getStringCellValue());
                break;
            case NUMERIC:
            case FORMULA:
                System.out.print(cell.getNumericCellValue());
                break;
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
        }
        System.out.print(" | ");
    }
}
