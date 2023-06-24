package by.sadovnick.exeloperations;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadDataFromFormulaCell {
    public static void main(String[] args) throws IOException {
        FileInputStream file = new FileInputStream(MyPaths.SALARY_FILE);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(0);
        int rows = sheet.getLastRowNum();
        int cols = sheet.getRow(0).getLastCellNum();

        ReadingExel.readLoop(sheet, rows, cols);
        file.close();
    }
}
