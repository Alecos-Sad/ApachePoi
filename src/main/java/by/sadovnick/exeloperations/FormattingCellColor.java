package by.sadovnick.exeloperations;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileOutputStream;
import java.io.IOException;

public class FormattingCellColor {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet1");
        XSSFRow row = sheet.createRow(1);

        //Setting background color
        XSSFCellStyle style = workbook.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);

        XSSFCell cell = row.createCell(1);
        cell.setCellValue("welcome");
        cell.setCellStyle(style);

        //Setting Foreground color
        style = workbook.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        cell = row.createCell(2);
        cell.setCellValue("Automation");
        cell.setCellStyle(style);

        FileOutputStream fos = new FileOutputStream(MyPaths.COLOR_FILE);
        workbook.write(fos);
        workbook.close();
        fos.close();

        System.out.println("Done!");
    }
}
