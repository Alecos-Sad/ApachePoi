package by.sadovnick.exeloperations;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteFormulaCell {
    public static void main(String[] args) throws IOException {
        File file = new File(MyPaths.CALC_FILE);
        FileOutputStream outputStream = new FileOutputStream(file);

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Numbers");
        XSSFRow row = sheet.createRow(0);

        row.createCell(0).setCellValue(100);
        row.createCell(1).setCellValue(200);
        row.createCell(2).setCellValue(300);

        row.createCell(3).setCellFormula("A1*B1*C1");

        workbook.write(outputStream);
        outputStream.close();

        System.out.println("calc.xlsx is created");

    }
}
