package by.sadovnick.exeloperations;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class WriteFormulaCell2 {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream(MyPaths.BOOK_FILE);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);
        sheet.getRow(7).getCell(2).setCellFormula("СУММ(C2:C6)");
        fis.close();

        FileOutputStream fos = new FileOutputStream(MyPaths.BOOK_FILE);
        workbook.write(fos);
        workbook.close();
        fos.close();
        System.out.println("DONE!");
    }
}
