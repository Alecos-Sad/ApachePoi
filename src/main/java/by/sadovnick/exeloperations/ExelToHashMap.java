package by.sadovnick.exeloperations;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;

public class ExelToHashMap {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream(MyPaths.STUDENT_FILE);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rows = sheet.getLastRowNum();

        HashMap<String, String> data = new HashMap<>();
        for (int r = 0; r <= rows; r++) {
            String key = sheet.getRow(r).getCell(0).getStringCellValue();
            String value = sheet.getRow(r).getCell(1).getStringCellValue();
            data.put(key, value);
        }

        data.forEach((k,v) -> System.out.println("key: " + k + " value: " + v));
    }
}
