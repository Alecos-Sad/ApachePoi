package by.sadovnick.exeloperations;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

public class HashMapToExel {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Student data");

        Map<String, String> data = new HashMap<>();
        data.put("101", "John");
        data.put("102", "Smith");
        data.put("103", "Scott");
        data.put("104", "Kim");
        data.put("105", "Mary");

        int rowno = 0;

        for (Map.Entry<String, String> entry : data.entrySet()){
            XSSFRow row = sheet.createRow(rowno++);
            row.createCell(0).setCellValue(entry.getKey());
            row.createCell(1).setCellValue(entry.getValue());
        }

        FileOutputStream fos = new FileOutputStream(MyPaths.STUDENT_FILE);
        workbook.write(fos);

        workbook.close();
        fos.close();
        System.out.println("Done!");
    }
}
