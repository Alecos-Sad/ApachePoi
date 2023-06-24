package by.sadovnick.exeloperations;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class ReadingPasswordProtectedExel {
    public static void main(String[] args) throws IOException {
        FileInputStream fis = new FileInputStream(MyPaths.CUSTOMERS_FILE);
        //Пароль на лист в экселе
        String password = "123";

        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(fis, password);
        XSSFSheet sheet = workbook.getSheetAt(0);
        ReadingExel.readIterator(sheet);
        workbook.close();
        fis.close();
    }
}
