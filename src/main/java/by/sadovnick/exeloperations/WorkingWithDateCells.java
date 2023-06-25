package by.sadovnick.exeloperations;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;

public class WorkingWithDateCells {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Date formats");

        //date in number format
        XSSFCell cell = sheet.createRow(0).createCell(0);
        cell.setCellValue(new Date());

        CreationHelper creationHelper = workbook.getCreationHelper();

        //format1: dd-mm-yyy
        CellStyle style1 = workbook.createCellStyle();
        style1.setDataFormat(creationHelper.createDataFormat().getFormat("dd-mm-yyyy"));

        XSSFCell cell1 = sheet.createRow(1).createCell(0);
        cell1.setCellStyle(style1);
        cell1.setCellValue(new Date());

        FileOutputStream outputStream = new FileOutputStream(MyPaths.DATE_FILE);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
        System.out.println("DONE!");
    }
}
