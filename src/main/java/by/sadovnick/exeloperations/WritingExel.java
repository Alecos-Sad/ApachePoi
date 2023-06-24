package by.sadovnick.exeloperations;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jetbrains.annotations.NotNull;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

//WorkBook --> Sheet --> Rows --> Cells
public class WritingExel {
    public static void main(String[] args) throws IOException {
        String exelFilePath = MyPaths.EMPLOYEE_FILE;
        FileOutputStream outputStream = new FileOutputStream(exelFilePath);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Emp Info");

        ArrayList<Object[]> empdataArray = new ArrayList<>();
        fillArray(empdataArray);

        Object[][] empdata = getObjects();

        writeWithForeEachLoopFromArray(sheet, empdataArray);


        workbook.write(outputStream);
        outputStream.flush();
        outputStream.close();
        System.out.println("Employee.xlsx file written successfully");
    }

    private static void writeWithForeEachLoopFromArray(XSSFSheet sheet, ArrayList<Object[]> empdataArray) {
        int rownum = 0;
        for (Object[] emp : empdataArray){
            XSSFRow row = sheet.createRow(rownum++);
            int cellnum = 0;
            writeRowData(emp, row, cellnum);
        }
    }

    private static void writeRowData(Object[] emp, XSSFRow row, int cellnum) {
        for (Object value : emp){
            XSSFCell cell = row.createCell(cellnum++);
            checkInstance(value, cell);
        }
    }

    private static void writeWithForeEachLoop(XSSFSheet sheet, Object[][] empdata) {
        int rowCount = 0;
        for (Object[] emp : empdata) {
            XSSFRow row = sheet.createRow(rowCount++);
            int columnCount = 0;
            writeRowData(emp, row, columnCount);
        }
    }

    private static void fillArray(ArrayList<Object[]> empdataArray) {
        empdataArray.add(new Object[]{"EmpID", "Name", "Job"});
        empdataArray.add(new Object[]{101, "David", "Engineer"});
        empdataArray.add(new Object[]{102, "Smith", "Manager"});
        empdataArray.add(new Object[]{103, "Scott", "Analyst"});
    }

    @NotNull
    private static Object[][] getObjects() {
        Object[][] empdata = {
                {"EmpID", "Name", "Job"},
                {101, "David", "Engineer"},
                {102, "Smith", "Manager"},
                {103, "Scott", "Analyst"}
        };
        return empdata;
    }

    private static void checkInstance(Object value, XSSFCell cell) {
        if (value instanceof String) {
            cell.setCellValue((String) value);
        }
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        }
        if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        }
    }
}
