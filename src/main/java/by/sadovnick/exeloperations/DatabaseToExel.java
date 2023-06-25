package by.sadovnick.exeloperations;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.*;

public class DatabaseToExel {
    public static void main(String[] args) throws SQLException, IOException {
        //connect to database
        Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/my_db", "root", "root");
        //statement/query
        Statement statement = connection.createStatement();
        ResultSet resultSet = statement.executeQuery("select * from children");

        //Exel
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Data");
        XSSFRow row = sheet.createRow(0);
        row.createCell(0).setCellValue("ID");
        row.createCell(1).setCellValue("Name");
        row.createCell(2).setCellValue("Age");

        int r = 1;
        while (resultSet.next()) {
            int id = resultSet.getInt("ID");
            String name = resultSet.getString("name");
            int age = resultSet.getInt("age");

            row = sheet.createRow(r++);
            row.createCell(0).setCellValue(id);
            row.createCell(1).setCellValue(name);
            row.createCell(2).setCellValue(age);
        }

        FileOutputStream outputStream = new FileOutputStream(MyPaths.DATABASE_FILE);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
        connection.close();
        System.out.println("Complete!");
    }
}
