package by.sadovnick.exeloperations;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.SQLException;
import java.sql.Statement;

public class ExelToDatabase {
    public static void main(String[] args) throws SQLException, IOException {
        //Database connection
        //connect to database
        Connection connection = DriverManager.getConnection("jdbc:mysql://localhost:3306/my_db", "root", "root");
        //statement/query
        Statement statement = connection.createStatement();
        //create a new database 'fromExel'
        String sql = "CREATE TABLE my_db.fromexel (\n" +
                "  Id int(11) NOT NULL AUTO_INCREMENT,\n" +
                "  name varchar(15) DEFAULT NULL,\n" +
                "  age int(11) DEFAULT NULL,\n" +
                "  PRIMARY KEY (Id)\n" +
                ")";
        statement.execute(sql);

        //Exel
        FileInputStream fileInputStream = new FileInputStream(MyPaths.DATABASE_FILE);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        int rows = sheet.getLastRowNum();
        for (int r = 1; r <= rows; r++) {
            XSSFRow row = sheet.getRow(r);
            int id = (int) row.getCell(0).getNumericCellValue();
            String name = row.getCell(1).getStringCellValue();
            int age = (int) row.getCell(2).getNumericCellValue();

            sql = "insert into my_db.fromexel values ('" + id + "', '" + name + "', '" + age + "' )";
            statement.execute(sql);
            statement.execute("commit ");
        }
        workbook.close();
        fileInputStream.close();
        connection.close();
        System.out.println("OK");
    }
}
