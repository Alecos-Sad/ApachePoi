package by.sadovnick.exeloperations;

import by.sadovnick.exeloperations.util.XLUtility;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

public class WevTableToExel {
    public static void main(String[] args) throws IOException {
        System.setProperty("webdriver.chrome.driver", "C:\\chrome driver\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();

        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        driver.manage().window().maximize();

        driver.get("https://en.wikipedia.org/wiki/List_of_countries_and_dependencies_by_population");

        XLUtility xlUtility = new XLUtility(MyPaths.WEB_FILE);

        //Write headers in exel sheet
        xlUtility.setCellData("Sheet1", 0, 0, "Country");
        xlUtility.setCellData("Sheet1", 0, 1, "Population");
        xlUtility.setCellData("Sheet1", 0, 2, "% of world");
        xlUtility.setCellData("Sheet1", 0, 3, "Date");
        xlUtility.setCellData("Sheet1", 0, 4, "Source");

        //Capture table
        WebElement table = driver.findElement(By.xpath("//table[@class='wikitable sortable jquery-tablesorter']/tbody"));
        int rows = table.findElements(By.xpath("tr")).size();//rows present in web table
        for (int r = 1; r <= rows; r++) {
            String country = table.findElement(By.xpath("tr[" + r + "]/td[1]")).getText();
            String population = table.findElement(By.xpath("tr[" + r + "]/td[2]")).getText();
            String perOfWorld = table.findElement(By.xpath("tr[" + r + "]/td[3]")).getText();
            String date = table.findElement(By.xpath("tr[" + r + "]/td[4]")).getText();
            String source = table.findElement(By.xpath("tr[" + r + "]/td[5]")).getText();

            System.out.println(country + " " + population + " " + perOfWorld + " " + date + " " + source);

            //writing data
            xlUtility.setCellData("Sheet1", r, 0, country);
            xlUtility.setCellData("Sheet1", r, 1, population);
            xlUtility.setCellData("Sheet1", r, 2, perOfWorld);
            xlUtility.setCellData("Sheet1", r, 3, date);
            xlUtility.setCellData("Sheet1", r, 4, source);
        }

        System.out.println("Web scrapping is done succssesfully!");
        driver.close();


    }
}
