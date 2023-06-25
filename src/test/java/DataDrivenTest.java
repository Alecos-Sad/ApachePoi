import by.sadovnick.exeloperations.MyPaths;
import by.sadovnick.exeloperations.util.XLUtility;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.IOException;
import java.util.concurrent.TimeUnit;

//https://www.youtube.com/watch?v=1nP9UlwzpgU&list=PLUDwpEzHYYLsN1kpIjOyYW6j_GLgOyA07&index=10
public class DataDrivenTest {
    private WebDriver driver;

    @BeforeClass
    public void setup() {
        System.setProperty("webdriver.chrome.driver", "C:\\chrome driver\\chromedriver.exe");
        driver = new ChromeDriver();
        driver.manage().timeouts().implicitlyWait(5, TimeUnit.SECONDS);
        driver.manage().window().maximize();
    }

    @Test(dataProvider = "LoginData")
    public void loginTest(String user, String pwd, String exp) {
        driver.get("https://admin-demo.nopcommerce.com/login");
        WebElement txtEmail = driver.findElement(By.id("Email"));
        txtEmail.clear();
        txtEmail.sendKeys(user);

        WebElement txtPassword = driver.findElement(By.id("Password"));
        txtPassword.clear();
        txtPassword.sendKeys(pwd);

        driver.findElement(By.xpath("//button[normalize-space()='Log in']")).click(); //Login button

        String exp_title = "Dashboard / nopCommerce administration";
        String act_title = driver.getTitle();

        if (exp.equals("Valid")) {
            if (exp_title.equals(act_title)) {
                driver.findElement(By.linkText("Logout")).click(); //Logout button
                Assert.assertTrue(true);
            } else {
                Assert.fail();
            }
        } else if (exp.equals("Invalid")) {
            if (exp_title.equals(act_title)) {
                driver.findElement(By.linkText("Logout")).click();
                Assert.fail();
            } else {
                Assert.assertTrue(true);
            }
        }
    }

    @DataProvider(name = "LoginData")
    public Object[][] getData() throws IOException {
        XLUtility xlUtility = new XLUtility(MyPaths.LOGIN_FILE);
        int totalRows = xlUtility.getRowCount("Лист1");
        int totalCols = xlUtility.getCellCount("Лист1", 1);
        String[][] loginData = new String[totalRows][totalCols];
        for (int i = 1; i <= totalRows; i++) {
            for (int j = 0; j < totalCols; j++) {
                // -1 потому что не учитываем заголовок таблицы
                loginData[i - 1][j] = xlUtility.getCellData("Лист1", i, j);
            }
        }
        return loginData;
    }

    @AfterClass
    public void tearDown() {
        driver.close();
    }
}
