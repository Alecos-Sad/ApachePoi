import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.AfterClass;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

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
    public Object[][] getData() {
        return new String[][]{
                {"admin@yourstore.com", "admin", "Valid"},
                {"admin@yourstore.com", "adm", "Invalid"},
                {"adm@yourstore.com", "admin", "Invalid"},
                {"adm@yourstore.com", "adm", "Invalid"}
        };
    }

    @AfterClass
    public void tearDown() {
        driver.close();
    }
}
