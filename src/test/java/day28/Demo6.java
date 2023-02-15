package day28;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.Duration;
import java.util.Date;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Demo6 {

	public static void main(String[] args) throws EncryptedDocumentException, IOException, InterruptedException {

		Workbook wb = WorkbookFactory.create(new FileInputStream("./data/Google.xlsx"));
//		String data = wb.getSheet("java").getRow(1).getCell(1).getStringCellValue();
//		System.out.println(data);
		WebDriver driver = new ChromeDriver();
		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(20));
		WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(20));

		driver.get("https://www.google.com");

		driver.findElement(By.name("q")).sendKeys("java");
		List<WebElement> list = driver.findElements(By.xpath("//span[contains(.,'java')]"));

		int count = list.size();
		System.out.println(count);

		for (int i = 0; i < count; i++) {
			
			String text = list.get(i).getText();
			wait.until(ExpectedConditions.titleContains("Google"));
			System.out.println((i + 1) + ". " + text);

			wb.getSheet("java").getRow(i).createCell(1).setCellValue(text);
			wb.write(new FileOutputStream("./data/Google.xlsx"));

		}
		driver.quit();
		wb.close();
	}

}
