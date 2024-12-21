package Twarita.uploadDownloadWithSelenium;

import java.io.FileInputStream;
import java.time.Duration;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Assertions;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Test {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		DataFormatter dataFormatter = new DataFormatter();
		String fruitName = "Apple";
		String requiredColumnName = "Price";
		String filePath = "C:\\Users\\twari\\OneDrive\\Documents\\download.xlsx";
		String updatedValue = "299";

		UpdateExcelMethods updateExcel = new UpdateExcelMethods();

		int rowNumber = updateExcel.getRowNum(filePath, fruitName);
		System.out.println(rowNumber);
		int columnNumber = updateExcel.getColumnNum(filePath, requiredColumnName);
		System.out.println(columnNumber);

		updateExcel.updateValue(filePath, rowNumber, columnNumber, updatedValue);

		WebDriver driver = new ChromeDriver();
		driver.get("https://rahulshettyacademy.com/upload-download-test/index.html");

		driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(5));

		WebDriverWait explicitWait = new WebDriverWait(driver, Duration.ofSeconds(10));

		WebElement downloadButton = driver.findElement(By.id("downloadButton"));
		downloadButton.click();

		WebElement uploadButton = driver.findElement(By.id("fileinput"));
		uploadButton.sendKeys(filePath);

		WebElement alertMessage = driver.findElement(By.className("Toastify"));
		explicitWait.until(ExpectedConditions.visibilityOf(alertMessage.findElement(By.tagName("button"))));
		explicitWait.until(ExpectedConditions.invisibilityOf(alertMessage.findElement(By.tagName("button"))));

		FileInputStream downloadedFile = new FileInputStream(filePath);
		XSSFWorkbook downloadedWorkbook = new XSSFWorkbook(downloadedFile);

		XSSFSheet sheet = downloadedWorkbook.getSheetAt(0);

		int rowCount = sheet.getPhysicalNumberOfRows();
		int columnCount = sheet.getRow(0).getLastCellNum();
		String[][] data = new String[rowCount][columnCount];
		for (int i = 1; i < rowCount; i++) {
			for (int j = 0; j < columnCount; j++) {
				Row row = sheet.getRow(i);
				Cell cellContent = row.getCell(j);
				data[i - 1][j] = dataFormatter.formatCellValue(cellContent);
			}
		}
		WebElement cellContent;

		for (int i = 0; i < rowCount - 1; i++) {
			for (int j = 0; j < columnCount; j++) {
				int colIndex = j + 1;
				cellContent = driver
						.findElement(By.xpath("//div[@id='row-" + i + "']/div[@data-column-id=" + colIndex + "]"));
				Assertions.assertEquals(data[i][j], cellContent.getText());
			}
		}

		String columnIdOfRequiredData = driver.findElement(By.xpath("//div[text()='" + requiredColumnName + "']"))
				.getDomAttribute("data-column-id");
		WebElement fruitData = driver.findElement(By.xpath("//div[text()='" + fruitName
				+ "']/parent::div/parent::div/div[@data-column-id='" + columnIdOfRequiredData + "']"));
		System.out.println(fruitData.getText());
		downloadedWorkbook.close();
		downloadedFile.close();
		driver.quit();

	}

}
