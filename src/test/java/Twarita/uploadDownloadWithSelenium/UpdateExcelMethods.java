package Twarita.uploadDownloadWithSelenium;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateExcelMethods {

	public void updateValue(String filePath, int rowNumber, int columnNumber, String updatedValue) throws Exception {
		FileInputStream downloadedFile = new FileInputStream(filePath);
		XSSFWorkbook downloadedWorkbook = new XSSFWorkbook(downloadedFile);

		XSSFSheet sheet = downloadedWorkbook.getSheetAt(0);

		Row desiredRow = sheet.getRow(rowNumber);
		Cell desiredCell = desiredRow.getCell(columnNumber);
		desiredCell.setCellValue(updatedValue);

		FileOutputStream fileOutputStream = new FileOutputStream(filePath);
		downloadedWorkbook.write(fileOutputStream);

		downloadedWorkbook.close();
		downloadedFile.close();
	}

	public int getColumnNum(String filePath, String requiredColumnName) throws IOException {
		FileInputStream downloadedFile = new FileInputStream(filePath);
		XSSFWorkbook downloadedWorkbook = new XSSFWorkbook(downloadedFile);

		XSSFSheet sheet = downloadedWorkbook.getSheetAt(0);

		Row firstRow = sheet.getRow(0);
		int columnCount = 0;
		Iterator<Cell> cellIterator = firstRow.iterator();
		while (cellIterator.hasNext()) {
			if (cellIterator.next().getStringCellValue().equalsIgnoreCase(requiredColumnName))
				break;
			columnCount++;

		}
		downloadedWorkbook.close();
		downloadedFile.close();
		return columnCount;
	}

	public int getRowNum(String filePath, String fruitName) throws Exception {
		// TODO Auto-generated method stub
		FileInputStream downloadedFile = new FileInputStream(filePath);
		XSSFWorkbook downloadedWorkbook = new XSSFWorkbook(downloadedFile);

		XSSFSheet sheet = downloadedWorkbook.getSheetAt(0);

		int rowNumber = 0;
		int rowCount = sheet.getPhysicalNumberOfRows();
		Row firstRow = sheet.getRow(0);
		int columnCount = 0;
		Iterator<Cell> cellIterator = firstRow.iterator();
		while (cellIterator.hasNext()) {
			if (cellIterator.next().getStringCellValue().equalsIgnoreCase("fruit_name"))
				break;
			columnCount++;

		}
		// columnCount=count1;
		System.out.println(columnCount);
		int count = 0;
		while (count < rowCount) {
			if (sheet.getRow(count).getCell(columnCount).getStringCellValue().equalsIgnoreCase(fruitName)) {
				rowNumber = count;
				break;
			}
			count++;
		}
		downloadedWorkbook.close();
		downloadedFile.close();
		return rowNumber;
	}

}
