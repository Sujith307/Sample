package com.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Sample_One {

	public static void excel_Read() {

		try {
			// file locate
			File f = new File("C:\\Users\\HP\\eclipse-workspace\\Oct_2023_Project"
					+ "\\src\\test\\resources\\TestData\\Test_One.xlsx");
			// fileread
			FileInputStream fis = new FileInputStream(f);
			// workbook access
			Workbook wb = new XSSFWorkbook(fis);
			Sheet sheet = wb.getSheet("Sheet1");
			Row row = sheet.getRow(0);
			Cell cell = row.getCell(1);
			System.out.println(cell);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void excel_Read_All() {

		try {
			// file locate
			File f = new File("C:\\Users\\HP\\eclipse-workspace\\Oct_2023_Project"
					+ "\\src\\test\\resources\\TestData\\Test_One.xlsx");
			// fileread
			FileInputStream fis = new FileInputStream(f);
			// workbook access
			Workbook wb = new XSSFWorkbook(fis);
			Sheet sheet = wb.getSheet("Sheet1");
			int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
			for (int i = 0; i < physicalNumberOfRows; i++) {
				Row row = sheet.getRow(i);
				int physicalNumberOfCells = row.getPhysicalNumberOfCells();
				for (int j = 0; j < physicalNumberOfCells; j++) {
					Cell cell = row.getCell(j);
					System.out.println(cell);
				}
			}
		} catch (FileNotFoundException e) {

		} catch (IOException e) {

		}

	}

	public static void excel_Read_Reuse() {

		try {
			// file locate
			File f = new File("C:\\Users\\HP\\eclipse-workspace\\Oct_2023_Project"
					+ "\\src\\test\\resources\\TestData\\Test_One.xlsx");
			// fileread
			FileInputStream fis = new FileInputStream(f);
			// workbook access
			Workbook wb = new XSSFWorkbook(fis);
			Sheet sheet = wb.getSheet("Sheet1");
			int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
			for (int i = 0; i < physicalNumberOfRows; i++) {
				Row row = sheet.getRow(i);
				int physicalNumberOfCells = row.getPhysicalNumberOfCells();
				for (int j = 0; j < physicalNumberOfCells; j++) {
					Cell cell = row.getCell(j);
					int cellType = cell.getCellType();
					if (cellType == 1) {
						String stringCellValue = cell.getStringCellValue();
						System.out.println(stringCellValue);
					} else if (cellType == 0) {
						if (DateUtil.isCellDateFormatted(cell)) {
							Date dateCellValue = cell.getDateCellValue();
							SimpleDateFormat sm = new SimpleDateFormat("dd/MM/yy");
							String format = sm.format(dateCellValue);
							System.out.println(format);
						} else {
							double numericCellValue = cell.getNumericCellValue();
							long l = (long) numericCellValue;
							String valueOf = String.valueOf(l);
							System.out.println(valueOf);
						}
					}
				}
			}
		} catch (FileNotFoundException e) {

		} catch (IOException e) {

		}

	}

	public static String excel_Read_Base_ReUse(int row1, int cell1) {
		String value= null;
		try {
			// file locate
			File f = new File("C:\\Users\\HP\\eclipse-workspace\\Oct_2023_Project"
					+ "\\src\\test\\resources\\TestData\\Test_One.xlsx");
			// fileread
			FileInputStream fis = new FileInputStream(f);
			// workbook access
			Workbook wb = new XSSFWorkbook(fis);
			Sheet sheet = wb.getSheet("Sheet1");
			Row row2 = sheet.getRow(row1);
			Cell cell2 = row2.getCell(cell1);
			int cellType = cell2.getCellType();
			if (cellType == 1) {
				value = cell2.getStringCellValue();
				System.out.println(value);
			} else if (cellType == 0) {
				if (DateUtil.isCellDateFormatted(cell2)) {
					Date dateCellValue = cell2.getDateCellValue();
					SimpleDateFormat sm = new SimpleDateFormat("dd/MM/yy");
					value = sm.format(dateCellValue);
					System.out.println(value);
				} else {
					double numericCellValue = cell2.getNumericCellValue();
					long l = (long) numericCellValue;
					value = String.valueOf(l);
					System.out.println(value);
				}
			}

		} catch (FileNotFoundException e) {

		} catch (IOException e) {
		}
		return value;

	}

	public static void main(String[] args) {
		WebDriver driver = new ChromeDriver();
		driver.get("https://www.facebook.com");
		driver.manage().window().maximize();
		driver.findElement(By.id("email")).sendKeys(excel_Read_Base_ReUse(2,2));
	}

}
