package com.curr.mindtree;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.NumberFormat;
import java.text.ParseException;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import io.github.bonigarcia.wdm.WebDriverManager;



public class conv_currency {

	public static void main(String[] args) throws IOException, InterruptedException, ParseException {
		Scanner sc=new Scanner(System.in); 
		System.out.println("Please enter Date for conversion int the format MMDDYYYY(Eg:07/01/2020 for Jul 1,2020)");
		String dateInput = sc.nextLine();
		System.out.println("Please enter the full path of input file(Right click on file -> properties -> security tab -> object name)");
		String filePath = sc.nextLine();
		sc.close();

		String sheetName = "Sheet1";
		
		//Using google chrome driver
		//System.setProperty("webdriver.chrome.driver","chromedriver.exe");
		WebDriverManager.chromedriver().setup();
		WebDriver driver = new ChromeDriver();
		driver.manage().window().maximize();
		
		//Website used for conversion
		driver.get("https://www1.oanda.com/currency/converter/");
		
		WebElement date = driver.findElement(By.id("end_date_input"));
		date.sendKeys(Keys.CONTROL,Keys.chord("a"));
		date.sendKeys(Keys.BACK_SPACE);
		date.sendKeys(dateInput);
		
		File file = new File(filePath);
		FileInputStream input = new FileInputStream(file);
		
		//Using XSSF for .xlsx files
		XSSFWorkbook book = new XSSFWorkbook(input);
		Sheet sheet = book.getSheet(sheetName);
		
		//Calculates total rows in excel
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		
		for(int i=1;i<rowCount+1;i++)
		{
			Row row = sheet.getRow(i);

			driver.findElement(By.id("base_currency_input")).sendKeys(row.getCell(0).getStringCellValue());
			driver.findElement(By.id("base_currency_input")).sendKeys(Keys.ENTER);
			
			driver.findElement(By.id("quote_currency_input")).sendKeys(row.getCell(1).getStringCellValue());
			driver.findElement(By.id("quote_currency_input")).sendKeys(Keys.ENTER);
			
			WebElement amount = driver.findElement(By.id("base_amount_input"));
			
			Thread.sleep(2000);
			String Fathernum = amount.getAttribute("value"); 
			
			//Removing commas in the output amount
			Number num = NumberFormat.getNumberInstance(java.util.Locale.US).parse(Fathernum);
			
			sheet.getRow(i).createCell(5).setCellValue(num.doubleValue());
			
			driver.findElement(By.id("quote_currency_input")).clear();
			driver.findElement(By.id("base_currency_input")).clear();
			driver.findElement(By.id("base_amount_input")).clear();
			
		}
		driver.close();
		input.close();
		
		FileOutputStream output = new FileOutputStream(file);
		book.write(output);
		output.close();
		book.close();

	}

}
