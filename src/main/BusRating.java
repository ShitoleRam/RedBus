package main;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class BusRating {

	public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
		
		org.apache.poi.ss.usermodel.Sheet sheet;
		
		System.setProperty("webdriver.gecko.driver",  "C:\\selenium\\geckodriver.exe");
		WebDriver driver = new FirefoxDriver();
		
		driver.get("https://www.redbus.in/");
		
		driver.manage().window().maximize();
		Thread.sleep(2000);
		
		driver.findElement(By.xpath("//input[@id='src']")).sendKeys("Mumbai");
		Thread.sleep(2000);
		driver.findElement(By.xpath("//li[@class='selected']")).click();
		
		driver.findElement(By.xpath("//input[@id='dest']")).sendKeys("Pune");
		Thread.sleep(2000);
		driver.findElement(By.xpath("//li[@class='selected']")).click();
		
		driver.findElement(By.xpath("//input[@id='onward_cal']")).click();
		
		Actions act = new Actions(driver);
		
		act.moveToElement(driver.findElement(By.xpath("//td[@class='current day']"))).click().perform();
		Thread.sleep(2000);
		driver.findElement(By.xpath("//button[@id='search_btn']")).click();
		
		Thread.sleep(10000);
	try {
			WebDriverWait wait= new WebDriverWait(driver, Duration.ofSeconds(30));
			wait.until(ExpectedConditions.visibilityOf(driver.findElement(By.xpath("//i[@class='icon icon-close']")))).click();
		
		}
		catch(Exception e) {
			System.out.println(e);
			
		}
		finally{
		
		
		JavascriptExecutor js = (JavascriptExecutor)driver;
		js.executeScript("window.scrollBy(0,500)","");
		
		Thread.sleep(2000);
	List<WebElement> element  = new ArrayList<>(driver.findElements(By.xpath("//div[@class='rating-sec lh-24']")));
	
	System.out.println(element.size());
	for(int i = 0; i<element.size(); i++)
	{
		String a = element.get(i).getText().trim();

		float b = Float.parseFloat(a);
		
		FileInputStream file = new FileInputStream("Excel Sheet\\ram2.xlsx");
		Workbook src= WorkbookFactory.create(file);
		
		FileOutputStream file1 = new FileOutputStream("Excel Sheet\\ram2.xlsx");
		
		src.write(file1);
		
		if(b>4.0)
		{
			System.out.println(b + " is > 4");
			
			 sheet = src.getSheet("Sheet1");
			 sheet.createRow(i).createCell(0).setCellValue(b);
		}
		else
		{
			System.out.println(b + " < 4");
			
			 sheet = src.getSheet("Sheet2");
			 sheet.createRow(i).createCell(0).setCellValue(b);
		}
	}
		

	}
	}

}
