package com.main;

import java.awt.AWTException; 
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.PointerInput.MouseButton;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Register {
	
	FileInputStream excelFile;
	XSSFWorkbook workbook;
	XSSFSheet datatypeSheet; 
	
	private static final int ROWSTART = 2;
	private static final int ROWEND = 10;
	private static final String PHONE = "01668641345";
	private static final String LINK = "https://hub.docker.com/";

	public static void main(String[] args) {  
		
		Register register = new Register();
		try {
			register.run();
		} catch (InterruptedException e) {
			e.printStackTrace();
		} catch (AWTException e) {
			e.printStackTrace();
		}
	}

	private void run() throws InterruptedException, AWTException { 
		System.setProperty("webdriver.chrome.driver", "E:\\Lib\\chromedriver.exe");
		System.setProperty("webdriver.gecko.driver", "E:\\Lib\\geckodriver.exe");
		WebDriver driver;
		
		
		
		  
		try {
			excelFile = new FileInputStream(new File(StringStatic.FILE_NAME_REGISTER));
			workbook = new XSSFWorkbook(excelFile);
			datatypeSheet = workbook.getSheetAt(0);
			for (int i = 1; i < 10; i++) {
				Cell cell = datatypeSheet.getRow(i).getCell(0);
				System.out.println("A" + (i + 1) + ": " + cell.getStringCellValue());
			}
			
			int j = 0;
			for (int i = ROWSTART - 1; i <= ROWEND - 1; i++) {
			
				Cell cellStatus = datatypeSheet.getRow(i).getCell(3);
				if (cellStatus.getNumericCellValue() != 2) {
					j++;
					
					Cell twitterUsernameCell = datatypeSheet.getRow(i).getCell(0);
					String twitterUsername = twitterUsernameCell.getStringCellValue().trim();
					 
					Thread.sleep(1000);
					
					ChromeOptions options = new ChromeOptions();
					// FirefoxOptions options = new FirefoxOptions();
					options.addArguments("disable-infobars");
					options.addArguments("--start-maximized");
					options.setExperimentalOption("useAutomationExtension", false);
					options.setExperimentalOption("excludeSwitches", Arrays.asList("enable-automation"));

					Map<String, Object> prefs = new HashMap<String, Object>();
					prefs.put("credentials_enable_service", false);
					options.setExperimentalOption("prefs", prefs);
					driver = new ChromeDriver(options); 
					Thread.sleep(500);
					registerTwitter(driver, twitterUsername, i, j);
				}
				
			} 
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}

	private void registerTwitter(WebDriver driver, String twitterUsername, int i, int j) throws IOException, AWTException, InterruptedException { 
		 
		
		WebDriverWait wait = new WebDriverWait(driver, 30000);
		
		((JavascriptExecutor) driver)
				.executeScript("window.open('" + LINK + "', '_blank')");
		ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
		
		// close some tab
		
		if (tabs.size() > 10) {
			for (int k = 0; k < 6; k++) {
				ArrayList<String> tabs2 = new ArrayList<String>(driver.getWindowHandles());
				driver.switchTo().window(tabs2.get(tabs2.size()-2));
				((JavascriptExecutor) driver)
				.executeScript("window.close()");
			}
		}
		 
		
		driver.switchTo().window(tabs.get(tabs.size()-1));
		Thread.sleep(1000);
		
		//Username
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"login-signin\"]")));
		driver.findElement(By.xpath("//*[@id=\"login-username\"]")).clear();
		driver.findElement(By.xpath("//*[@id=\"login-username\"]")).sendKeys(twitterUsername);
		driver.findElement(By.xpath("//*[@id=\"login-signin\"]")).click();
		
		// Phone
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"login-passwd\"]")));
		driver.findElement(By.xpath("//*[@id=\"login-passwd\"]")).clear();
		driver.findElement(By.xpath("//*[@id=\"login-passwd\"]")).sendKeys(PHONE);
		driver.findElement(By.xpath("//*[@id=\"login-signin\"]")).click();
		 
		
		// Take screen shot
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try { 
			FileUtils.copyFile(src, new File("D:/Temp/BTC/TEST_TWITTER_USER_LINK/" + twitterUsername + ".png"));
		} catch (ElementNotVisibleException e) {
			System.out.println(e.getMessage());
			System.out.println("jkgffg");

		}
		
		// Write to file 
		FileOutputStream output_file = new FileOutputStream(new File(StringStatic.FILE_NAME_REGISTER));
		workbook.write(output_file);
		output_file.close();
		   
		System.out.println("success >>>>>>>>>>>>>>>>>>> " + twitterUsername);
	}

}
