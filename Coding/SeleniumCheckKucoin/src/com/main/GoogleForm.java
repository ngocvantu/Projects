package com.main;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.InputEvent;
import java.awt.event.KeyEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
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

public class GoogleForm {
	
	FileInputStream excelFile;
	XSSFWorkbook workbook;
	XSSFSheet datatypeSheet;
	
	private static final int ROWSTART = 2;
	private static final int ROWEND = 50;

	public static void main(String[] args) { 

		GoogleForm googleForm = new GoogleForm();
		try {
			googleForm.run();
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (AWTException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	private void run() throws InterruptedException, AWTException { 
		System.setProperty("webdriver.chrome.driver", "E:\\Lib\\chromedriver.exe");
		System.setProperty("webdriver.gecko.driver", "E:\\Lib\\geckodriver.exe");
		WebDriver driver;
		
		
		ChromeOptions options = new ChromeOptions();
		// FirefoxOptions options = new FirefoxOptions();
		options.addArguments("disable-infobars");
		options.addArguments("--start-maximized");
		options.setExperimentalOption("useAutomationExtension", false);
		options.setExperimentalOption("excludeSwitches", Arrays.asList("enable-automation"));

		Map<String, Object> prefs = new HashMap<String, Object>();
		prefs.put("credentials_enable_service", false);
		options.setExperimentalOption("prefs", prefs);
		
		 

		try {
			excelFile = new FileInputStream(new File(StringStatic.FILE_NAME_GOOGLE_FORM));
			workbook = new XSSFWorkbook(excelFile);
			datatypeSheet = workbook.getSheetAt(0);
			for (int i = 1; i < 10; i++) {
				Cell cell = datatypeSheet.getRow(i).getCell(0);
				System.out.println("A" + (i + 1) + ": " + cell.getStringCellValue());
			}
			
			int j = 0;
			for (int i = ROWSTART - 1; i <= ROWEND - 1; i++) {
			
				Cell cellStatus = datatypeSheet.getRow(i).getCell(3);
				if (cellStatus.getNumericCellValue() == 0) {
					j++;
					
					Cell twitterUsernameCell = datatypeSheet.getRow(i).getCell(0);
					String twitterUsername = twitterUsernameCell.getStringCellValue().trim();
					
					Cell cell = datatypeSheet.getRow(i).getCell(2);
					String email = cell.getStringCellValue().trim();
					Cell yourNameCell = datatypeSheet.getRow(i).getCell(4);
					String yourName = yourNameCell.getStringCellValue().trim();
					Cell yourTwitterAccountCell = datatypeSheet.getRow(i).getCell(7);
					String yourTwitterAccount = yourTwitterAccountCell.getStringCellValue().trim();
					System.out.println("email: " + email);
					System.out.println("yourName: " + yourName);
					System.out.println("yourTwitterAccount: " + yourTwitterAccount);
					Thread.sleep(3000);
					driver = new ChromeDriver(options);
					WebDriverWait wait = new WebDriverWait(driver, 30);
					
					driver.get(StringStatic.FILE_NAME_GOOGLE_FORM_LINK); 
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[2]/div/div[2]/div[3]/div[2]")));
					WebElement dangNhapButton = driver.findElement(By.xpath("/html/body/div[2]/div/div[2]/div[3]/div[2]"));
					dangNhapButton.click(); 
					
					// next button to be clickable
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"identifierNext\"]")));
					WebElement emailInput = driver.findElement(By.xpath("//*[@id=\"identifierId\"]"));
					emailInput.sendKeys(StringStatic.GOOGLE_ACCOUNT);  
					driver.findElement(By.xpath("//*[@id=\"identifierNext\"]")).click();
					
					// next button to be clickable
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"passwordNext\"]")));
					WebElement pass = driver.findElement(By.xpath("//*[@id=\"password\"]/div[1]/div/div[1]/input"));
					pass.sendKeys("Thongtinaz@12");  
					Thread.sleep(500);
					driver.findElement(By.xpath("//*[@id=\"passwordNext\"]")).click();
					driver.get(StringStatic.FILE_NAME_GOOGLE_FORM_LINK); 
					 
					Thread.sleep(1000);
					fillForm(driver, email, yourName, yourTwitterAccount, twitterUsername, i, j);
					driver.close();
				}
				
			} 
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) { 
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}

	private void fillForm(WebDriver driver, String email, String yourName, String yourTwitterAccount, String twitterUsername, int i, int j) throws IOException, AWTException, InterruptedException { 
		WebDriverWait wait = new WebDriverWait(driver, 30000);
		
		((JavascriptExecutor) driver)
				.executeScript("window.open('" + StringStatic.FILE_NAME_GOOGLE_FORM_LINK + "', '_blank')");
		ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
		driver.switchTo().window(tabs.get(j));

		// driver.get(StringStatic.FILE_NAME_GOOGLE_FORM_LINK);

		
		// submit button to be clickable
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[3]/div[1]/div/div")));
		// inser5t email
		WebElement yourEmail = driver.findElement(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[2]/div[1]/div[2]/div/div[1]/div/div[1]/input"));
		yourEmail.sendKeys(email);
		// insert name
		WebElement yourName1 = driver.findElement(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[2]/div[2]/div[2]/div/div[1]/div/div[1]/input"));
		yourName1.sendKeys(yourName);
		// insert name
		WebElement yourTwitterAcc = driver.findElement(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[2]/div[3]/div[2]/div/div[1]/div/div[1]/input"));
		yourTwitterAcc.sendKeys(yourTwitterAccount);
		
		// copy username to clipboard
		String str = twitterUsername;

		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Clipboard clipboard = toolkit.getSystemClipboard();
		StringSelection strSel = new StringSelection(str);
		clipboard.setContents(strSel, null);
		
		
		
		// Press add file
		driver.findElement(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[2]/div[4]/div[3]")).click();
		
		// Press select file button
//		wait.until(ExpectedConditions.elementToBeClickable(By.linkText("Select files from your computer")));
		Thread.sleep(700);
//		driver.findElement(By.linkText("Select files from your computer")).click();
		
		// paste the image name
		Robot r = new Robot();
		
		r.mouseMove(680, 418);
		
		r.mousePress(InputEvent.BUTTON1_MASK);
		r.mouseRelease(InputEvent.BUTTON1_MASK);
		
		Thread.sleep(1000);
		
		r.keyPress(KeyEvent.VK_CONTROL);
		r.keyPress(KeyEvent.VK_V);
		r.keyRelease(KeyEvent.VK_V);
		r.keyRelease(KeyEvent.VK_CONTROL);
		
		Thread.sleep(500);
		
		r.keyPress(KeyEvent.VK_DOWN);
		r.keyRelease(KeyEvent.VK_DOWN);
		
		// press enter
		r.keyPress(KeyEvent.VK_ENTER);
		r.keyRelease(KeyEvent.VK_ENTER);
		
		// upload button
		Thread.sleep(500);
		r.mouseMove(270, 654); 
		r.mousePress(InputEvent.BUTTON1_MASK);
		r.mouseRelease(InputEvent.BUTTON1_MASK);
		
		Thread.sleep(500);
		 
		// submit button  
//		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"jd1ybd\"]/a")));
//		driver.findElement(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[3]/div[1]/div/div")).click();
		boolean ok = false;
		
		while (!ok) {

			try {
				driver.findElement(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[3]/div[1]/div/div")).click(); 
				ok = true;
			} catch (Exception e) {

//				System.out.println("Element is InVisible"); 

			}
		}
		
		
		// link button (submit xong)
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div/div[2]/div[1]/div[2]/div[3]/a"))); 
		
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			// now copy the screenshot to desired location using copyFile
			// //method
			FileUtils.copyFile(src, new File("D:/Temp/BTC/SUBMITFORMOK/" + email + ".png"));
		} catch (ElementNotVisibleException e) {
			System.out.println(e.getMessage());
			System.out.println("jkgffg");

		}
		
		 
		FileOutputStream output_file = new FileOutputStream(new File(StringStatic.FILE_NAME_GOOGLE_FORM));
		Cell cell0 = datatypeSheet.getRow(i).getCell(3);
		cell0.setCellValue(1);
		  
		workbook.write(output_file);
		output_file.close();
		   
		System.out.println("success >>>>>>>>>>>>>>>>>>> " + email);
	}

}
