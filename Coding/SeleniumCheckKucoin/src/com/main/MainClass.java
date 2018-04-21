package com.main;

import java.awt.AWTException;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Random;

import javax.imageio.ImageIO;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.gargoylesoftware.htmlunit.WebWindowListener;

public class MainClass {
	private static final String FILE_NAME = "E:\\GoogleDriver\\BTC\\taikhoan_ok.xlsx";
//	private static final String FILE_NAME = "E:\\GoogleDriver\\BTC\\taikhoan_ok-form.xlsx";
//	private static final String FILE_NAME = "E:\\GoogleDriver\\BTC\\taikhoan_ok-security.xlsx";
	 
//	private static final String FILE_NAME = "C:\\Users\\tunv\\Google Drive\\BTC\\taikhoan_ok-getname.xlsx";
	
	 
	
	
	FileInputStream excelFile;
	XSSFWorkbook workbook;
	XSSFSheet datatypeSheet;

	// data in cell file
	private static final String COLUMN_DATA = "D";
	private static final String COLUMN_STATUS = "D";
	private static final int ROWSTART = 110;
	private static final int ROWEND = 210;

	// ACtion (retweet, follow kucoin)
	private static final String ACTION = "RETWEET";

	// FOLDER for image
	private static final String KUCOINSECURITY_FOLDER = "D:/Temp/BTC/KUCOINSECURITY/";
	private static final String FOLLOW_FOLDER = "D:/Temp/BTC/FOLLOW/";
	private static final String RETWEET_FOLDER = "D:/Temp/BTC/RETWEET/";
	private static final String FORM_FOLDER = "D:/Temp/BTC/FORM/";
	private static final String MAT_KHAU = "Thongtinaz@12";

	public static void main(String[] args) throws InterruptedException, IOException, AWTException {
		MainClass main = new MainClass();
		main.doAction();
	}

	public void doAction() throws InterruptedException, IOException, AWTException {
		System.setProperty("webdriver.chrome.driver", "E:\\Lib\\chromedriver.exe");
		System.setProperty("webdriver.gecko.driver", "E:\\Lib\\geckodriver.exe");
		WebDriver driver;

		excelFile = new FileInputStream(new File(FILE_NAME));
		workbook = new XSSFWorkbook(excelFile);
		datatypeSheet = workbook.getSheetAt(0);
		for (int i = 1; i < 10; i++) {
			Cell cell = datatypeSheet.getRow(i).getCell(0);
			System.out.println("A" + (i + 1) + ": " + cell.getStringCellValue());
		}


		
		for (int i = ROWSTART - 1; i <= ROWEND - 1; i++) {
			Cell cellStatus = datatypeSheet.getRow(i).getCell(3);
			if (cellStatus == null) {
				System.out.println("null cell");
				cellStatus = datatypeSheet.getRow(i).createCell(3);
				cellStatus.setCellValue(1);
			}
			if (cellStatus.getNumericCellValue() == 0) {
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
				// driver = new FirefoxDriver(options);

			// Get userstring from cel file
			Cell cell = datatypeSheet.getRow(i).getCell(0);
			String userString = cell.getStringCellValue().trim();
			
			

				switch (ACTION) {
				case "RETWEET":
					thuchien(userString, driver, i);
					break;
				case "FOLLOW":
					followKucoin(userString, driver, i);
					break;
				case "FORMKUCOIN":
					Cell celltwitterusername = datatypeSheet.getRow(i).getCell(4);
					String twitterusername = celltwitterusername.getStringCellValue().trim();
					
					Cell cellemail = datatypeSheet.getRow(i).getCell(2);
					String email = cellemail.getStringCellValue().trim();
					WebDriverWait wait = new WebDriverWait(driver, 60);
					driver.get("https://docs.google.com/forms/d/e/1FAIpQLSfeWHLEcQ5Q1VGXUzK-00oXQX4zu_uBnNKVyxAebfRr_OCX1w/viewform");
					
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[2]/div/div[2]/div[3]/div[2]/content/span")));
					WebElement googleLoginButton  = driver.findElement(By.xpath("/html/body/div[2]/div/div[2]/div[3]/div[2]/content/span"));
					googleLoginButton.click();
					wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"identifierId\"]")));
					WebElement inputEmail = driver.findElement(By.xpath("//*[@id=\"identifierId\"]"));
					 inputEmail.sendKeys("meohoang13.sinhvien@gmail.com");
					 WebElement nextEmail = driver.findElement(By.xpath("//*[@id=\"identifierNext\"]/content/span"));
					 nextEmail.click();
					 wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"password\"]/div[1]/div/div[1]/input")));
					 WebElement inputpass = driver.findElement(By.xpath("//*[@id=\"password\"]/div[1]/div/div[1]/input"));
					 inputpass.sendKeys("Thongtinaz@12");
					 WebElement nextPass = driver.findElement(By.xpath("//*[@id=\"passwordNext\"]/content/span"));
					 nextPass.click();
					 wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[2]/div[1]/div[2]/div/div[1]/div/div[1]/input")));
					fillFormKucoin(userString, driver, i, twitterusername, email);
					break;
				case "GETNAME":
					getTwittername(userString, driver, i);
					break;
				case "KUCOINSECURITY":
					Cell celltwitterusername1 = datatypeSheet.getRow(i).getCell(4);
					String twitterusername1 = celltwitterusername1.getStringCellValue().trim();
					
					Cell cellemail1 = datatypeSheet.getRow(i).getCell(2);
					String email1 = cellemail1.getStringCellValue().trim();
					kucoinSecurity(userString, driver, i, twitterusername1, email1);
					break;
				default:
					break;
				}
			}
			// Runtime.getRuntime().exec("taskkill /F /IM chromedriver.exe");
		}

		workbook.close();
	}

	private void kucoinSecurity(String userString, WebDriver driver, int i, String twitterusername1, String email1) throws InterruptedException, IOException {
		driver.get("https://www.kucoin.com/#/login");
		System.out.println(email1);
		WebElement checkBoc1 = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div/div/div[2]/div/div[1]/div[1]/label/span/input"));
		checkBoc1.click();
		
		WebElement checkBoc2 = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div/div/div[2]/div/div[1]/div[2]/label/span/input"));
		checkBoc2.click();
		
		WebElement checkBoc3 = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div/div/div[2]/div/div[1]/div[3]/label/span/input"));
		checkBoc3.click();
		
		WebElement checkBoc4 = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div/div/div[2]/div/div[1]/div[4]/label/span/input"));
		checkBoc4.click();
		
		WebElement checkBoc5 = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div/div/div[2]/div/div[1]/div[5]/label/span/input"));
		checkBoc5.click(); 
		
		Thread.sleep(1000);
		
		WebElement btnconfirm = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div/div/div[2]/div/div[3]/button"));
		btnconfirm.click();
		 
		
		WebDriverWait wait = new WebDriverWait(driver, 30);
		wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//*[@id=\"username\"]")));
		
		WebElement email = driver
				.findElement(By.xpath("//*[@id=\"username\"]"));
		email.clear();
		email.sendKeys(email1);
		
		WebElement pass = driver
				.findElement(By.xpath("//*[@id=\"password\"]"));
		pass.clear();
		pass.sendKeys("Thongtinaz@12"); 
		
		WebElement btnLogin = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div/div/div[2]/form/div[4]/div/div/button"));
		btnLogin.click();
		  
		
		wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("/html/body/div[4]/div/div[2]/div/div[1]/div[3]/div/button")));
		  
		Thread.sleep(2000);
		
		WebElement btnTips = driver
				.findElement(By.xpath("/html/body/div[4]/div/div[2]/div/div[1]/div[3]/div/button"));
		btnTips.click();
		
		
		wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//*[@id=\"safeWords\"]")));
		
		JavascriptExecutor jsExecutor = ((JavascriptExecutor) driver); 
		jsExecutor.executeScript("window.scrollBy(0,340)", "");
		
		
		WebElement safeWords = driver
				.findElement(By.xpath("//*[@id=\"safeWords\"]"));
		safeWords.sendKeys("tunguyen"); 
		 
		WebElement secureQuestion1 = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div[2]/div/div/div/div[2]/div/div[2]/div/div[2]/div/form/div[3]/div[1]/div[2]/div/div/div/span"));
		secureQuestion1.click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("/html/body/div[4]/div/div/div/ul/li[5]")).click(); 
		WebElement answereQuestion1 = driver.findElement(By.xpath("//*[@id=\"a1\"]"));
		answereQuestion1.sendKeys("ngo gia tu");
		Thread.sleep(1000);
		
		WebElement secureQuestion2 = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div[2]/div/div/div/div[2]/div/div[2]/div/div[2]/div/form/div[4]/div[1]/div[2]/div/div/div/span"));
		secureQuestion2.click();
		Thread.sleep(1000);
		driver.findElement(By.xpath("/html/body/div[5]/div/div/div/ul/li[6]")).click(); 
		WebElement answereQuestion2 = driver.findElement(By.xpath("//*[@id=\"a2\"]"));
		answereQuestion2.sendKeys("cho");
		Thread.sleep(1000);
		
		WebElement secureQuestion3 = driver
				.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div[2]/div/div/div/div[2]/div/div[2]/div/div[2]/div/form/div[5]/div[1]/div[2]/div/div/div/span"));
		secureQuestion3.click();
		Thread.sleep(1000); 
		driver.findElement(By.xpath("/html/body/div[6]/div/div/div/ul/li[5]")).click(); 
		WebElement answereQuestion3 = driver.findElement(By.xpath("//*[@id=\"a3\"]"));
		answereQuestion3.sendKeys("nguyenkha");
		Thread.sleep(2000);
		
		jsExecutor.executeScript("window.scrollBy(0,-50)", "");
		
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			// now copy the screenshot to desired location using copyFile
			// //method
			FileUtils.copyFile(src, new File(KUCOINSECURITY_FOLDER + email1 + ".png"));
		} catch (ElementNotVisibleException e) {
			System.out.println(e.getMessage());
			System.out.println("jkgffg");

		}
		
		Thread.sleep(2000);
		
		WebElement btnSubbmit = driver.findElement(By.xpath("//*[@id=\"StyleRootContentWrap\"]/div/div[2]/div/div/div/div[2]/div/div[2]/div/div[2]/div/form/div[6]/div/div/button"));
		btnSubbmit.click();
		
		
		FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
		Cell cell0 = datatypeSheet.getRow(i).getCell(3);
		cell0.setCellValue(1);
		  
		workbook.write(output_file);
		output_file.close();
		   
		System.out.println("success " + email1);
		Thread.sleep(5000);
//		driver.close();
	}

	private void getTwittername(String usernamestring, WebDriver driver, int i) throws InterruptedException, IOException { 
		driver.get("https://twitter.com/login/");

		WebDriverWait wait = new WebDriverWait(driver, 5);
		wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/div[2]/button")));
		wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[1]/input")));

		WebElement username = driver
				.findElement(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[1]/input"));
		WebElement pass = driver
				.findElement(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[2]/input"));

		username.clear();
		username.sendKeys(usernamestring);
		Thread.sleep(1000);
		pass.clear();
		pass.sendKeys("thongtin27592");
		Thread.sleep(1000);
		WebElement btnLogin = driver
				.findElement(By.xpath("/html/body/div[1]/div[2]/div/div/div[1]/form/div[2]/button"));
		btnLogin.click();
		Thread.sleep(3000);
		
		WebElement userlink = driver
				.findElement(By.xpath("//*[@id=\"page-container\"]/div[1]/div[1]/div/div[2]/div/a"));
		FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
		Cell cell0 = datatypeSheet.getRow(i).getCell(3);
		cell0.setCellValue(1);
		
		Cell cell = datatypeSheet.getRow(i).getCell(4);
		String name = userlink.getText();
		cell.setCellValue(name);
		System.out.println(name);
		workbook.write(output_file);
		output_file.close();
		driver.close();
	}

	private void fillFormKucoin(String userString, WebDriver driver, int i, String twittername, String email) throws AWTException, IOException { 
		try {
			System.out.println("username: " + userString);
			System.out.println("twittername: " + twittername);
			System.out.println("email: " + email);
			System.out.println("image: " + "D:/Temp/BTC/Retweet1/" +  userString + ".png");
			WebDriverWait wait = new WebDriverWait(driver, 20);
		 wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[2]/div[1]/div[2]/div/div[1]/div/div[1]/input")));
		 
			ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());
			driver.switchTo().window(tabs.get(0));
			//To navigate to new link/URL in 2nd new tab
			driver.get("http://facebook.com");
			Thread.sleep(10000);
		
		wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//*[@id=\"form_submit\"]")));
		WebElement yourname = driver
				.findElement(By.xpath("//*[@id=\"202113127\"]/div[2]/div/div[1]/input"));
		yourname.clear();
		yourname.sendKeys(twittername); 
			Thread.sleep(1000);
			
			WebElement kucoinemail= driver
					.findElement(By.xpath("//*[@id=\"202113126\"]/div[2]/div/div/div/input"));
			kucoinemail.clear();
			kucoinemail.sendKeys(email);
			Thread.sleep(1000);
			
			WebElement twitteraccount= driver
					.findElement(By.xpath("//*[@id=\"202113128\"]/div[2]/div/div/div/input"));
			twitteraccount.clear();
			twitteraccount.sendKeys("https://twitter.com/" + userString);
			Thread.sleep(1000);
			
			WebElement snapshot= driver
					.findElement(By.xpath("//*[@id=\"202113129\"]/div[2]/ul/li[1]/div/input"));
			snapshot.clear(); 
			
			snapshot.sendKeys("D:/Temp/BTC/Retweet1/" +  userString + ".png");
			
			Thread.sleep(3000);
			
			WebElement btnsubmit= driver
					.findElement(By.xpath("//*[@id=\"form_submit\"]"));
			btnsubmit.click();
			
			
			
			
			JavascriptExecutor jsExecutor = ((JavascriptExecutor) driver);
			Random rand = new Random(); 
			int n = rand.nextInt((190 - 185) + 1) + 185;
			jsExecutor.executeScript("scroll(0, " + n + ");");
			File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

			Robot robot = new Robot();
			String fileName = FORM_FOLDER + userString + ".png";

			// Define an area of size 500*400 starting at coordinates (10,50)
			int rong = rand.nextInt((1200 - 1000) + 1) + 1000;
			int cachTren =  rand.nextInt((30 - 10) + 1) + 10;
			Rectangle rectArea = new Rectangle(30, cachTren, rong, 700);
			BufferedImage screenFullImage = robot.createScreenCapture(rectArea);
			ImageIO.write(screenFullImage, "png", new File(fileName));

			// try {
			// // now copy the screenshot to desired location using copyFile
			// // //method
			// FileUtils.copyFile(src, new File(RETWEET_FOLDER + usernamestring
			// + ".png"));
			// } catch (ElementNotVisibleException e) {
			// System.out.println(e.getMessage());
			// System.out.println("jkgffg");
			//
			// }

			FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
			
			Cell cell = datatypeSheet.getRow(i).getCell(3);
			cell.setCellValue(1);
			workbook.write(output_file);
			output_file.close();

			System.out.println("success: " + userString);
			Thread.sleep(1000);
//			driver.close();
//			driver.quit();
		} catch (InterruptedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	private void followKucoin(String usernamestring, WebDriver driver, int i) throws InterruptedException, IOException {
		driver.get("https://twitter.com/login/");

		WebDriverWait wait = new WebDriverWait(driver, 5);
		wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/div[2]/button")));
		wait.until(ExpectedConditions
				.elementToBeClickable(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[1]/input")));

		WebElement username = driver
				.findElement(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[1]/input"));
		WebElement pass = driver
				.findElement(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[2]/input"));

		username.clear();
		username.sendKeys(usernamestring);
		Thread.sleep(1000);
		pass.clear();
		pass.sendKeys("Thongtinaz@12");

		WebElement btnLogin = driver
				.findElement(By.xpath("/html/body/div[1]/div[2]/div/div/div[1]/form/div[2]/button"));
		btnLogin.click();
		Thread.sleep(3000);
		driver.navigate().to("https://twitter.com/kucoincom");
		Thread.sleep(3000);
		WebElement folloewButton = driver.findElement(By.xpath(
				"//*[@id=\"page-container\"]/div[1]/div/div[2]/div/div/div[2]/div/div/ul/li[6]/div/div/span[2]/button[1]"));
		folloewButton.click();
		Thread.sleep(2000);
		driver.navigate().to("https://twitter.com/" + usernamestring);
		Thread.sleep(3000);
		JavascriptExecutor jsExecutor = ((JavascriptExecutor) driver);
		Random rand = new Random();

		int n = rand.nextInt((190 - 185) + 1) + 185;
		jsExecutor.executeScript("scroll(0, " + n + ");");
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			// now copy the screenshot to desired location using copyFile
			// //method
			FileUtils.copyFile(src, new File(FOLLOW_FOLDER + usernamestring + ".png"));
		} catch (ElementNotVisibleException e) {
			System.out.println(e.getMessage());
			System.out.println("jkgffg");

		}

		FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
		Cell cell = datatypeSheet.getRow(i).getCell(3);
		cell.setCellValue(1);
		workbook.write(output_file);
		output_file.close();
		System.out.println("success: " + usernamestring);
		Thread.sleep(3000);
		driver.close();

	}

	private void thuchien(String usernamestring, WebDriver driver, int i) throws InterruptedException {
		try {
			driver.get("https://twitter.com/login/");

			WebDriverWait wait = new WebDriverWait(driver, 5);
			wait.until(ExpectedConditions
					.elementToBeClickable(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/div[2]/button")));
			wait.until(ExpectedConditions.elementToBeClickable(
					By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[1]/input")));

			WebElement username = driver
					.findElement(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[1]/input"));
			WebElement pass = driver
					.findElement(By.xpath("//*[@id=\"page-container\"]/div/div[1]/form/fieldset/div[2]/input"));

			username.clear();
			username.sendKeys(usernamestring);
			Thread.sleep(1000);
			pass.clear();
			if ("btckhongphaila1".equals(usernamestring) || 
					"tatcadaunha1".equals(usernamestring) || 
					"nenphaicogang1".equals(usernamestring) ||
					"hochanhvask1".equals(usernamestring)) {
				pass.sendKeys("thongtin27592");
			} else {
				pass.sendKeys(MAT_KHAU);
			}
			
			Thread.sleep(1000);
			WebElement btnLogin = driver
					.findElement(By.xpath("/html/body/div[1]/div[2]/div/div/div[1]/form/div[2]/button"));
			btnLogin.click();
			Thread.sleep(3000);
			driver.navigate().to("https://twitter.com/kucoincom");
			Thread.sleep(3000);
			
			if (driver.findElement(By.xpath( "//*[@id=\"stream-item-tweet-986248291573641216\"]/div[1]/div[2]/div[4]/div[2]/div[2]/button[1]/div/span[1]")).isDisplayed()) {
				WebElement reTweetIcon = driver.findElement(By.xpath(
						"//*[@id=\"stream-item-tweet-986248291573641216\"]/div[1]/div[2]/div[4]/div[2]/div[2]/button[1]/div/span[1]"));
				reTweetIcon.click();
				Thread.sleep(5000);
				WebElement reTweetButton = driver.findElement(
						By.xpath("/html/body/div[24]/div/div[2]/form/div[2]/div[3]/button"));
				reTweetButton.click();
			} else {
				WebElement unTweetIcon = driver.findElement(By.xpath(
						"//*[@id=\"stream-item-tweet-986248291573641216\"]/div[1]/div[2]/div[4]/div[2]/div[2]/button[2]/div/span[1]"));
				unTweetIcon.click();
				Thread.sleep(2000);
				WebElement reTweetIcon = driver.findElement(By.xpath(
						"//*[@id=\"stream-item-tweet-986248291573641216\"]/div[1]/div[2]/div[4]/div[2]/div[2]/button[1]/div/span[1]"));
				reTweetIcon.click();
				Thread.sleep(5000);
				WebElement reTweetButton = driver.findElement(
						By.xpath("/html/body/div[24]/div/div[2]/form/div[2]/div[3]/button")); 
				reTweetButton.click();
			}
			
			Thread.sleep(10000);
			// WebElement userLink = driver
			// .findElement(By.xpath("//*[@id=\"page-container\"]/div[1]/div[1]/div/div[2]/span/a/span"));
			// userLink.click();

			driver.navigate().to("https://twitter.com/" + usernamestring);
			Thread.sleep(3000);
			JavascriptExecutor jsExecutor = ((JavascriptExecutor) driver);
			Random rand = new Random();

			int n = rand.nextInt((250 - 200) + 1) + 200;
			jsExecutor.executeScript("scroll(0, " + n + ");");
			File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

			Robot robot = new Robot();
			String fileName = RETWEET_FOLDER + usernamestring + ".png";

			// Define an area of size 500*400 starting at coordinates (10,50)
			int rong = rand.nextInt((1200 - 1000) + 1) + 1000;
			int cachTren =  rand.nextInt((30 - 10) + 1) + 10;
			int cachPhai =  rand.nextInt((50 - 30) + 1) + 30;
			Rectangle rectArea = new Rectangle(cachPhai, cachTren, rong, 700);
			BufferedImage screenFullImage = robot.createScreenCapture(rectArea);
			ImageIO.write(screenFullImage, "png", new File(fileName));

			// try {
			// // now copy the screenshot to desired location using copyFile
			// // //method
			// FileUtils.copyFile(src, new File(RETWEET_FOLDER + usernamestring
			// + ".png"));
			// } catch (ElementNotVisibleException e) {
			// System.out.println(e.getMessage());
			// System.out.println("jkgffg");
			//
			// }

			FileOutputStream output_file = new FileOutputStream(new File(FILE_NAME));
			Cell cell = datatypeSheet.getRow(i).getCell(3);
			cell.setCellValue(1);
			workbook.write(output_file);
			output_file.close();

			System.out.println("success: " + usernamestring);
			Thread.sleep(3000);
			driver.close();
			driver.quit();
		} catch (Exception e) {
			e.printStackTrace();
			return;
		}
	}
}
