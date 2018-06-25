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
import java.util.Random;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.ElementNotVisibleException;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
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

public class CreateGoogleAcc {

	FileInputStream excelFile;
	XSSFWorkbook workbook;
	XSSFSheet datatypeSheet;

	private static final int ROWSTART = 12;
	private static final int ROWEND = 14;
	private static final String PHONE = "923806387";
	private static final String LINK = "https://twitter.com/i/flow/signup";
	private static final String LINK_GOOGLE_CREATE_ACC = "https://accounts.google.com/signup/v2/webcreateaccount?service=mail&continue=https%3A%2F%2Fmail.google.com%2Fmail%2F&ltmpl=default&flowName=GlifWebSignIn&flowEntry=SignUp";
	public static final String LINK_KUCOIN_REGISTER = "https://www.kucoin.com/#/signup";
	String pass = "Thongtinaz@12";

	public static void main(String[] args) {

		CreateGoogleAcc register = new CreateGoogleAcc();
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
				Cell cellOK = datatypeSheet.getRow(i).getCell(8);
				if (cellStatus.getNumericCellValue() != 2 && !cellOK.getStringCellValue().equals("OK")) {
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
//					registerTwitter(driver, twitterUsername, i, j);
					registerGoogle(driver, twitterUsername, i, j);
					String yahooEmail = datatypeSheet.getRow(i).getCell(2).getStringCellValue();
					registerKucoin(driver, yahooEmail, i, j);
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

	private void registerKucoin(WebDriver driver, String yahooemail, int i, int j) throws InterruptedException, IOException { 
		WebDriverWait wait = new WebDriverWait(driver, 30000);

		((JavascriptExecutor) driver).executeScript("window.open('" + LINK_KUCOIN_REGISTER + "', '_blank')");
		ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());

		// close some tab

		if (tabs.size() > 10) {
			for (int k = 0; k < 6; k++) {
				ArrayList<String> tabs2 = new ArrayList<String>(driver.getWindowHandles());
				driver.switchTo().window(tabs2.get(tabs2.size() - 2));
				((JavascriptExecutor) driver).executeScript("window.close()");
			}
		}

		driver.switchTo().window(tabs.get(tabs.size() - 1));
		Thread.sleep(1000);

		 
//		WebElement checkbox = waitForElement(driver, wait, "//*[@id=\"StyleRootContentWrap\"]/div/div/div[2]/div[2]/div/p[1]/label/span[1]/input");
//		checkbox.sendKeys(Keys.SPACE);
		WebElement next = waitForElement(driver, wait, "//*[@id=\"StyleRootContentWrap\"]/div/div/div[2]/div[2]/div/p[2]/button");
		next.click(); 
		
		WebElement email = waitForElement(driver, wait, "//*[@id=\"email\"]");
		email.clear();
		email.sendKeys(yahooemail);
		
		WebElement pass = waitForElement(driver, wait, "//*[@id=\"password\"]");
		pass.clear();
		pass.sendKeys(this.pass);
		
		WebElement pass2 = waitForElement(driver, wait, "//*[@id=\"confirm\"]");
		pass2.clear();
		pass2.sendKeys(this.pass);
		
		Thread.sleep(1000);
		
		WebElement next2 = waitForElement(driver, wait, "//*[@id=\"StyleRootContentWrap\"]/div/div/div[2]/div[2]/div/form/div[6]/div/div/button");
		next2.click(); 
		
		WebElement resendButton = waitForElement(driver, wait, "//*[@id=\"StyleRootContentWrap\"]/div/div/div[2]/div[2]/div/p[2]/button");
		// Write to file
		FileOutputStream output_file = new FileOutputStream(new File(StringStatic.FILE_NAME_REGISTER));
		Cell cell0 = datatypeSheet.getRow(i).getCell(8);
		cell0.setCellValue("OK");
		workbook.write(output_file);
		output_file.close();
		
		System.out.println("success kucoin >>>>>>>>>>>>>> " + yahooemail);
	}

	private String registerGoogle(WebDriver driver, String twitterUsername, int i, int j)
			throws InterruptedException, IOException {
		Random rand = new Random();
		int  n = rand.nextInt(200) + 20;
		WebDriverWait wait = new WebDriverWait(driver, 30000);

		((JavascriptExecutor) driver).executeScript("window.open('" + LINK_GOOGLE_CREATE_ACC + "', '_blank')");
		ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());

		// close some tab

		if (tabs.size() > 10) {
			for (int k = 0; k < 6; k++) {
				ArrayList<String> tabs2 = new ArrayList<String>(driver.getWindowHandles());
				driver.switchTo().window(tabs2.get(tabs2.size() - 2));
				((JavascriptExecutor) driver).executeScript("window.close()");
			}
		}

		driver.switchTo().window(tabs.get(tabs.size() - 1));
		Thread.sleep(1000);

		// Username
		WebElement username = waitForElement(driver, wait, "//*[@id=\"usernamereg-firstName\"]");
		username.clear();
		username.sendKeys("mua");
		WebElement lastName = waitForElement(driver, wait, "//*[@id=\"usernamereg-lastName\"]");
		lastName.clear();
		lastName.sendKeys("xuan");

		WebElement email = waitForElement(driver, wait, "//*[@id=\"usernamereg-yid\"]");
		email.clear();
		email.sendKeys(twitterUsername + n);

		WebElement pass = waitForElement(driver, wait, "//*[@id=\"usernamereg-password\"]");
		pass.clear();
		pass.sendKeys("Thongtinaz@12");

		WebElement phone = waitForElement(driver, wait, "//*[@id=\"usernamereg-phone\"]");
		phone.clear();
		phone.sendKeys(PHONE);

		// date birth
		driver.findElement(By.xpath("//*[@id=\"usernamereg-month\"]")).sendKeys(Keys.ARROW_DOWN);
		driver.findElement(By.xpath("//*[@id=\"usernamereg-month\"]")).sendKeys(Keys.ARROW_DOWN);
//		driver.findElement(By.xpath("//*[@id=\"usernamereg-month\"]")).sendKeys(Keys.ENTER);

		// month
		WebElement month = waitForElement(driver, wait, "//*[@id=\"usernamereg-day\"]");
		month.clear();
		month.sendKeys("8");
		// year
		WebElement year = waitForElement(driver, wait, "//*[@id=\"usernamereg-year\"]");
		year.clear();
		year.sendKeys("1992");
		
		WebElement continueButton = waitForElement(driver, wait, "//*[@id=\"reg-submit-button\"]");
		continueButton.click();
		
		WebElement sendKey = waitForElement(driver, wait, "//*[@id=\"phone-verify-challenge\"]/form/div[2]/button");
		sendKey.click();
		
		WebElement tieptuc = waitForElement(driver, wait, "//*[@id=\"account-attributes-challenge\"]/form/div/div[3]/button");
		tieptuc.click();
		 
		Thread.sleep(2000);
		  

		// Take screen shot
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			FileUtils.copyFile(src, new File("D:/Temp/BTC/REGISTER/" + twitterUsername + ".png"));
		} catch (ElementNotVisibleException e) {
			System.out.println(e.getMessage());
			System.out.println("jkgffg");

		}

		// Write to file
		FileOutputStream output_file = new FileOutputStream(new File(StringStatic.FILE_NAME_REGISTER));
		Cell cell0 = datatypeSheet.getRow(i).getCell(2);
		cell0.setCellValue(twitterUsername + n + "@yahoo.com");
		workbook.write(output_file);
		output_file.close();
		
		System.out.println("success yahoo >>>>>>>>>>>>>> " + twitterUsername);
		return twitterUsername + n + "@yahoo.com";

	}

	private void registerTwitter(WebDriver driver, String twitterUsername, int i, int j)
			throws IOException, AWTException, InterruptedException {

		WebDriverWait wait = new WebDriverWait(driver, 30000);

		((JavascriptExecutor) driver).executeScript("window.open('" + LINK + "', '_blank')");
		ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());

		// close some tab

		if (tabs.size() > 10) {
			for (int k = 0; k < 6; k++) {
				ArrayList<String> tabs2 = new ArrayList<String>(driver.getWindowHandles());
				driver.switchTo().window(tabs2.get(tabs2.size() - 2));
				((JavascriptExecutor) driver).executeScript("window.close()");
			}
		}

		driver.switchTo().window(tabs.get(tabs.size() - 1));
		Thread.sleep(1000);

		// Username
		WebElement username = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/input");
		username.clear();
		username.sendKeys(twitterUsername);

		// Phone
		WebElement phone = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div[3]/div[2]/div/input");
		phone.clear();
		phone.sendKeys("0"+PHONE);

		// Next button
		WebElement nextButton = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div/div[3]/div");
		while (nextButton.getAttribute("aria-disabled") != null) {
		}
		nextButton.click();

		// sign up button
		WebElement signupButton = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div[5]");
		signupButton.click();

		// confirm send text ( ok button)
		WebElement okButton = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div[2]/div/div/div[2]/div[2]/div/div[3]");
		okButton.click();

		// verification field
		WebElement verificationField = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div[3]/div/div/input");
		while (!(verificationField.getAttribute("value").length() == 6)) {

		}
		// Next button after receive code
		WebElement nextButton2 = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div/div[3]/div");
		while (nextButton2.getAttribute("aria-disabled") != null) {
		}
		nextButton2.click();

		// password field
		WebElement pass = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div[3]/div/div[2]/div/input");
		pass.clear();
		pass.sendKeys("Thongtinaz@12");
		// Next button after set password
		WebElement nextButton3 = waitForElement(driver, wait,
				"//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div/div[3]/div");
		while (nextButton3.getAttribute("aria-disabled") != null) {
		}
		nextButton3.click();

		Thread.sleep(2000);

		// if
		// (driver.findElement(By.xpath("/html/body/div[2]/div/form/input[6]")).isEnabled())
		// {
		// driver.findElement(By.xpath("/html/body/div[2]/div/form/input[6]")).click();
		// }

		// skip for now button
		// WebElement skipforNow = waitForElement(driver, wait,
		// "//*[@id=\"react-root\"]/div[2]/div/div/div/div[2]/div[2]/div/div/div[2]/div[1]/div/div/div/div[3]/div");
		// skipforNow.click();

		// @username link
		WebElement usernameLink = waitForElement(driver, wait,
				"//*[@id=\"page-container\"]/div[1]/div[1]/div/div[2]/span/a");
		String myUsername = usernameLink.getText();
		System.out.println(myUsername);
 
		driver.get("https://twitter.com/kucoincom");

		// folow button
		WebElement folowbutton = waitForElement(driver, wait,
				"//*[@id=\"page-container\"]/div[1]/div/div[2]/div/div/div[2]/div/div/ul/li[6]/div/div/span[2]/button[1]");
		folowbutton.click();

		// Take screen shot
		File src = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		try {
			FileUtils.copyFile(src, new File("D:/Temp/BTC/REGISTER/" + twitterUsername + ".png"));
		} catch (ElementNotVisibleException e) {
			System.out.println(e.getMessage());
			System.out.println("jkgffg");

		}

		// Write to file
		FileOutputStream output_file = new FileOutputStream(new File(StringStatic.FILE_NAME_REGISTER));
		Cell cell0 = datatypeSheet.getRow(i).getCell(1);
		cell0.setCellValue(myUsername.replaceAll("@", ""));
		datatypeSheet.getRow(i).getCell(9).setCellValue("0"+PHONE);
		workbook.write(output_file);
		output_file.close();

		System.out.println("success >>>>>>>>>>>>>>>>>>> " + twitterUsername);
	}

	public WebElement waitForElement(WebDriver driver, WebDriverWait wait, String xpath) {
		wait.until(ExpectedConditions.elementToBeClickable(By.xpath(xpath)));
		return driver.findElement(By.xpath(xpath));
	}

}
