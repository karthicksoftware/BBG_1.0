package Poc.BBG;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.logging.Level;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.logging.LogEntries;
import org.openqa.selenium.logging.LogEntry;
import org.openqa.selenium.logging.LogType;
import org.openqa.selenium.logging.LoggingPreferences;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.Assert;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

public class JSErrorExtractor {
	WebDriver driver = null;

	@BeforeTest
	@Parameters(value = { "browser" })
	public void setUp(String value) throws Exception {
		DesiredCapabilities caps = null;
		if (value.equals("chrome")) {
			System.out.println("Chrome constructed");
			System.setProperty("webdriver.chrome.driver", "D:\\Softwares\\chromedriver.exe");
			caps = DesiredCapabilities.chrome();
			getLogs(caps);
			driver = new ChromeDriver(caps);
		} else if (value.equals("firefox")) {
			System.out.println("Firefox constructed");
			/*System.setProperty("webdriver.gecko.driver", "D:\\Softwares\\geckodriver.exe");
			File pathToBinary = new File("C:\\Program Files (x86)\\Mozilla Firefox\\firefox.exe");
			FirefoxBinary ffBinary = new FirefoxBinary(pathToBinary);
			FirefoxProfile firefoxProfile = new FirefoxProfile();*/
			caps = DesiredCapabilities.firefox();
			getLogs(caps);
			driver = new FirefoxDriver(caps);
		} else if (value.equals("ie")) {
			System.out.println("IE constructed");
			System.setProperty("webdriver.ie.driver", "D:\\Softwares\\IEDriverServer.exe");
			caps = DesiredCapabilities.internetExplorer();
			getLogs(caps);
			driver = new InternetExplorerDriver(caps);
		} else if (value.equals("phantom")) {
			System.out.println("Phantom constructed");
			File file = new File("D:\\Softwares\\phantomjs-2.1.1-windows\\phantomjs-2.1.1-windows\\bin\\phantomjs.exe");
			System.setProperty("phantomjs.binary.path", file.getAbsolutePath());
			caps = DesiredCapabilities.phantomjs();
			getLogs(caps);
			driver = new PhantomJSDriver(caps);
		} else {
			Assert.assertTrue(false, "Invalid browser");
		}

	}

	private static void getLogs(DesiredCapabilities caps) throws Exception {
		LoggingPreferences logPrefs = new LoggingPreferences();
		logPrefs.enable(LogType.BROWSER, Level.SEVERE);
		caps.setCapability(CapabilityType.LOGGING_PREFS, logPrefs);
	}

	public List<String> analyzeLog() {
		LogEntries logEntries = driver.manage().logs().get(LogType.BROWSER);
		List<String> errors = new ArrayList<String>();
		for (LogEntry entry : logEntries) {
			System.out.println("JS Error: " + entry.getMessage());
			errors.add(entry.getMessage());
		}
		return errors;
	}

	@Test(dataProvider = "readURLFromExcel")
	public void testMethod(String url) throws Exception {
		String failureURL = "http://www.javascriptkit.com/javatutors/errortest.htm";
		System.out.println("URL : " + url);
		driver.get(url);
		Thread.sleep(1000);
		if (url.equals(failureURL)) {
			driver.switchTo().alert().accept();
		}
		driver.manage().window().maximize();
		Thread.sleep(1000);
		updateStatusInExcel(analyzeLog(), url);
	}

	public synchronized void updateStatusInExcel(List<String> errors, String url) throws Exception {
		File excel = new File("D:\\Karthick\\2017\\January\\UselessFile.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheetAt(0);
		DataFormatter df = new DataFormatter();
		Iterator<Row> row = sheet.iterator();
		XSSFCellStyle style = book.createCellStyle();
		XSSFCellStyle style1 = book.createCellStyle();
		while (row.hasNext()) {
			Row currentRow = row.next();
			if (df.formatCellValue(currentRow.getCell(0)).trim().equals(url.trim())) {
				if (errors.size() == 0) {
					currentRow.getCell(1).setCellValue("No");
					style.setAlignment(style.ALIGN_CENTER);
					style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
					style.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
					style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
					style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
					style.setFillForegroundColor(new XSSFColor(java.awt.Color.GREEN));
					style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					currentRow.getCell(1).setCellStyle(style);
					break;
				} else {
					currentRow.getCell(1).setCellValue("Yes");
					style.setAlignment(style.ALIGN_CENTER);
					style1.setAlignment(style.ALIGN_CENTER);
					currentRow.getCell(2).setCellValue(errors.size());
					currentRow.getCell(2).setCellStyle(style1);
					Iterator<String> itr = errors.iterator();
					StringBuilder str = new StringBuilder();
					while(itr.hasNext()) {
						str.append(itr.next().trim()+"\n");
					}
					currentRow.getCell(3).setCellValue(str.toString());
					style.setBorderBottom(XSSFCellStyle.BORDER_MEDIUM);
					style.setBorderTop(XSSFCellStyle.BORDER_MEDIUM);
					style.setBorderLeft(XSSFCellStyle.BORDER_MEDIUM);
					style.setBorderRight(XSSFCellStyle.BORDER_MEDIUM);
					style.setFillForegroundColor(new XSSFColor(java.awt.Color.RED));
					style.setFillPattern(CellStyle.SOLID_FOREGROUND);
					currentRow.getCell(1).setCellStyle(style);
					break;
				}
			} else {
				continue;
			}
		}
		FileOutputStream fos = new FileOutputStream(excel);
		book.write(fos);
		book.close();
		fos.flush();
		fos.close();
	}

	@DataProvider
	public synchronized Object[][] readURLFromExcel() throws Exception {
		File excel = new File("D:\\Karthick\\2017\\January\\UselessFile.xlsx");
		FileInputStream fis = new FileInputStream(excel);
		XSSFWorkbook book = new XSSFWorkbook(fis);
		XSSFSheet sheet = book.getSheetAt(0);
		List<String> urls = new ArrayList<String>();
		Object[][] data = null;
		DataFormatter df = new DataFormatter();
		Iterator<Row> row = sheet.iterator();
		while (row.hasNext()) {
			Row currentRow = row.next();
			if ((!df.formatCellValue(currentRow.getCell(0)).equals("Urls"))
					&& (!df.formatCellValue(currentRow.getCell(0)).trim().equals(""))) {
				urls.add(df.formatCellValue(currentRow.getCell(0)));
			}
		}
		book.close();
		fis.close();
		data = new Object[urls.size()][1];
		for (int i = 0; i < urls.size(); i++) {
			data[i][0] = urls.get(i);
		}

		return data;
	}

	@AfterTest
	public void tearDown() throws Exception {
		driver.quit();
	}

}
