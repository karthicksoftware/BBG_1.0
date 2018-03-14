package Poc.BBG;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.List;
import java.util.logging.Level;

import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.HttpClientBuilder;
import org.json.JSONException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
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
import org.testng.annotations.Parameters;
import org.testng.annotations.Test;

public class PhantomTest {
	WebDriver driver = null;

	// @BeforeTest
	public void setup() throws Exception {
		File file = new File("D:\\Softwares\\phantomjs-2.1.1-windows\\phantomjs-2.1.1-windows\\bin\\phantomjs.exe");
		System.setProperty("phantomjs.binary.path", file.getAbsolutePath());
		System.setProperty("webdriver.chrome.driver", "D:\\Softwares\\chromedriver.exe");
	}

	@Test(enabled = false)
	@Parameters(value = { "browser" })
	public void sampleMethod(String value) throws Exception {
		if (value.equals("chrome"))
			driver = new ChromeDriver();
		else
			driver = new PhantomJSDriver();

		driver.get(
				"http://stackoverflow.com/questions/24138398/how-to-implement-phantomjs-with-selenium-webdriver-using-java");
		Thread.sleep(3000L);
		String title = driver.getTitle();
		String expected = "How to implement PhantomJS with Selenium WebDriver";
		Assert.assertTrue(title.contains(expected), "Title doesn't match");
	}

	@Test(enabled = false)
	public class Test1 {

		// String url = "https://httpbin.org/get?show_env=1";
		String url = "http://www.javascriptkit.com/javatutors/errortest.htm";

		public void getResponse() throws IOException, JSONException {
			HttpClient client = HttpClientBuilder.create().build();
			HttpGet request = new HttpGet(url);
			request.addHeader("accept", "application/json");
			HttpResponse response = client.execute(request);
			System.out.println("Reason Phrase: " + response.getStatusLine().getReasonPhrase());
			System.out.println("Status code : " + response.getStatusLine().getStatusCode());

			BufferedReader data = new BufferedReader(new InputStreamReader(response.getEntity().getContent()));
			StringBuffer result = new StringBuffer();
			String line = "";
			while ((line = data.readLine()) != null) {
				result.append(line + "\n");
			}
			System.out.println(result);
		}
	}

	@Test(enabled = false)
	public void getJSErrors() throws Exception {
		// System.setProperty("webdriver.gecko.driver", "C:\\Program
		// Files\\Mozilla Firefox\\firefox.exe");
		/*
		 * FirefoxBinary bin = new FirefoxBinary(new
		 * File("C:\\Program Files\\Mozilla Firefox\\firefox.exe"));
		 * FirefoxProfile firefoxProfile = new FirefoxProfile();
		 */
		FirefoxProfile ffProfile = new FirefoxProfile();

		final WebDriver driver = new FirefoxDriver(ffProfile);
		driver.get("http://www.javascriptkit.com/javatutors/errortest.htm");
		driver.switchTo().alert().dismiss();
		Thread.sleep(2000L);
		JavaScriptError.addExtension(ffProfile);
		final List<JavaScriptError> jsErrors = JavaScriptError.readErrors(driver);
		System.out.println("JS Errors: " + jsErrors.size());
	}

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
			caps = DesiredCapabilities.firefox();
			getLogs(caps);
			driver = new FirefoxDriver(caps);
		} else if (value.equals("ie")) {
			System.out.println("IE constructed");
			System.setProperty("webdriver.ie.driver", "D:\\Softwares\\IEDriverServer.exe");
			caps = DesiredCapabilities.internetExplorer();
			getLogs(caps);
			driver = new InternetExplorerDriver(caps);
		} else {
			Assert.assertTrue(false, "Invalid browser");
		}

	}

	private static void getLogs(DesiredCapabilities caps) throws Exception {
		LoggingPreferences logPrefs = new LoggingPreferences();
		logPrefs.enable(LogType.BROWSER, Level.SEVERE);
		caps.setCapability(CapabilityType.LOGGING_PREFS, logPrefs);
	}

	public void analyzeLog(String value, String URL) {
		System.out.println("Test what:" + value);
		LogEntries logEntries = driver.manage().logs().get(LogType.BROWSER);
		if (value.equals("chrome")) {
			for (LogEntry entry : logEntries) {
				System.out.println("JS Error: " + entry.getMessage());
			}
		} else if (value.equals("firefox")) {
			for (int i = 0; i < logEntries.getAll().size(); i++) {
				System.out.println("JS Error: " + logEntries.getAll().get(logEntries.getAll().size() - i - 1).getMessage());
			}
		} else {
			Assert.assertTrue(false, "Browser Doesn't matched");
		}

	}

	@Test
	@Parameters(value = { "browser" })
	public void testMethod(String browser) throws InterruptedException {
		String url = "http://www.javascriptkit.com/javatutors/errortest.htm";
		//String url = "http://www.rferl.org/";
		driver.get(url);
		driver.switchTo().alert().dismiss();
		driver.manage().window().maximize();
		Thread.sleep(2000);
		analyzeLog(browser, url);
	}

	@AfterTest
	public void tearDown() throws Exception {
		driver.quit();
	}
}
