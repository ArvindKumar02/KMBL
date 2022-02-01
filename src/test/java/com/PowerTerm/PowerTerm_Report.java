package com.PowerTerm;

import java.net.MalformedURLException;
import java.net.URL;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

import com.aventstack.extentreports.ExtentReports;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import com.aventstack.extentreports.reporter.ExtentHtmlReporter;

import io.appium.java_client.windows.WindowsDriver;

public class PowerTerm_Report {
	public static WindowsDriver driver = null;
	ExtentHtmlReporter htmlReporter;
	ExtentReports extent;

	@BeforeTest
	public void SetUp() {

		htmlReporter = new ExtentHtmlReporter("PowerTerm.html");
		extent = new ExtentReports();
		extent.attachReporter(htmlReporter);
	}

	@BeforeMethod
	public void LaunchApp() throws MalformedURLException {
		
		ExtentTest test = extent.createTest("Application Launch", "Application Launch OK");

		DesiredCapabilities cap = new DesiredCapabilities();

		cap.setCapability("app", "C:\\Program Files (x86)\\Ericom Software\\PowerTerm\\ptw32.exe");
		cap.setCapability("platform", "Windows");
		cap.setCapability("deviceName", "WindowsPC");

		driver = new WindowsDriver(new URL("http://127.0.0.1:4723"), cap);
		
		test.info("Launching Application");
		

		driver.manage().timeouts().implicitlyWait(500, TimeUnit.SECONDS);

	}

	@Test
	public void Launch() throws InterruptedException {
		
		ExtentTest test = extent.createTest("Application Automation", "Application Automation OK");
		
		test.log(Status.INFO, "This shows usage of log(Powerterm, Pass)");

		test.info("Closing the PopUp Menu");
		driver.findElementByName("Close").click();
		
		

		test.info("Accessing File Tab");
		driver.findElementByName("File").click();
		driver.findElementByName("Print Setup...").click();
		driver.findElementByName("OK").click();
		

		test.info("Accessing Communication Tab");
		driver.findElementByName("Communication").click();
		driver.findElementByName("Utilities").click();
		driver.findElementByName("Break").click();
		driver.findElementByName("OK").click();
		System.out.println("Communication Status: OK");

		test.info("Acessing Options Tab");
		driver.findElementByName("Options").click();
		driver.findElementByName("Power Pad Setup...").click();
		driver.findElementByName("Cancel").click();
		System.out.println("OPtions Status: OK");

		test.info("Accessing Help Tab");
		driver.findElementByName("Help").click();
		driver.findElementByName("About PowerTerm InterConnect Demo...").click();
		driver.findElementByName("OK").click();
		System.out.println("Help Status: OK");

	}
	
	@AfterMethod
	private void Report() {
		
		extent.flush();

	}

	@AfterTest
	public void TearDown() {

		driver.quit();

	}

}
