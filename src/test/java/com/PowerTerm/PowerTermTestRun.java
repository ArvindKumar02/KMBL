package com.PowerTerm;

import java.net.MalformedURLException;
import java.net.URL;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.remote.DesiredCapabilities;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

import io.appium.java_client.windows.WindowsDriver;

public class PowerTermTestRun {

	public static WindowsDriver driver = null;

	@BeforeMethod
	public void SetUp() throws MalformedURLException {

		DesiredCapabilities cap = new DesiredCapabilities();

		cap.setCapability("app", "C:\\Program Files (x86)\\Ericom Software\\PowerTerm\\ptw32.exe");
		cap.setCapability("platform", "Windows");
		cap.setCapability("deviceName", "WindowsPC");
		// cap.setCapability("ms:waitForAppLaunch", "25"); -----> App Launch
		// Wait(max-50Sec)

		driver = new WindowsDriver (new URL("http://127.0.0.1:4723"), cap);

		driver.manage().timeouts().implicitlyWait(5000, TimeUnit.SECONDS);

	}

	@Test
	public void PowerTerm_Launch() throws InterruptedException {

		
		System.out.println("Connet PopUp");
		driver.findElementByName("Close").click();
		System.out.println("Connect PopUp Status: OK");

		System.out.println("Process On File");
		driver.findElementByName("File").click();
		driver.findElementByName("Print Setup...").click();
		driver.findElementByName("OK").click();
		System.out.println("FileStatus: OK");

		System.out.println("Proccess On Communication");
		driver.findElementByName("Communication").click();
		driver.findElementByName("Utilities").click();
		driver.findElementByName("Break").click();
		driver.findElementByName("OK").click();
		System.out.println("Communication Status: OK");

		System.out.println("Process On Options");
		driver.findElementByName("Options").click();
		driver.findElementByName("Power Pad Setup...").click();
		driver.findElementByName("Cancel").click();
		System.out.println("OPtions Status: OK");

		System.out.println("Process On Help");
		driver.findElementByName("Help").click();
		driver.findElementByName("About PowerTerm InterConnect Demo...").click();
		driver.findElementByName("OK").click();
		System.out.println("Help Status: OK");
			
		
		System.out.println("All Automated Properly");
	}

	@AfterMethod
	public void TearDown() {

		driver.quit();

	}

}
