package com.data_driven_framework.qa;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import io.github.bonigarcia.wdm.WebDriverManager;

public class DataDriven_framework {
	
	@Test(dataProvider="wordpressData")
	public void login(String username,String password) throws InterruptedException {
		
		WebDriverManager.chromedriver().setup();
        WebDriver driver = new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://www.facebook.com/");
        
        driver.findElement(By.id("email")).sendKeys(username);
        driver.findElement(By.name("pass")).sendKeys(password);
        driver.findElement(By.name("login")).click();
        
        Thread.sleep(5000);
        
        System.out.println(driver.getTitle());
        
        driver.close();
		
	}
	
	@DataProvider(name="wordpressData")
    public Object[][] passData() {
        
        ExcelDataConfig config=new ExcelDataConfig("C:\\Users\\anura\\eclipse-workspace\\Data_Driven_Framework\\TestData\\UserName.xlsx");
        
        int rows=config.getRowcount(0);
        
        Object[][] data=new Object[rows][2];
        
        
        for(int i=0;i<rows;i++) {
            
            data[i][0]=config.getData(0, i, 0);
            
            data[i][1]=config.getData(0, i, 1);
            
        }
        return data;
        
    }

}
