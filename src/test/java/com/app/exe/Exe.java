package com.app.exe;

import org.openqa.selenium.By;

import com.app.base.BaseClass;

public class Exe extends BaseClass {
	
	public static void main(String[] args) {
		browserlaunch();
		input_Data(driver.findElement(By.id("username")), excel_Read_Base_ReUse(0,0));
		input_Data(driver.findElement(By.id("password")), excel_Read_Base_ReUse(1,0));
	//	driver.findElement(By.id("username")).sendKeys(excel_Read_Base_ReUse(0,0));
	//	driver.findElement(By.id("password")).sendKeys(excel_Read_Base_ReUse(1,0));
		click_Value(driver.findElement(By.id("login")));
		dropdpown_Value(driver.findElement(By.id("location")), excel_Read_Base_ReUse(2,0));
	}
	
	

}
