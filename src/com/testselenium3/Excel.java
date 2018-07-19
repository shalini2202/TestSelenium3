package com.testselenium3;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class Excel {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		

		String FilePath = "C:\\Users\\a631020\\eclipse-workspace\\TestSelenium3.xlsx";
		FileInputStream fs = new FileInputStream(FilePath);
		Workbook a =new XSSFWorkbook(fs);

		// TO get the access to the sheet
		Sheet sh = a.getSheet("Sheet1");
		
		String name = sh.getRow(0).getCell(0).getStringCellValue();
		System.out.println(name);

	}

}
