package Utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class ExcelUtils {
	  
	    
	
	public ExcelUtils() throws IOException {
		

		  String excelpath = "C:\\mySelenium\\projects\\ExcelReading\\src\\test\\java\\data\\TestData.xlsx"; 
		  XSSFWorkbook workbook = new XSSFWorkbook(excelpath);
		  XSSFSheet sheet = workbook.getSheet("Sheet1");	
		  System.out.println(System.getProperty("user.dir"));
		
		 
		  int rowcount = sheet.getPhysicalNumberOfRows();
		  System.out.println("Total Row Count " + rowcount);
		  String name = sheet.getRow(1).getCell(0).getStringCellValue();
		  String place = sheet.getRow(1).getCell(1).getStringCellValue();
		  int age = (int) sheet.getRow(1).getCell(2).getNumericCellValue();
		 System.out.println("name " + name + " place " + place + " age " + age);
	}
}

