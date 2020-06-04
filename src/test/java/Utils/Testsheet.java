package Utils;

import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.Test;

public class Testsheet {
	 public XSSFWorkbook workbook;
	 public XSSFSheet sheet;
	    
  @BeforeClass
  public void beforeClass() throws IOException {
  workbook = new XSSFWorkbook("C:\\\\mySelenium\\\\projects\\\\ExcelReading\\\\src\\\\test\\\\java\\\\data\\\\TestData.xlsx");
  sheet = workbook.getSheet("Sheet1");
    }
  
  @Test
  public void readdata() throws IOException {
	    
	  System.out.println(System.getProperty("user.dir"));		 
	  int rowcount = sheet.getPhysicalNumberOfRows();
	  System.out.println("Total Row Count " + rowcount);
	  String name = sheet.getRow(1).getCell(0).getStringCellValue();
	  String place = sheet.getRow(1).getCell(1).getStringCellValue();
	  int age = (int) sheet.getRow(1).getCell(2).getNumericCellValue();
	 System.out.println("name " + name + " place " + place + " age " + age);
	  
  }
}
