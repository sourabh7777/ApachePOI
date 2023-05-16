package library;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ReadExcelData {

	public static void main(String[] args) throws Exception{
		
		String filepath = ".\\TestData\\ApachePOITestData.xlsx";
		FileInputStream inputStream = new FileInputStream(filepath);
		XSSFWorkbook wb = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = wb.getSheet("Sheet1");
		
		String firstName = sheet.getRow(1).getCell(2).getStringCellValue();
		String lastName = sheet.getRow(1).getCell(3).getStringCellValue();

		System.out.println(firstName +"   " + lastName);
		
		wb.close();

	}

}
