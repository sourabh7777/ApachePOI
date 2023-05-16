package library;

import java.io.FileInputStream;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelFoorLoop {

	public static void main(String[] args) throws Exception {
		String filepath = ".\\TestData\\ApachePOITestData.xlsx";
		FileInputStream inputStream = new FileInputStream(filepath);
		XSSFWorkbook wb = new XSSFWorkbook(inputStream);
		XSSFSheet sheet = wb.getSheet("Sheet1");
		
		int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
		for(int i=1;i<rowCount;i++) {
			for(int j=1;j<rowCount;j++) {
				if(j==1) {
				System.out.print((int)sheet.getRow(i).getCell(j).getNumericCellValue());
				}
				else {
					System.out.print(sheet.getRow(i).getCell(j).getStringCellValue());
				}
				System.out.print(" ");
			}
			System.out.println("\n");
		}
		
		wb.close();

	}

}
