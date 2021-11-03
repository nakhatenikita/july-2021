import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

public class readexcel {
	@Test
	public void readdatafromExcel() throws IOException   {
		//Read Data From Excel
		FileInputStream file = new FileInputStream("E:\\Selenium jar\\Sheetdata.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);
		//XSSFSheet sheet = workbook.getSheet("string"); //using string sheet name=creditionals
		XSSFSheet sheet = workbook.getSheetAt(0);  //using index sheet name=0,1,2...so on
		
		System.out.println(sheet.getRow(0).getCell(0).getNumericCellValue());
		System.out.println(sheet.getRow(1).getCell(1).getStringCellValue());
		
		//Writing data into excel
		Row row = sheet.createRow(10);
		Cell cell = row.createCell(5);
		cell.setCellValue("Utkarshaa Academy");
		FileOutputStream fos = new FileOutputStream("E:\\Selenium jar\\Sheetdata.xlsx");
		workbook.write(fos);
		fos.close();
		System.out.println("end of writing data in excel");
		
		
		
		
	
		
		

		
		
		
	}

}
