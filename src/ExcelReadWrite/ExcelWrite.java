package ExcelReadWrite;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelWrite { 

	public static void main(String[] args) {
		
		try {
			
        File file = new File("/Users/sara2cena/Documents/Workbook1.xlsx");
		
		FileInputStream fis = new FileInputStream(file);
		
	    XSSFWorkbook wbk = new XSSFWorkbook(fis);
		
		XSSFSheet sheet = wbk.getSheetAt(0); //first column
		
		int rowcnt = sheet.getLastRowNum() + 1; //row count
		
		
		//Values you want to add
		String [] value = {"Fire", "70", "3000"};
		
		XSSFRow row = sheet.createRow(rowcnt); //create a row where you want to add these values
		
		/*In the for loop, i < number is determined by the number of values you want to add
		 
		 or by the number of values in the array you just created.
		 
		*/ 
		
		for(int i=0; i<3; i++) 
		{
			//Need cell object to store the value
			
			XSSFCell cell = row.createCell(i);
			
			cell.setCellValue(value[i]);
			
		}
		
		FileOutputStream fos = new FileOutputStream (file);
		
		wbk.write(fos); //write into Workbook
		
		fis.close(); //close FileInputStream
		fos.close(); //close FileOutputStream
		wbk.close(); //close Workbook
		
		
		} catch (Exception e) {
		System.out.println(e);
	}

	}
	
}
