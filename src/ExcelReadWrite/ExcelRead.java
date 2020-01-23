package ExcelReadWrite;
import java.io.File;
import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) { 
		
		//Read the workbook by creating an object of Workbook
		
		try {
		
		File file = new File("/Users/sara2cena/Documents/Workbook1.xlsx");
		
		FileInputStream fis = new FileInputStream(file);
		
		XSSFWorkbook wbk = new XSSFWorkbook(fis);
		
		XSSFSheet sheet = wbk.getSheetAt(0); //first column
		
		int rowcnt = sheet.getLastRowNum() + 1; //row count
		
		for(int i=0; i<rowcnt; i++)
		{
			XSSFRow row = sheet.getRow(i);
			
			for(int j=0; j<row.getLastCellNum(); j++)
			{
				System.out.print(row.getCell(j) + "|"); //Cell value
				
				//"|" symbol is a separator between each cell value
				
				//we use print instead of println because we want the row to be printed in one line
					
			}
			
			System.out.println(); //adds new line after printing each row 
			
			}
		
		fis.close(); //FileInputStream will be closed
		wbk.close(); //Workbook will be closed
		
		
			
		} catch (Exception e) {
			System.out.println(e);
		}
		
		

	}

}
