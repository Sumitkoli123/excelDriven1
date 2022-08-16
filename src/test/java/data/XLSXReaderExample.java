package data;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;

public class XLSXReaderExample {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		// TODO Auto-generated method stub
		
		File file = new File("C:\\Users\\Admin\\Documents\\Book1.xlsx"); //creating a new file instance  
		FileInputStream fis = new FileInputStream(file); //obtaining bytes from the file 
		
		//creating Workbook instance that refers to .xlsx file  

		XSSFWorkbook wb = new XSSFWorkbook(fis);
		XSSFSheet sheet = wb.getSheetAt(1); //creating a Sheet object to retrieve object
		Iterator<Row> itr = sheet.rowIterator(); //iterating over excel file  
		
		while (itr.hasNext())
		{
			Row row = itr.next();
			Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
		while (cellIterator.hasNext())
		{
			Cell cell = cellIterator.next();
			
		switch (cell.getCellType())
		{
		
		case STRING:    //field that represents string cell type
			System.out.print(cell.getStringCellValue() + "\t\t\t");  
			break;  
			case NUMERIC:    //field that represents number cell type  
			System.out.print(cell.getNumericCellValue() + "\t\t\t");  
			break;  
			default:
		}
		}
		
		System.out.println("");
		}
	}
	
}

