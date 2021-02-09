package practice_poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CreateWorkBook {

	/**
	 * 		Create Workbook, Create Sheets, and 
	 * 		Create Row & Cells in the workbook
	 * 
	 * @param args
	 * @throws IOException
	 */
	public static void main(String[] args) throws IOException {
		// Create Workbook (.XLS)
		Workbook wb  =  new HSSFWorkbook();
		
		// Create FileOutput Stream Object
		try(OutputStream fileOut = new FileOutputStream("JavaWorbook.xls")){
			
			// Create Sheets in workbook
			Sheet sheet1 = wb.createSheet("First Sheet");
			Sheet sheet2 =  wb.createSheet("Second Sheet");
			
			// Add Rows and Cells
			Row row = sheet1.createRow(0);
			Cell cell1 = row.createCell(0);
			cell1.setCellValue("Subject");
			
			Cell cell2 = row.createCell(1);
			cell2.setCellValue("Fees");
			
			Row row1 = sheet1.createRow(1);
			Cell cell3 = row1.createCell(0);
			cell3.setCellValue("JAVA");
			
			Cell cell4 = row1.createCell(1);
			cell4.setCellValue("25000");
			
			wb.write(fileOut);
			System.out.println("WB Created");
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
	}
}
