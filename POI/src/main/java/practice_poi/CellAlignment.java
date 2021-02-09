package practice_poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import javax.swing.text.StyledEditorKit.BoldAction;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class CellAlignment {

	/**
	 * 		Create the workbook and align the cell 
	 * 
	 * @param args
	 * @throws FileNotFoundException
	 * @throws IOException
	 */
	public static void main(String[] args) throws FileNotFoundException, IOException {
		Workbook wb = new HSSFWorkbook();
		
		try(OutputStream fileOutput = new FileOutputStream("JavaWorkbook3.xls")){
			Sheet sheet1 = wb.createSheet("Sheet1");
			Row row = sheet1.createRow(0);
			Cell cell  = row.createCell(0);
			cell.setCellValue("Column TITLE");
			CellStyle cellStyle = wb.createCellStyle();
			
			
			// Alignment Left
			HSSFCellStyle style1 =  (HSSFCellStyle) wb.createCellStyle();
			sheet1.setColumnWidth(0, 8000);
			cell.setCellStyle(style1);
			
			// Change row Height
			sheet1.setColumnWidth(1, 4000);
			row = sheet1.createRow(1);
			cell = row.createCell(1);
			cell.setCellValue("height change");
			row.setHeight((short) 500);
			
			wb.write(fileOutput);		
			System.out.println("WB created");
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
	}
}
