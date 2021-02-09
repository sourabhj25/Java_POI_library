package practice_poi;

import java.io.FileOutputStream;  
import java.io.OutputStream;  
import org.apache.poi.hssf.usermodel.HSSFWorkbook;  
import org.apache.poi.ss.usermodel.BorderStyle;  
import org.apache.poi.ss.usermodel.Cell;  
import org.apache.poi.ss.usermodel.CellStyle;  
import org.apache.poi.ss.usermodel.IndexedColors;  
import org.apache.poi.ss.usermodel.Row;  
import org.apache.poi.ss.usermodel.Sheet;  
import org.apache.poi.ss.usermodel.Workbook; 

public class CellPresentation {

	public static void main(String[] args) {
		
		Workbook wb = new HSSFWorkbook();
		
		
		try(OutputStream fileOutput = new FileOutputStream("JavaWorkbook4.xls")){
			
			Sheet sheet1 = wb.createSheet("sheet1");
			Row row = sheet1.createRow(1);
			Cell cell = row.createCell(1);
			cell.setCellValue("Subjects");
			
			// styling border of the cell
			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setBorderBottom(CellStyle.BORDER_THIN);
			cellStyle.setBottomBorderColor(IndexedColors.RED.getIndex());
			cellStyle.setBorderLeft(CellStyle.BORDER_THICK);
			cellStyle.setLeftBorderColor(IndexedColors.BLUE.getIndex());
			cellStyle.setBorderRight(CellStyle.BORDER_SLANTED_DASH_DOT);  
			cellStyle.setRightBorderColor(IndexedColors.BRIGHT_GREEN.getIndex());  
	        cellStyle.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED);
	        cellStyle.setTopBorderColor(IndexedColors.BLACK.getIndex());  
	        cell.setCellStyle(cellStyle);
	        
			wb.write(fileOutput);
			System.out.println("WB created");
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
		
	}
}
