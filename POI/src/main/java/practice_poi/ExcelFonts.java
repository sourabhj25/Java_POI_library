package practice_poi;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelFonts {

	public static void main(String[] args) throws IOException {
		
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("sheet1");
		Row row = sheet.createRow(5);
		Cell cell = row.createCell(5);
		cell.setCellValue("Hello! Welcome To POI");
		
		// merge row
		sheet.addMergedRegion(new CellRangeAddress(5, 6, 5, 8));
		
		// Creating Font Setting
		Font font = wb.createFont();
		font.setBoldweight((short)12);
		font.setFontHeightInPoints((short)16);
		font.setItalic(true);
		font.setFontName("Times New Roman");
		
		CellStyle style = wb.createCellStyle();
		style.setFont(font);
		cell.setCellStyle(style);
		
		try(OutputStream fileOut = new FileOutputStream("JavaWorkbook7.xls")){
			wb.write(fileOut);
			System.out.println("File Created...");
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
		
	}
}
