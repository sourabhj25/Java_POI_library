package practice_poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CellBackgroundColour {

	public static void main(String[] args) throws FileNotFoundException, IOException {

		Workbook wb = new XSSFWorkbook();
		Sheet sheet1 = wb.createSheet();
		Row row = sheet1.createRow(0);
		Cell cell = row.createCell(0);

		CellStyle style = wb.createCellStyle();

		style.setFillBackgroundColor(IndexedColors.BRIGHT_GREEN.getIndex());
		style.setFillPattern(CellStyle.BIG_SPOTS);
		cell.setCellValue("Subject");
		cell.setCellStyle(style);
		
		try(OutputStream fileOut = new FileOutputStream("JavaWorkbook5.xls")){
			wb.write(fileOut);
			System.out.println("WB created...");
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
	}
}
