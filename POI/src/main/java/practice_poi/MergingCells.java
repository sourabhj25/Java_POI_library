package practice_poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class MergingCells {

	public static void main(String[] args) throws FileNotFoundException, IOException {

		Workbook wb = new HSSFWorkbook();
		
		Sheet sheet = wb.createSheet();
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Company Name and Address");
		
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
		
		try(OutputStream fileOutput = new FileOutputStream("JavaWorkbook6.xls")){
			wb.write(fileOutput);
			System.out.println("File Created...");
		}catch(Exception e){
			System.out.println(e.getMessage());
		}
	}
}
