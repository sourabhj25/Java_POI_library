package practice_poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 		Add date to the cell 
 * @author saurabh
 *
 */
public class ExcelFormating1 {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		Workbook wb = new HSSFWorkbook();
		CreationHelper creationHelper = wb.getCreationHelper();
		
		try(OutputStream fileOutputStream = new FileOutputStream("JavaWorkbook2.xls")){
			Sheet sheet1 = wb.createSheet("Sheet1");
			Row row = sheet1.createRow(0);
			
			CellStyle cellStyle = wb.createCellStyle();
			cellStyle.setDataFormat(creationHelper.createDataFormat().getFormat("dd/MM/yyyy hh:mm"));
			Cell cell = row.createCell(0);
			cell.setCellValue(new Date());
			cell.setCellStyle(cellStyle);
			
			
			wb.write(fileOutputStream);
			System.out.println("WB Created");
		}catch(Exception e){
			System.out.println("Exception :" + e.getMessage());
		}
	}
}
