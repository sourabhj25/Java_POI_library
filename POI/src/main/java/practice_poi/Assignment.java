package practice_poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.property.PropertyTable;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;

import lombok.extern.java.Log;

@Log
public class Assignment {
	public static void main(String[] args) throws FileNotFoundException, IOException {

		try (OutputStream fileOutput = new FileOutputStream("invoiceFile.xls")) {
			Workbook wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet1");
			Row row = sheet.createRow(0);
			Cell cell = row.createCell(0);
			CellStyle style = wb.createCellStyle();

			style.setFillBackgroundColor(IndexedColors.BLUE.getIndex());
			style.setFillPattern(CellStyle.BIG_SPOTS);
			style.setAlignment(HSSFCellStyle.ALIGN_CENTER);

			// Creating Font Setting
			Font font = wb.createFont();
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			font.setFontHeightInPoints((short) 14);
			cell.setCellValue("INVOICE");
			style.setFont(font);

			// Merge Cell
			sheet.addMergedRegion(new CellRangeAddress(0, 1, 0, 5));

			/*
			 * int rowCount = 5; for(int i=0; i<50; i++){ Row row1 =
			 * sheet.createRow(rowCount);
			 * 
			 * row1.createCell(0).setCellValue("Agile Soft Systems, Inc");
			 * rowCount++; }
			 */

			// add picture data to this workbook.
			InputStream is = new FileInputStream("/home/saurabh/Downloads/agile.jpeg");
			byte[] bytes = IOUtils.toByteArray(is);
			int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_JPEG);
			is.close();

			CreationHelper helper = wb.getCreationHelper();
			// Create the drawing patriarch. This is the top level container for
			// all shapes.
			Drawing drawing = sheet.createDrawingPatriarch();
			// add a picture shape
			ClientAnchor anchor = helper.createClientAnchor();
			// set top-left corner of the picture,
			// subsequent call of Picture#resize() will operate relative to it
			anchor.setCol1(0);
			anchor.setRow1(3);
			Picture pict = drawing.createPicture(anchor, pictureIdx);
			// auto-size picture relative to its top-left corner
			pict.resize(0.3);

			// Logo and Name
			sheet.addMergedRegion(new CellRangeAddress(4, 6, 0, 1));
			Row row1 = sheet.createRow(7);
			Row row2 = sheet.createRow(8);
			Row row3 = sheet.createRow(9);	
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			row1.createCell(0).setCellValue("Agile Soft Systems, Inc");
			style.setFont(font);
			row2.createCell(0).setCellValue("38930 Blacow Rd Ste B3");
			row3.createCell(0).setCellValue("Fremont, CA 94536");

			// Bill to
			sheet.createRow(6).createCell(3).setCellValue("Bill To : ");
			row1.createCell(3).setCellValue("My Energy, Inc.");
			row2.createCell(3).setCellValue("5425 Airport Blvd, Ste 100");
			row3.createCell(3).setCellValue("Boulder, CO 80301");

			// Invoice date
			Row row4 = sheet.createRow(11);
			Row row5 = sheet.createRow(12);
			Row row6 = sheet.createRow(13);
			Row row7 = sheet.createRow(14);
			row4.createCell(3).setCellValue("Invoice Date:");
			row5.createCell(3).setCellValue("Invoice #:");
			row6.createCell(3).setCellValue("Invoice Amount:");
			row7.createCell(3).setCellValue("Invoice Due:");
			// Invoice date
			row4.createCell(4).setCellValue("date");
			row5.createCell(4).setCellValue("asda5");
			row6.createCell(4).setCellValue("$550");
			row7.createCell(4).setCellValue("On Reciept");
			// Agreement and Period
			row4.createCell(0).setCellValue("PO/Agreement#");
			row5.createCell(0).setCellValue("Period");
			row4.createCell(1).setCellValue("7/7/17");
			row5.createCell(1).setCellValue("Jan 2018");

			
			// table
			row = sheet.createRow(20);
			row.createCell(0).setCellValue("Project");
			row.createCell(1).setCellValue("Resource");
			row.createCell(2).setCellValue("Period");
			row.createCell(3).setCellValue("Type");
			row.createCell(4).setCellValue("Amt Billed");
			font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
			style.setFont(font);
			
			row = sheet.createRow(36);
			row.createCell(0).setCellValue("Total Invoice");
			row.createCell(4).setCellValue("$500");

			// Logo and Name
			sheet.addMergedRegion(new CellRangeAddress(4, 6, 0, 1));
			sheet.createRow(39).createCell(0).setCellValue("Pay through Wells Fargo Bank:");
			sheet.createRow(41).createCell(0).setCellValue("Agile Soft Systems, Inc");
			sheet.createRow(42).createCell(0).setCellValue("38930 Blacow Rd Ste B3");
			sheet.createRow(43).createCell(0).setCellValue("Fremont, CA 94536");
			Row row8 = sheet.createRow(44);
			Row row9 = sheet.createRow(45);
			row8.createCell(0).setCellValue("Wells Fargo Bank Acct #");
			row9.createCell(0).setCellValue("ABA #");
			row8.createCell(1).setCellValue("123456789");
			row9.createCell(1).setCellValue("12345678");

			sheet.autoSizeColumn(0);
			for (int i = 1; i <= 6; i++) {
				sheet.setColumnWidth(i, 5000);
			}
			cell.setCellStyle(style);
			// Write the file
			wb.write(fileOutput);
			log.info("File Created...");

		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

}
