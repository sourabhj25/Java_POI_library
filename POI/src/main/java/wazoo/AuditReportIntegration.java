package wazoo;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import lombok.extern.java.Log;

@Log
public class AuditReportIntegration {

	public static void main(String[] args) throws FileNotFoundException, IOException {
		try (OutputStream fileOutput = new FileOutputStream("auditReport.xls")) {
			Workbook wb = new HSSFWorkbook();
			Sheet sheet = wb.createSheet("Sheet1");
			int row = 0;
			Row row1 = sheet.createRow(0);
			Cell cell1 = row1.createCell(0);
			CellStyle style = wb.createCellStyle();

			cell1.setCellValue("AUDIT REPORT");
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 5));
			cell1.setCellStyle(style);

			Row eventName = sheet.createRow(2);
			Cell cell2 = eventName.createCell(0);
			cell2.setCellValue("Event Name");
			sheet.addMergedRegion(new CellRangeAddress(2, 2, 0, 1));
			Cell cell2_value = eventName.createCell(1);
			cell2_value.setCellValue(eventAuditResponseDTO.getEventName());
			sheet.addMergedRegion(new CellRangeAddress(2, 2, 1, 5));

			Row eventAddress = sheet.createRow(3);
			Cell cell3 = eventAddress.createCell(0);
			cell3.setCellValue("Event Address");
			sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 1));
			Cell cell3_value = eventAddress.createCell(1);
			cell3_value.setCellValue(event.getAddress);
			sheet.addMergedRegion(new CellRangeAddress(3, 3, 1, 5));

			Row eventDate = sheet.createRow(4);
			Cell cell4 = eventDate.createCell(0);
			cell4.setCellValue("Event Date & Time");
			sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 1));
			Cell cell4_value = eventDate.createCell(0);
			cell4_value.setCellValue(event.geteventDate);
			sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 1));

			for () {
			Row ticketDetails = sheet.createRow(5);
				
			Cell cell5 = ticketDetails.createCell(0);
			cell5.setCellValue(tickets.);
			Cell cell6 = ticketDetails.createCell(1);
			cell6.setCellValue("Ticket Price(Service Charge) $");
			}

			Cell cell7 = sheet.createRow(7).createCell(0);
			cell7.setCellValue("Today's Sale");
			sheet.addMergedRegion(new CellRangeAddress(7, 7, 0, 5));
			cell7.setCellStyle(style);

			// =================================
			// Box Office
			// =================================
			Cell cell8 = sheet.createRow(8).createCell(0);
			cell8.setCellValue("Box Office");
			sheet.addMergedRegion(new CellRangeAddress(8, 8, 0, 1));
			cell8.setCellStyle(style);

			Row row2 = sheet.createRow(9);
			Cell cell9 = row2.createCell(2);
			cell9.setCellValue("Tickets Sold");
			sheet.addMergedRegion(new CellRangeAddress(9, 9, 2, 3));
			cell9.setCellStyle(style);

			Cell cell10 = row2.createCell(4);
			cell10.setCellValue("Ticket Sales($)");
			sheet.addMergedRegion(new CellRangeAddress(9, 9, 4, 5));
			cell10.setCellStyle(style);

			Row row3 = sheet.createRow(10);
			Cell cell11 = row3.createCell(2);
			cell11.setCellValue("Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(10, 10, 2, 3));
			cell11.setCellStyle(style);

			Cell cell12 = row3.createCell(4);
			cell12.setCellValue("Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(10, 10, 4, 5));
			cell12.setCellStyle(style);

			Cell cell13 = sheet.createRow(12).createCell(0);
			cell13.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(12, 12, 0, 1));
			cell13.setCellStyle(style);

			// ==================================
			// Web
			// ==================================

			Cell cell14 = sheet.createRow(13).createCell(0);
			cell14.setCellValue("Web");
			sheet.addMergedRegion(new CellRangeAddress(13, 13, 0, 1));
			cell14.setCellStyle(style);

			Row row5 = sheet.createRow(14);
			Cell cell15 = row5.createCell(2);
			cell15.setCellValue("Tickets Sold");
			sheet.addMergedRegion(new CellRangeAddress(14, 14, 2, 3));
			cell15.setCellStyle(style);

			Cell cell16 = row5.createCell(4);
			cell16.setCellValue("Ticket Sales($)");
			sheet.addMergedRegion(new CellRangeAddress(14, 14, 4, 5));
			cell16.setCellStyle(style);

			Row row6 = sheet.createRow(15);
			Cell cell17 = row6.createCell(2);
			cell17.setCellValue("Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(15, 15, 2, 3));
			cell17.setCellStyle(style);

			Cell cell18 = row6.createCell(4);
			cell18.setCellValue("Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(15, 15, 4, 5));
			cell18.setCellStyle(style);

			Cell cell19 = sheet.createRow(17).createCell(0);
			cell19.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(17, 17, 0, 1));
			cell19.setCellStyle(style);

			// ===================================
			// Today's Total Sale
			// ===================================
			Cell cell20 = sheet.createRow(18).createCell(0);
			cell20.setCellValue("Today's Total Sale");
			sheet.addMergedRegion(new CellRangeAddress(18, 18, 0, 5));
			cell20.setCellStyle(style);

			Row row7 = sheet.createRow(19);
			Cell cell21 = row7.createCell(2);
			cell21.setCellValue("Total Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(19, 19, 2, 3));
			cell21.setCellStyle(style);

			Cell cell22 = row7.createCell(4);
			cell22.setCellValue("Total Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(19, 19, 4, 5));
			cell22.setCellStyle(style);

			Cell cell23 = sheet.createRow(21).createCell(0);
			cell23.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(21, 21, 0, 1));
			cell23.setCellStyle(style);
			// ===================================

			Cell cell24 = sheet.createRow(22).createCell(0);
			cell24.setCellValue("Sales For Event");
			sheet.addMergedRegion(new CellRangeAddress(22, 22, 0, 5));
			cell24.setCellStyle(style);

			// =================================
			// Box Office(Sales For Event)
			// =================================
			Cell cell25 = sheet.createRow(23).createCell(0);
			cell25.setCellValue("Box Office");
			sheet.addMergedRegion(new CellRangeAddress(23, 23, 0, 1));
			cell25.setCellStyle(style);

			Row row8 = sheet.createRow(24);
			Cell cell26 = row8.createCell(2);
			cell26.setCellValue("Tickets Sold");
			sheet.addMergedRegion(new CellRangeAddress(24, 24, 2, 3));
			cell26.setCellStyle(style);

			Cell cell27 = row8.createCell(4);
			cell27.setCellValue("Ticket Sales($)");
			sheet.addMergedRegion(new CellRangeAddress(24, 24, 4, 5));
			cell27.setCellStyle(style);

			Row row9 = sheet.createRow(25);
			Cell cell28 = row9.createCell(2);
			cell28.setCellValue("Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(25, 25, 2, 3));
			cell28.setCellStyle(style);

			Cell cell29 = row9.createCell(4);
			cell29.setCellValue("Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(25, 25, 4, 5));
			cell29.setCellStyle(style);

			Cell cell30 = sheet.createRow(26).createCell(0);
			cell30.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(26, 26, 0, 1));
			cell30.setCellStyle(style);

			// ==================================
			// Web(Sales For Event)
			// ==================================

			Cell cell31 = sheet.createRow(27).createCell(0);
			cell31.setCellValue("Web");
			sheet.addMergedRegion(new CellRangeAddress(27, 27, 0, 1));
			cell31.setCellStyle(style);

			Row row10 = sheet.createRow(28);
			Cell cell32 = row10.createCell(2);
			cell32.setCellValue("Tickets Sold");
			sheet.addMergedRegion(new CellRangeAddress(28, 28, 2, 3));
			cell32.setCellStyle(style);

			Cell cell33 = row10.createCell(4);
			cell33.setCellValue("Ticket Sales($)");
			sheet.addMergedRegion(new CellRangeAddress(28, 28, 4, 5));
			cell33.setCellStyle(style);

			Row row11 = sheet.createRow(29);
			Cell cell34 = row11.createCell(2);
			cell34.setCellValue("Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(29, 29, 2, 3));
			cell34.setCellStyle(style);

			Cell cell35 = row11.createCell(4);
			cell35.setCellValue("Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(29, 29, 4, 5));
			cell35.setCellStyle(style);

			Cell cell36 = sheet.createRow(30).createCell(0);
			cell36.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(30, 30, 0, 1));
			cell36.setCellStyle(style);

			// ===================================
			// Today's Total Sale(Sales For Event)
			// ===================================
			Cell cell37 = sheet.createRow(31).createCell(0);
			cell37.setCellValue("Today's Total Sale");
			sheet.addMergedRegion(new CellRangeAddress(31, 31, 0, 5));
			cell37.setCellStyle(style);

			Row row12 = sheet.createRow(32);
			Cell cell38 = row12.createCell(2);
			cell38.setCellValue("Total Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(32, 32, 2, 3));
			cell38.setCellStyle(style);

			Cell cell39 = row12.createCell(4);
			cell39.setCellValue("Total Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(32, 32, 4, 5));
			cell39.setCellStyle(style);

			Cell cell40 = sheet.createRow(33).createCell(0);
			cell40.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(33, 33, 0, 1));
			cell40.setCellStyle(style);
			// ===================================

			Cell cell41 = sheet.createRow(34).createCell(0);
			cell41.setCellValue("Unsold Tickets");
			sheet.addMergedRegion(new CellRangeAddress(34, 34, 0, 5));
			cell41.setCellStyle(style);

			Row row13 = sheet.createRow(35);
			Cell cell42 = row13.createCell(2);
			cell42.setCellValue("Tickets Unsold");
			sheet.addMergedRegion(new CellRangeAddress(35, 35, 2, 3));
			cell42.setCellStyle(style);

			Cell cell43 = row13.createCell(4);
			cell43.setCellValue("Ticket Sales($)");
			sheet.addMergedRegion(new CellRangeAddress(35, 35, 4, 5));
			cell43.setCellStyle(style);

			Cell cell44 = sheet.createRow(36).createCell(0);
			cell44.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(36, 36, 0, 1));
			cell44.setCellStyle(style);

			for (int i = 1; i <= 5; i++) {
				sheet.autoSizeColumn(i);
				sheet.setColumnWidth(i, 5000);
			}
			style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
			wb.write(fileOutput);
			log.info("File Created...");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

	}
	
	{
		if (eventAuditResponseDTO.getSalesDTO().get(i).getTimeOfSales().compareTo("Today") == 0
				&& eventAuditResponseDTO.getSalesDTO().get(i).getVisibility().equals("BOX_OFFICE")) {
			
			// =================================
			// Box Office
			// =================================

			Cell cell8 = sheet.createRow(rowCount).createCell(0);
			cell8.setCellValue("Box Office");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 1));
			cell8.setCellStyle(style);
			rowCount++;

			Row ticketsSale = sheet.createRow(rowCount);
			Cell cell9 = ticketsSale.createCell(2);
			cell9.setCellValue("Tickets Sold");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell9.setCellStyle(style);

			Cell cell10 = ticketsSale.createCell(4);
			cell10.setCellValue("Ticket Sales($)");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell10.setCellStyle(style);
			rowCount++;

			Row row2 = sheet.createRow(rowCount);
			Cell cell9_value = row2.createCell(2);
			for (int j = 0; j < eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().size(); j++) {
				cell9_value.setCellValue(
						eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().get(j).getSalesCount());
				sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
				cell9_value.setCellStyle(style);

				Cell cell10_value = row2.createCell(4);
				cell10_value.setCellValue(
						eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().get(j).getTotal());
				sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
				cell10_value.setCellStyle(style);
				rowCount++;
			}

			Row row3 = sheet.createRow(rowCount);
			Cell cell11 = row3.createCell(2);
			cell11.setCellValue("Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell11.setCellStyle(style);

			Cell cell12 = row3.createCell(4);
			cell12.setCellValue("Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell12.setCellStyle(style);
			rowCount++;

			Row refund = sheet.createRow(rowCount);
			Cell cell11_value = refund.createCell(2);
			cell11_value.setCellValue(eventAuditResponseDTO.getSalesDTO().get(i).getTotalRefundedTickets());
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell11_value.setCellStyle(style);

			Cell cell12_value = refund.createCell(4);
			cell12_value.setCellValue(eventAuditResponseDTO.getSalesDTO().get(i).getTotalAmountOfRefund());
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell12_value.setCellStyle(style);
			rowCount++;

			int todaysSaleBoxOfficeTotalTickets = eventAuditResponseDTO.getSalesDTO().get(i)
					.getTotalTicketSold()
					- eventAuditResponseDTO.getSalesDTO().get(i).getTotalRefundedTickets();
			double todaysSaleBoxOfficeTotalAmount = eventAuditResponseDTO.getSalesDTO().get(i)
					.getTotalAmountOfTickets()
					- eventAuditResponseDTO.getSalesDTO().get(i).getTotalAmountOfRefund();
			Row todayTotal = sheet.createRow(rowCount);
			Cell cell13 = todayTotal.createCell(0);
			cell13.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 1));
			cell13.setCellStyle(style);

			Cell cell13_value1 = todayTotal.createCell(2);
			cell13_value1.setCellValue(todaysSaleBoxOfficeTotalTickets);
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell13_value1.setCellStyle(style);
			Cell cell13_value2 = todayTotal.createCell(4);
			cell13_value2.setCellValue(todaysSaleBoxOfficeTotalAmount);
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell13_value2.setCellStyle(style);
			rowCount++;
		}

		else if (eventAuditResponseDTO.getSalesDTO().get(i).getTimeOfSales().compareTo("Today") == 0
				&& eventAuditResponseDTO.getSalesDTO().get(i).getVisibility().equals("ONLINE")) {

			// ==================================
			// Web
			// ==================================
			rowCount++;
			Cell cell14 = sheet.createRow(rowCount).createCell(0);
			cell14.setCellValue("Web");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 1));
			cell14.setCellStyle(style);

			rowCount++;
			Row webTicketsSold = sheet.createRow(rowCount);
			Cell cell15 = webTicketsSold.createCell(2);
			cell15.setCellValue("Tickets Sold");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell15.setCellStyle(style);

			Cell cell16 = webTicketsSold.createCell(4);
			cell16.setCellValue("Ticket Sales($)");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell16.setCellStyle(style);

			rowCount++;

			Row row5 = sheet.createRow(rowCount);
			for (int j = 0; j < eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().size(); j++) {
				Cell cell15_value = row5.createCell(2);
				cell15_value.setCellValue(
						eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().get(j).getSalesCount());
				sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
				cell15_value.setCellStyle(style);

				Cell cell16_value = row5.createCell(4);
				cell16_value.setCellValue(
						eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().get(j).getTotal());
				sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
				cell16_value.setCellStyle(style);
				rowCount++;
			}

			Row row6 = sheet.createRow(rowCount);
			Cell cell17 = row6.createCell(2);
			cell17.setCellValue("Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell17.setCellStyle(style);

			Cell cell18 = row6.createCell(4);
			cell18.setCellValue("Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell18.setCellStyle(style);
			rowCount++;

			Row webRefund = sheet.createRow(rowCount);
			Cell cell17_value = webRefund.createCell(2);
			cell17_value.setCellValue(eventAuditResponseDTO.getSalesDTO().get(i).getTotalRefundedTickets());
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell17_value.setCellStyle(style);

			Cell cell18_value = webRefund.createCell(4);
			cell18_value.setCellValue(eventAuditResponseDTO.getSalesDTO().get(i).getTotalAmountOfRefund());
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell18_value.setCellStyle(style);
			rowCount++;

			int todaysSaleWebTotalTickets = eventAuditResponseDTO.getSalesDTO().get(i).getTotalTicketSold()
					- eventAuditResponseDTO.getSalesDTO().get(i).getTotalRefundedTickets();
			double todaysSaleWebTotalAmount = eventAuditResponseDTO.getSalesDTO().get(i)
					.getTotalAmountOfTickets()
					- eventAuditResponseDTO.getSalesDTO().get(i).getTotalAmountOfRefund();

			Row todayTotalWeb = sheet.createRow(rowCount);
			Cell cell19 = todayTotalWeb.createCell(0);
			cell19.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 1));
			cell19.setCellStyle(style);

			Cell cell19_value1 = todayTotalWeb.createCell(2);
			cell19_value1.setCellValue(todaysSaleWebTotalTickets);
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell19_value1.setCellStyle(style);
			Cell cell19_value2 = todayTotalWeb.createCell(4);
			cell19_value2.setCellValue(todaysSaleWebTotalAmount);
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell19_value1.setCellStyle(style);
			rowCount++;
		}
		
		else if (eventAuditResponseDTO.getSalesDTO().get(i).getTimeOfSales().compareTo("TodayTotal") == 0
				&& eventAuditResponseDTO.getSalesDTO().get(i).getVisibility().equals("Total")) {
			// ===================================
			// Today's Total Sale
			// ===================================
			rowCount++;
			Cell cell20 = sheet.createRow(rowCount).createCell(0);
			cell20.setCellValue("Today's Total Sale");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 5));
			cell20.setCellStyle(style);

			rowCount++;
			Row row7 = sheet.createRow(rowCount);
			Cell cell21 = row7.createCell(2);
			cell21.setCellValue("Total Refunded Tickets");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell21.setCellStyle(style);

			Cell cell22 = row7.createCell(4);
			cell22.setCellValue("Total Amount Refunded");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell22.setCellStyle(style);
			
			rowCount++;
			for (int k = 0; k < eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().size(); k++) {
				Row totalRefundBoxOfc = sheet.createRow(rowCount);
				Cell cell21_value = totalRefundBoxOfc.createCell(2);
				cell21_value.setCellValue(eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().get(k).getSalesCount());
				sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
				cell21_value.setCellStyle(style);

				Cell cell22_value = totalRefundBoxOfc.createCell(4);
				cell22_value.setCellValue(eventAuditResponseDTO.getSalesDTO().get(i).getTicketSales().get(k).getTotal());
				sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
				cell22_value.setCellStyle(style);
				rowCount++;
			}
			Row totalofTotalSales = sheet.createRow(rowCount);
			Cell cell23 = totalofTotalSales.createCell(0);
			cell23.setCellValue("Total");
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 0, 1));
			cell23.setCellStyle(style);
			
			Cell cell23_value1 = totalofTotalSales.createCell(2);
			cell23_value1.setCellValue(eventAuditResponseDTO.getSalesDTO().get(i).getTotalRefundedTickets());
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 2, 3));
			cell23_value1.setCellStyle(style);
			Cell cell23_value2 = totalofTotalSales.createCell(4);
			cell23_value2.setCellValue(eventAuditResponseDTO.getSalesDTO().get(i).getTotalAmountOfRefund());
			sheet.addMergedRegion(new CellRangeAddress(rowCount, rowCount, 4, 5));
			cell23_value2.setCellStyle(style);
		}
	}

}
