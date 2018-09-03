package demos;

import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MergingCells {

	public static void main(String[] args) throws IOException {
		Workbook wb = new XSSFWorkbook(); // or new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");

		CellStyle style = wb.createCellStyle();
		Font font = wb.createFont();
		font.setBold(true);
		font.setFontName("Arial");
		font.setColor(IndexedColors.GREEN.getIndex());
		font.setFontHeightInPoints((short) 18);
		style.setFont(font);
		style.setVerticalAlignment(VerticalAlignment.CENTER);

		style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

		style.setBorderBottom(BorderStyle.MEDIUM);
		style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
		style.setBorderRight(BorderStyle.MEDIUM);
		style.setRightBorderColor(IndexedColors.BLACK.getIndex());
		style.setWrapText(true);

		Row row = sheet.createRow((short) 1);
		Cell cell = row.createCell((short) 1);
		cell.setCellStyle(style);
		cell.setCellValue(new XSSFRichTextString("This is a test"));
		sheet.addMergedRegion(new CellRangeAddress(1, 3, 1, 2));
		
		RegionUtil.setBorderTop(BorderStyle.MEDIUM, sheet.getMergedRegion(0), sheet);
		RegionUtil.setBorderRight(BorderStyle.THICK, sheet.getMergedRegion(0), sheet);
		RegionUtil.setBorderBottom(BorderStyle.THICK, sheet.getMergedRegion(0), sheet);
		RegionUtil.setBottomBorderColor(IndexedColors.AQUA.getIndex(), sheet.getMergedRegion(0), sheet);
		
		Row row4 = sheet.createRow((short) 4);
		row4.createCell(0).setCellValue(" hello hello fdasfds dfafdff ");
		row4.createCell(1).setCellValue(" hello hello fdasfds dfafdff");
		row4.createCell(2).setCellValue(" hello hello \r\n fdasfds dfafdff");
		row4.createCell(3).setCellValue(" hello hello fdasfds dfafdffhello");
		row4.createCell(4).setCellValue(" hello hello \r\n fdasfds dfafdff");
		row4.getCell(3).setCellStyle(style);
        //设置自动列宽
        /*for (int i = 0; i < 4; i++) {
        	sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i)*17/10);
        }*/
		row.setRowStyle(style);
		row4.setRowStyle(style);
		
		// Write the output to a file
		FileOutputStream fileOut = new FileOutputStream("D:\\merging_cells.xlsx");
		wb.write(fileOut);
		fileOut.close();
		wb.close();
	}

}
