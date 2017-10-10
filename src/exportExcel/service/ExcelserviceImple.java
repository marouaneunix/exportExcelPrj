package exportExcel.service;

import java.util.List;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelserviceImple implements ExcelService{

	@Override
	public XSSFWorkbook setTitle(XSSFWorkbook workBook, String sheet, int rowNumber, int cellNumber, String color,
			String backgroundColor, int fondSize,String title) {
		 // Create a new font and alter it.
	    Font font = workBook.createFont();
	    font.setFontHeightInPoints((short)fondSize);
	    font.setFontName("Courier New");
	    font.setItalic(true);

	    // Fonts are set into a style so create a new one to use.
	    CellStyle style = workBook.createCellStyle();
	    style.setFont(font);
	    style.setFillBackgroundColor(HSSFColor.GOLD.index);
	    style.setFillForegroundColor(HSSFColor.BLUE.index);//(new XSSFColor(new java.awt.Color(128, 0, 128)));
	    style.setFillPattern(CellStyle.SOLID_FOREGROUND);
	    XSSFSheet sheett = workBook.getSheet(sheet) != null ? workBook.getSheet(sheet) : workBook.createSheet(sheet);
	    XSSFRow row = sheett.getRow(rowNumber) != null ? sheett.getRow(rowNumber) : sheett.createRow(rowNumber);
	    XSSFCell cell = row.getCell(cellNumber) != null ? row.getCell(cellNumber) : row.createCell(cellNumber);
		cell.setCellStyle(style);
		cell.setCellValue(title);
		return workBook;
	}

	@Override
	public XSSFWorkbook SetBodycContent(Object content, XSSFWorkbook workBook, String sheet, int rowNumberStart,
			int cellNumberStart) {
		List<String> lites = (List<String>) content;
		return null;
	}

}
