package exportExcel.service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public interface ExcelService {

	public XSSFWorkbook setTitle(XSSFWorkbook workBook,String sheet, int rowNumber, int cellNumber, String color, String backgroundColor, int fondSize,String title);
	public XSSFWorkbook SetBodycContent(Object content,XSSFWorkbook workBook,String sheet, int rowNumberStart, int cellNumberStart);
}
