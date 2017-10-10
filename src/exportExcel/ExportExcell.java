package exportExcel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import exportExcel.service.ExcelService;
import exportExcel.service.ExcelserviceImple;



/**
 * Servlet implementation class ExportExcell
 */
public class ExportExcell extends HttpServlet {
	private static final long serialVersionUID = 1L;
       
    /**
     * @see HttpServlet#HttpServlet()
     */
    public ExportExcell() {
        super();
        // TODO Auto-generated constructor stub
    }

	/**
	 * @see HttpServlet#doGet(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doGet(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		ExcelService excelService = new ExcelserviceImple();
		//Create blank workbook
	      XSSFWorkbook workbook = new XSSFWorkbook(); 
	      //Create a blank sheet
	      XSSFSheet spreadsheet = workbook.createSheet( 
	      " Employee Info ");
	      //Create row object
	      XSSFRow row;
	      //This data needs to be written (Object[])
	      Map < String, Object[] > empinfo = 
	      new TreeMap < String, Object[] >();
	      empinfo.put( "1", new Object[] { 
	      "EMP ID", "EMP NAME", "DESIGNATION" });
	      empinfo.put( "2", new Object[] { 
	      "tp01", "Gopal", "Technical Manager" });
	      empinfo.put( "3", new Object[] { 
	      "tp02", "Manisha", "Proof Reader" });
	      empinfo.put( "4", new Object[] { 
	      "tp03", "Masthan", "Technical Writer" });
	      empinfo.put( "5", new Object[] { 
	      "tp04", "Satish", "Technical Writer" });
	      empinfo.put( "6", new Object[] { 
	      "tp05", "Krishna", "Technical Writer" });
	      //Iterate over data and write to sheet
	      Set < String > keyid = empinfo.keySet();
	      int rowid = 0;
	      for (String key : keyid)
	      {
	         row = spreadsheet.createRow(rowid++);
	         Object [] objectArr = empinfo.get(key);
	         int cellid = 0;
	         for (Object obj : objectArr)
	         {
	            Cell cell = row.createCell(cellid++);
	            cell.setCellValue((String)obj);
	         }
	      }
	      XSSFSheet spreadsheet1 = workbook.createSheet( 
	    	      "cccc");
	      XSSFRow r = spreadsheet1.createRow(2);
	      Cell ce  = r.createCell(2);
	      ce.setCellValue("XXXXXXXXXXXX");
	      workbook = excelService.setTitle(workbook, spreadsheet1.getSheetName(), 2, 1, "red", "black", 40, "TOTLE");
	      System.out.println( 
	      "Writesheet.xlsx written successfully" );
	      OutputStream os = response.getOutputStream(); ;
          response.setContentType("application/vnd.ms-excel");
          response.setHeader("Content-Disposition", "attachment; filename=myexcel.xls");
          workbook.write(os);
           
           os.flush();
          
	//	response.getWriter().append("Served at: ").append(request.getContextPath());
	}

	/**
	 * @see HttpServlet#doPost(HttpServletRequest request, HttpServletResponse response)
	 */
	protected void doPost(HttpServletRequest request, HttpServletResponse response) throws ServletException, IOException {
		// TODO Auto-generated method stub
		doGet(request, response);
	}

}
