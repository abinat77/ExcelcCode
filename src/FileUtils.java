import java.io.FileInputStream;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.formula.CollaboratingWorkbooksEnvironment;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import org.testng.annotations.Test;


public class FileUtils {    
	
	 static XSSFSheet sheet= null;
	final static String path ="C:\\Users\\User\\cucucmberWorkshop\\Excel\\book.xlsx";
	public static int rowNum =0; 
	
	
	@Test
	public static void test1() throws IOException, Exception, Throwable {
			writeData( "Sheet1","nlesh","Pincode","Radha");
		}  

	public static void writeData( String sheetName, String rowname, String cellNum, String value)
			throws IOException,EncryptedDocumentException, OpenXML4JException{
	  
		FileInputStream fis = new FileInputStream(path);
		Workbook wb = WorkbookFactory.create(fis);
		 int col_Num = -1;
         sheet = (XSSFSheet) wb.getSheet(sheetName);
         Row row = sheet.getRow(0);
         for(int m=1 ; m<= sheet.getPhysicalNumberOfRows() ;m++) {
         	
     	 	Row rows = CellUtil.getRow(m, sheet);
     	    Cell cell = CellUtil.getCell(rows,0);
     	     System.out.println(cell.toString());
     	     
     	     if(cell.toString().equalsIgnoreCase(rowname)) {
     	    	 rowNum = m+1;
     	    	 break;
     	     } 
      }
     
         for (int i = 0; i < row.getLastCellNum(); i++) {
             if (row.getCell(i).getStringCellValue().trim().equals(cellNum))
             {
                 col_Num = i;
             }
         }   
         sheet.autoSizeColumn(col_Num);
         row = sheet.getRow(rowNum - 1);
         if(row==null)
             row = sheet.createRow(rowNum - 1);

         Cell cell = row.getCell(col_Num);
         if(cell == null)
             cell = row.createCell(col_Num);

         cell.setCellValue(value);
        FileOutputStream fos = new FileOutputStream(path);
        wb.write(fos);
        fos.close();
       
        
     }

}