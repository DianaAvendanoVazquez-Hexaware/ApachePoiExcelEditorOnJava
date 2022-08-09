package excelDemo;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class createExcelFile {

	public static void main(String[] args)  throws Exception{
		
		// workbook object
        XSSFWorkbook workbook = new XSSFWorkbook();
  
        // spreadsheet object
        XSSFSheet spreadsheet
            = workbook.createSheet("page1");
  
        // creating a row object
        XSSFRow row;
        
        //-------------------------------------------------------------------------
     // Creating an empty TreeMap
        
        Map<String, Object[]> page1Data = new TreeMap<String, Object[]>();
        //the key it's gonna be the number of the column 
        
        page1Data.put("1", new Object[] {"NAME", "LASTNAME", "EMAIL", "PASSWORD", "COMPANY", "ADDRESS", "CITY",
        		"ZIP_CODE", "MOBILE_PHONE"}); //TITLES
        
        page1Data.put("2", new Object[] {"SomeName", "SomeLastName", "SomePassword", "SomeCompany", "SomeAddress", "SomeCity",
        		"SomePostCode", "SomeMobilePhone"}); //contents
        
        Set<String> keyid = page1Data.keySet();
        
        int rowid = 0;
        
     // writing the data into the sheet
        
        for (String key : keyid) {
  
            row = spreadsheet.createRow(rowid++);
            Object[] objectArr = page1Data.get(key);
            int cellid = 0;
  
            for (Object obj : objectArr) {
                Cell cell = row.createCell(cellid++);
                cell.setCellValue((String)obj);
            }
        }
        
        
        //-------------------------------------------------------------------------
		
		//creates an excel file at the specified location  
        FileOutputStream out = new FileOutputStream(
        new File("C:\\Users\\iamCa\\Desktop\\NotPorn\\cursos\\Hexaware\\excelDemo\\Excel1.xlsx"));
        System.out.println("Excel File has been created successfully.");   //message to be sure it was created
      
        workbook.write(out);
        out.close();
	}

}
