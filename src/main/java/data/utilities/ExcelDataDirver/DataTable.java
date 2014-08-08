package data.utilities.ExcelDataDirver;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import data.utilities.ExcelProperty;

public class DataTable {

 //private static final Cell Cell = null;
private static final int MY_MINIMUM_COLUMN_COUNT = 0;
private static final String excelFilePath = ExcelProperty.DataTable_Path;
private static final String excelSheetName = ExcelProperty.DataSheet_Name;

public static void main( String[] args ) throws IOException{
//	String csvFile = "./src/main/resources/TestDataSheet.xls";
	FileInputStream ExcelFile = new FileInputStream(excelFilePath);
	Workbook wb = new HSSFWorkbook(ExcelFile);
	Sheet sheet1 = wb.getSheet(excelSheetName);
	//int RowCount = sheet1.getLastRowNum();
	System.out.println("Number of Columns In Data Sheet:  " + getLastColNum(sheet1));	
	System.out.println("Number of Rows In Data Sheet:  " + getLastRowNum(sheet1));
	
}

public static int getLastColNum(Object DataSheet) {
  	 // Decide which rows to process
      int rowStart = Math.min(15, ((Sheet) DataSheet).getFirstRowNum());
    //  byte rowEnd = 1;// Math.max(1400, sheet1.getLastRowNum());
      Row r = ((Sheet) DataSheet).getRow(rowStart);
         int lastColumn = Math.max(r.getLastCellNum(), MY_MINIMUM_COLUMN_COUNT);
         return lastColumn;
  }

public static int getLastRowNum(Object DataSheet) {
 	int lastRow;
 	lastRow = ((Sheet) DataSheet).getLastRowNum();
        return lastRow;  	
 }

//This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method
/*public static void setExcelFile(String Path,String SheetName) throws Exception {
       try {
           // Open the Excel file
        FileInputStream ExcelFile = new FileInputStream(Path);
        // Access the required test data sheet
        HSSFWorkbook ExcelWBook = new HSSFWorkbook(ExcelFile);
        HSSFSheet Sheet = ExcelWBook.getSheet(SheetName);
        } catch (Exception e){
            throw (e);
        }
}
        
 */
	
} // class
