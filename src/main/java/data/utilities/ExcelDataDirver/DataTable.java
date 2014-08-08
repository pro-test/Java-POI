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


public static void main( String[] args ) throws IOException{
	String csvFile = "C:\\Workspace\\ExcelDataDirver\\src\\main\\resources\\TestDataSheet.xls";
	FileInputStream ExcelFile = new FileInputStream(csvFile);
	Workbook wb = new HSSFWorkbook(ExcelFile);
	Sheet sheet1 = wb.getSheet("Global");
	//int RowCount = sheet1.getLastRowNum();
	System.out.println(getLastColNum(sheet1));	
	System.out.println(getLastRowNum(sheet1));
	
}

public static int getLastColNum(Object DataSheet) {
  	 // Decide which rows to process
      int rowStart = Math.min(15, ((Sheet) DataSheet).getFirstRowNum());
    //  byte rowEnd = 1;// Math.max(1400, sheet1.getLastRowNum());
      Row r = ((Sheet) DataSheet).getRow(rowStart);
         int lastColumn = Math.max(r.getLastCellNum(), MY_MINIMUM_COLUMN_COUNT);
         return lastColumn;
       //  System.out.println("Last Column aa" + lastColumn);
       	
  }

public static int getLastRowNum(Object DataSheet) {
 	int lastRow;
 	lastRow = ((Sheet) DataSheet).getLastRowNum();
        return lastRow;  	
 }

//This method is to set the File path and to open the Excel file, Pass Excel Path and Sheetname as Arguments to this method
public static void setExcelFile(String Path,String SheetName) throws Exception {
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




/*

	  throws IOException
	    {
	        System.out.println( "Hello World!" );
	        // Open the Excel file
	        String csvFile = "C:\\Workspace\\ExcelDataDirver\\src\\main\\resources\\TestDataSheet.xls";
	        FileInputStream ExcelFile = new FileInputStream(csvFile);
	        Workbook wb = new HSSFWorkbook(ExcelFile);
	        Sheet sheet1 = wb.getSheet("Global");
	        int RowCount = sheet1.getLastRowNum();
	        //short ColCount = sheet1.getLeftCol(); 
	        System.out.println(RowCount);
	        //System.out.println(ColCount);
	         
	        
	        System.out.println(getLastCol(sheet1));
	        
	        /*
	        for (Row row : sheet1) {
	        	// System.out.print("Row Value In Loop is: "+ row);
	          for (Cell cell : row) {
	            	//System.out.print("Cell Value In Loop is: "+ cell);
	                CellReference cellRef = new CellReference(1,cell.getColumnIndex())  ; //  (row.getRowNum(), cell.getColumnIndex());
	                System.out.println(cellRef.formatAsString());
	                System.out.print(" - ");

	                switch (cell.getCellType()) {
	                    case Cell.CELL_TYPE_STRING:
	                        System.out.println(cell.getRichStringCellValue().getString());
	                        break;
	                    case Cell.CELL_TYPE_NUMERIC:
	                        if (DateUtil.isCellDateFormatted(cell)) {
	                            System.out.println(cell.getDateCellValue());
	                        } else {
	                            System.out.println(cell.getNumericCellValue());
	                        }
	                        break;
	                    case Cell.CELL_TYPE_BOOLEAN:
	                        System.out.println(cell.getBooleanCellValue());
	                        break;
	                    case Cell.CELL_TYPE_FORMULA:
	                        System.out.println(cell.getCellFormula());
	                        break;
	                    default:
	                        System.out.println();
	                }
	           }
	       
	        }
	        */ 
	        

	        
	    
	        // Decide which rows to process
	        //int rowStart = Math.min(15, sheet1.getFirstRowNum());
	        //byte rowEnd = 1; // Math.max(1400, sheet1.getLastRowNum());
	        //int MY_MINIMUM_COLUMN_COUNT = 1;
	      //  for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
	        //   Row r = sheet1.getRow(rowNum);
	        //   int lastColumn = Math.max(r.getLastCellNum(), MY_MINIMUM_COLUMN_COUNT);
	        //   System.out.println("Last Column aa" + lastColumn);
	        //   for (int cn = 0; cn < lastColumn; cn++) {
	        //	   System.out.println(cn);
	        //      Cell c = r.getCell(cn, Row.RETURN_BLANK_AS_NULL);
	         //     if (c == null) {
	                 // The spreadsheet is empty in this cell
	          //    } else {
	                 // Do something useful with the cell's contents
	          //    }
	         //  }
	         //  System.out.println("Last Column" + lastColumn);
	      //  }
	        
	        

	    
	 //   } //Main Method
	 
	 
 
        
 
	
} // class
