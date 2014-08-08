package data.utilities.ExcelDataDirver;

import java.io.FileInputStream;
//import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellReference;

public class App 
{
    public static void main( String[] args ) throws IOException
    {
        System.out.println( "Hello World!" );
        // Open the Excel file
        String csvFile = "C:\\Workspace\\ExcelDataDirver\\src\\main\\resources\\TestDataSheet.xls";
        FileInputStream ExcelFile = new FileInputStream(csvFile);
        Workbook wb = new HSSFWorkbook(ExcelFile);
        Sheet sheet1 = wb.getSheet("Global");
        
        for (Row row : sheet1) {
        	 System.out.print("Row Value In Loop is: "+ row);
            for (Cell cell : row) {
            	System.out.print("Cell Value In Loop is: "+ cell);
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                System.out.print(cellRef.formatAsString());
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
        
        //System.out.println(DataTable.getLastColNum(sheet1));
        
        
    }
}
