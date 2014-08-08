package data.utilities.ExcelDataDirver;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;

import junit.framework.Test;
import junit.framework.TestCase;
import junit.framework.TestSuite;

/**
 * Unit test for simple App.
 */
public class AppTest 
    extends TestCase
{
	private static Sheet sheet1 = null;
	
	static { // need to use static initialization block, because @BeforeClass doesn't work in junit if the test extends TestCase
        String csvFile = AppTest.class.getClassLoader().getResource("TestDataSheet.xls").getPath();
        System.out.println("path " + csvFile);
        Workbook wb = null;
        try {
			FileInputStream ExcelFile = new FileInputStream(csvFile);
			wb = new HSSFWorkbook(ExcelFile);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
        
        sheet1 = wb.getSheet("Global");
	}
	
    /**
     * Create the test case
     *
     * @param testName name of the test case
     */
    public AppTest( String testName )
    {
        super( testName );
    }

    /**
     * @return the suite of tests being tested
     */
    public static Test suite()
    {
        return new TestSuite( AppTest.class );
    }

    /**
     * Rigourous Test :-)
     */
    public void testApp()
    {
        assertTrue( true );
        
        if(sheet1 == null){
        	System.out.println("Sheet1 is null");
        }
        
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
    }
}
