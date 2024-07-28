package taskzen13;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel {

	public static void main(String[] args) throws IOException {

		// Specify the location of Excel File
    File src = new File ("F:\\Task13");

    // Load File
    FileInputStream fis = new FileInputStream(src);

    // Load Workbook
    XSSFWorkbook wb = new XSSFWorkbook(fis);

    // Load Worksheet
    XSSFSheet sh = wb.getSheet("Sheet1");

    // Print the name of Loaded sheet
    System.out.println(sh.getSheetName());

   // Print Username from Excel Sheet
    System.out.println(sh.getRow(0).getCell(0).getStringCellValue());

    // Print Total Number of Rows
    System.out.println("Total Rows : "+ sh.getPhysicalNumberOfRows());

    // Print Total Number of Columns
    System.out.println("Total Columns : "+ sh.getRow(0).getPhysicalNumberOfCells());

    int rows = (sh.getLastRowNum() - sh.getFirstRowNum()) +1;
    System.out.println("Total Rows : "+ rows);

    int columns = sh .getRow(0).getLastCellNum();
    System.out.println("Total Columns: "+ columns);

    // Print All Cells of Excel Sheet

    for (int i = 0; i < rows; i++) {
    	for (int j = 0; j < columns; j++) {
    		System.out.println(sh.getRow(i).getCell(j).getStringCellValue());
    	}
    }

	}

}