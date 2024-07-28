package taskzen13;



import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcel {

	public static void main(String[] args) throws IOException {

		// Create object of workbook
		XSSFWorkbook wb = new XSSFWorkbook();

		// Create Sheet
		XSSFSheet sh = wb.createSheet("Employee Data");

		// Create Arraylist
		ArrayList<Object[]> empData = new ArrayList<>();
		empData.add(new Object[] { "Name", "Age", "Email" });
		empData.add(new Object[] { "John Doe", "30", "john@test.com" });
		empData.add(new Object[] { "Jane Doe", "28", "john@test.com" });
		empData.add(new Object[] { "Bob Smith", "35", "jacky@example.com" });
		empData.add(new Object[] { "Swapnil", "37", "swapnil@example.com" });

		int rownum = 0;

		// Outer loop for Rows
		for (Object[] emp : empData) {

			// For Rows
			XSSFRow row = sh.createRow(rownum++);
			int cellnum = 0;

			// Inner loop for columns
			for (Object value : emp) {

				// For Columns
				XSSFCell cell = row.createCell(cellnum++);
			}
		}
		// Give File PAth Where employees.xlsx will create
		String filepath = ".\\datafiles\\employees.xlsx";

		// Create Object of FileOutputStream
		FileOutputStream fos = new FileOutputStream(filepath);

		fos.close();
		System.out.println("Employee.xlsx file writtern successfully.");

	}
}