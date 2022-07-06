package ExcelUtilities;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import net.bytebuddy.asm.Advice.Return;

public class ExcelUtilities {

	public static Workbook book;
	public static Sheet sheet;
	public static String excel_path_data= "C:\\Users\\Prithiv\\eclipse-workspace\\Boohoo.com\\Excel\\TestDataBoohoo.xlsx";
	public static Object[][] getExcelData(String sheetName) {
		FileInputStream file = null;
		try {
			file = new FileInputStream(excel_path_data);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		try {
			book = WorkbookFactory.create(file);
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		}
		catch (IOException e) {
			e.printStackTrace();
		}
		sheet=book.getSheet(sheetName);
		Object [][] data = new Object[sheet.getLastRowNum()][sheet.getRow(0).getLastCellNum()];
		for (int i = 0; i < sheet.getLastRowNum(); i++) {
			for (int j = 0; j < sheet.getRow(0).getLastCellNum(); j++) {
				data [i][j]= sheet.getRow(i).getCell(j).toString();
			}
		}
		return data;
	}

}
