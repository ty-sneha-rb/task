package utils;

import java.io.IOException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUtils {
	
	static XSSFWorkbook workbook;
	static XSSFSheet sheet;
	
	public ExcelUtils(String excelpath, String sheetName) {
		try {
			workbook = new XSSFWorkbook(excelpath);
			sheet = workbook.getSheet(sheetName);
		}catch(Exception exp) {
			System.out.println(exp.getCause());
			System.out.println(exp.getMessage());
			exp.printStackTrace();
		}
	}

	public static void getCellData(int rowNum , int colNum) throws IOException {
		DataFormatter formatter = new DataFormatter();
		Object value = formatter.formatCellValue(sheet.getRow(rowNum).getCell(colNum));
		//double value = sheet.getRow(1).getCell(8).getNumericCellValue(); 
		//String value = sheet.getRow(1).getCell(0).getStringCellValue();
		System.out.println(value);
	}

	public static void getRowCount() {
		int rowCount = sheet.getPhysicalNumberOfRows();
		System.out.println("No of rows:"+rowCount);

	}
}