package utils;

import java.io.IOException;

public class ExcelUtilsTest {
	
public static void main(String[] args) throws IOException {
	
	String excelpath = "./data/task.xlsx.xltx";
	String sheetName = "Sheet1";
	ExcelUtils excel = new ExcelUtils(excelpath , sheetName);
	
	excel.getRowCount();
	excel.getCellData(1,0);
	excel.getCellData(1,1);
	excel.getCellData(1,2);
	excel.getCellData(1,3);
	excel.getCellData(1,4);
	excel.getCellData(1,5);
	excel.getCellData(1,6);
	excel.getCellData(1,7);
	
}
}
