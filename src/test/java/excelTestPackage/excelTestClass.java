package excelTestPackage;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class excelTestClass {
	public static void excelTestMethod() throws IOException {
		File f=new File("C:\\WorkSpace\\Excel.Test\\target\\TestExcel.xlsx");
		FileInputStream fIS= new FileInputStream(f);
		XSSFWorkbook wB= new XSSFWorkbook(fIS);
		XSSFSheet sheet=wB.getSheet("Sheet1");
		String data=null;
		for (int i = 0; i < sheet.getPhysicalNumberOfRows(); i++) {
			XSSFRow row=sheet.getRow(i);
			for(int j=0;j<row.getPhysicalNumberOfCells();j++) {
				XSSFCell cell=row.getCell(j);
			}
			
		}
		
		}
	
	public static void main(String[] args) throws IOException {
		excelTestMethod();
		
		
	}

}
