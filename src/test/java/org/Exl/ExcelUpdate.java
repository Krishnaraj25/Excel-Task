package org.Exl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelUpdate {
	public static void main(String[] args) throws Throwable {
		File F=new File("C:\\Users\\ASUS\\eclipse-workspace\\ExcelTask\\src\\test\\resources\\Update.xlsx");
		FileInputStream F1=new FileInputStream(F);
		Workbook W=new XSSFWorkbook(F1);
		Sheet S=W.getSheet("Sheet1");
		for(int i=0;i<S.getPhysicalNumberOfRows();i++) {
		Row R=S.getRow(i);
		for(int j=0;j<R.getPhysicalNumberOfCells();j++) {
			Cell C=R.getCell(j);
			int CellType=C.getCellType();
			if(CellType==1) {
				String Value=C.getStringCellValue();
				if(Value.equals("Krishna")){
					C.setCellValue("Giri");
				}
				else if(Value.equals("Mugunthan")){
					C.setCellValue("Mugi");
				}
				else if(Value.equals("Praveen")) {
					C.setCellValue("Ram");
				}
				else if(Value.equals("Dhiva")){
					C.setCellValue("Dhivagar");
					
				}
			}
		}
		}
		FileOutputStream F2=new FileOutputStream(F);
		W.write(F2);
	}

}
