package org.Exl;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ExcelConcept {
	public static void main(String[] args) throws Throwable {
			File F=new File("C:\\Users\\ASUS\\eclipse-workspace\\ExcelTask\\src\\test\\resources\\Info.xlsx");
			FileInputStream F1=new FileInputStream(F);
			Workbook W=new XSSFWorkbook(F1);
			Sheet S = W.getSheet("Sheet1");
			for(int i=0;i<S.getPhysicalNumberOfRows();i++) {
				Row R=S.getRow(i);
			for(int j=0;j<R.getPhysicalNumberOfCells();j++) {
				Cell C=R.getCell(j);
				int CellType=C.getCellType();
				if(CellType==1) {
					String V=C.getStringCellValue();
					System.out.println(V);
				}
				else if(CellType==0)
					if(DateUtil.isCellDateFormatted(C)) {
					Date D=C.getDateCellValue();
					SimpleDateFormat SDF=new SimpleDateFormat("DD/MM/YYYY");
					String Date=SDF.format(D);
					System.out.println(Date);
			}
					else {
						double CV=C.getNumericCellValue();
						Long l=(long)CV;
						String L=String.valueOf(l);
						System.out.println(L);
					}
			}
	}
	
	}
}
