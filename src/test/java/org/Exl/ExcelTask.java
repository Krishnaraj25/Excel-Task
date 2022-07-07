package org.Exl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelTask {
	public static void main(String[] args) throws Throwable {
		Scanner sc=new Scanner(System.in);
		File F=new File("C:\\Users\\ASUS\\eclipse-workspace\\ExcelTask\\src\\test\\resources\\Write.xlsx");
		Workbook W=new XSSFWorkbook();
		Sheet S=W.createSheet("Excel");
		for(int i=0;i<4;i++) {
		Row R=S.createRow(i);
		System.out.println("Row No:"+i);
		for(int j=0;j<3;j++) {
			Cell C=R.createCell(j);
			System.out.println("Enter the Value of Cell:"+j);
			C.setCellValue(sc.next());
		}
		}
		FileOutputStream F1=new FileOutputStream(F);
		W.write(F1);
		
	}

}
