package com.fb.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilterInputStream;
import java.io.FilterOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelActions {

	public static void main(String[] args) throws Exception {
		
		File f = new File("C:\\Users\\selladurai\\Desktop\\name.xlsx");
		FileInputStream f1 = new FileInputStream(f);
		
		Workbook w = new XSSFWorkbook(f1);
		 Sheet s = w.createSheet("students");
		Row r = s.createRow(0);
			 Cell c = r.createCell(4);
	c.setCellValue("sella");
	FileOutputStream d = new FileOutputStream(f);
w.write(d);
			
	}
	
}


