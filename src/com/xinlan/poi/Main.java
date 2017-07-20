package com.xinlan.poi;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) {
		try {
			write();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void write() throws IOException {
		String excelFileName = "D:/潘易.xlsx";// name of excel file

		String sheetName = "Sheet1";// name of sheet

		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet(sheetName);

		// iterating r number of rows
		for (int r = 0; r < 10; r++) {
			XSSFRow row = sheet.createRow(r);
			// iterating c number of columns
			for (int c = 0; c < 5; c++) {
				XSSFCell cell = row.createCell(c);
				cell.setCellValue("Cell " + r + " " + c);
			}
		}
		
		XSSFSheet sheet2 = wb.createSheet("附加");

		// iterating r number of rows
		for (int r = 0; r < 10; r++) {
			XSSFRow row = sheet2.createRow(r);
			// iterating c number of columns
			for (int c = 0; c < 5; c++) {
				XSSFCell cell = row.createCell(c);
				cell.setCellValue("海军——"+c);
			}
		}
		
		FileOutputStream fileOut = new FileOutputStream(excelFileName);
		// write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}
}
