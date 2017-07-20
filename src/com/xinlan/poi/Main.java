package com.xinlan.poi;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) {
		try {
			write();
			
			readXLSXFile();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void readXLSXFile() throws IOException {
		InputStream ExcelFileToRead = new FileInputStream("D:/潘易.xlsx");
		XSSFWorkbook wb = new XSSFWorkbook(ExcelFileToRead);

		XSSFSheet sheet = wb.getSheet("附加");
		
		//wb.getSheet(arg0)
		
		XSSFRow row;
		XSSFCell cell;

		Iterator rows = sheet.rowIterator();
		while (rows.hasNext()) {
			row = (XSSFRow) rows.next();
			Iterator cells = row.cellIterator();
			while (cells.hasNext()) {
				cell = (XSSFCell) cells.next();

				if (cell.getCellType() == XSSFCell.CELL_TYPE_STRING) {
					System.out.print(cell.getStringCellValue() + " ");
				} else if (cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
					System.out.print(cell.getNumericCellValue() + " ");
				} else {
					// U Can Handel Boolean, Formula, Errors
				}
			}
			System.out.println();
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
				cell.setCellValue("海军——" + c);
			}
		}

		FileOutputStream fileOut = new FileOutputStream(excelFileName);
		// write this workbook to an Outputstream.
		wb.write(fileOut);
		fileOut.flush();
		fileOut.close();
	}
}
