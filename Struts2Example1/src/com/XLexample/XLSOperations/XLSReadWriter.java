package com.XLexample.XLSOperations;

//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileNotFoundException;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.sql.Date;
//import java.util.HashMap;
//import java.util.Iterator;
//import java.util.Map;
//import java.util.Set;
//import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.xssf.usermodel.XSSFSheet;
//import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * * Sample Java program to read and write Excel file in Java using Apache POI *
 */

public class XLSReadWriter {

	public static List<Object[]> uploadXLS(File inputFile, File outputFile) {
		// For storing data into CSV files
		List<Object[]> dataObjectList = new ArrayList();

		Object[] dataObject;
		int iCount = 0;

		StringBuffer data = new StringBuffer();
		try {
			FileOutputStream fos = new FileOutputStream(outputFile);

			// Get the workbook object for XLSX file
			XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(inputFile));

			// Get first sheet from the workbook
			XSSFSheet sheet = wBook.getSheetAt(0);
			Row row;
			Cell cell;

			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			
			if (iCount==0)
				dataObject = new Object[100];
			else
				dataObject = new Object[iCount];
			
			while (rowIterator.hasNext()) {
				row = rowIterator.next();

				// For each row, iterate through each columns
				Iterator<Cell> cellIterator = row.cellIterator();
				
				while (cellIterator.hasNext()) {

					cell = cellIterator.next();
					iCount++;
					switch (cell.getCellType()) {
					case Cell.CELL_TYPE_BOOLEAN:
						data.append(cell.getBooleanCellValue() + ",");

						break;
					case Cell.CELL_TYPE_NUMERIC:
						data.append(cell.getNumericCellValue() + ",");

						break;
					case Cell.CELL_TYPE_STRING:
						data.append(cell.getStringCellValue() + ",");
						break;

					case Cell.CELL_TYPE_BLANK:
						data.append("" + ",");
						break;
					default:
						data.append(cell + ",");

					}
				}
				if (dataObject.length == 100)				
					dataObject = Arrays.copyOf(dataObject, iCount);
			}

			fos.write(data.toString().getBytes());
			fos.close();
			return dataObjectList;

		} catch (Exception ioe) {
			ioe.printStackTrace();
		}
		return null;
	}

	public static void writeXLS(List<Object[]> dataObject, File outputFile) {

		try {
			FileOutputStream fos = new FileOutputStream(outputFile);

			// Get the workbook object for XLSX file
			XSSFWorkbook wBook = new XSSFWorkbook(new FileInputStream(outputFile));

			// Get first sheet from the workbook
			XSSFSheet sheet = wBook.getSheetAt(0);
			Row row = null;

			// Iterate through each rows from first sheet
			Iterator<Row> rowIterator = sheet.iterator();
			int rownum = sheet.getLastRowNum();
			int cellnum = 0;
			for (Object[] rowObj : dataObject) {
				for (Object obj : rowObj) {
					Cell cell = row.createCell(cellnum++);
					if (obj instanceof String) {
						cell.setCellValue((String) obj);
					} else if (obj instanceof Boolean) {
						cell.setCellValue((Boolean) obj);
					} else if (obj instanceof Date) {
						cell.setCellValue((Date) obj);
					} else if (obj instanceof Double) {
						cell.setCellValue((Double) obj);
					}
				}

			}

			wBook.write(fos);
			System.out.println("Writing on Excel file Finished ...");
			fos.close();
			wBook.close();
		} catch (FileNotFoundException fe) {
			fe.printStackTrace();
		} catch (IOException ie) {
			ie.printStackTrace();
		}

	}
}