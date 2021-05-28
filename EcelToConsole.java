package com.midtree.Files.ExcelToConsole;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EcelToConsole {

	public static void main(String[] args) {
		readExcel("Rajastan");
	}

	private static void readExcel(String string) {
		try {
			File file = new File("C:\\Users\\M1064478\\Desktop\\Documents Teams\\Shreevasta.xlsx"); // creating
																									// a
																									// new
																									// file
																									// instance
			FileInputStream fis = new FileInputStream(file); // obtaining bytes from the file
			// creating Workbook instance that refers to .xlsx file
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheetAt(0); // creating a Sheet object to retrieve object

			Iterator<Row> itr = sheet.iterator(); // iterating over excel file

			while (itr.hasNext()) {
				Row row = itr.next();
				Iterator<Cell> cellIterator = row.cellIterator(); // iterating over each column
				boolean flag = false;
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {
					case STRING: // field that represents string cell type

						if (cell.getStringCellValue().equalsIgnoreCase(string)) {
							flag = true;
						} else {
							continue;
						}

						break;
					case NUMERIC: // field that represents number cell type
						if (flag) {
							System.out.println(cell.getNumericCellValue());
							flag = false;
						} else {

						}

						break;
					default:
					}
				}
				System.out.println("");
			}
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

}
