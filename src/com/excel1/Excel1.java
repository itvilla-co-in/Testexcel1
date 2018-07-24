package com.excel1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class Excel1 {
	
	private static Object getCellValue(Cell cell) {
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_STRING:
		return cell.getStringCellValue();
	
		case Cell.CELL_TYPE_BOOLEAN:
		return cell.getBooleanCellValue();
		case Cell.CELL_TYPE_NUMERIC:
		return cell.getNumericCellValue();
		}
	
		return null;
		}
	
	public static void main(String[] args) throws IOException {
	  
		  String excelFilePath = "D:/nk0072025/TECHM/Book1.xlsx";
	        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
	         
	        Workbook workbook = new XSSFWorkbook(inputStream);
	        Sheet firstSheet = workbook.getSheetAt(0);
	        Iterator<Row> iterator = firstSheet.iterator();
	         
	        while (iterator.hasNext()) {
	            Row nextRow = iterator.next();
	            Iterator<Cell> cellIterator = nextRow.cellIterator();
	             
	            while (cellIterator.hasNext()) {
	                Cell cell = cellIterator.next();
	                int columnIndex = cell.getColumnIndex();
	               /* working swithc   
	                switch (cell.getCellType()) {
	                    case Cell.CELL_TYPE_STRING:
	                        System.out.print(cell.getStringCellValue());
	                        break;
	                    case Cell.CELL_TYPE_BOOLEAN:
	                        System.out.print(cell.getBooleanCellValue());
	                        break;
	                    case Cell.CELL_TYPE_NUMERIC:
	                        System.out.print(cell.getNumericCellValue());
	                        break;
	                }
	                */
	                switch (columnIndex) {
	                case 1:
	    	            System.out.println(getCellValue(cell));
                        break;
	                case 2:
	                	System.out.println(getCellValue(cell));
	                	break;
	                case 3:
	                	System.out.println(getCellValue(cell));
	                    break;
	                case 4:
	                	System.out.println(getCellValue(cell));
	                    break;
	                case 5:
	                	System.out.println(getCellValue(cell));
	                    break;
	                case 6:
	                	System.out.println(getCellValue(cell));
	                    break;
	                case 7:
	                	System.out.println(getCellValue(cell));
	                    break;
	                case 8:
	                	System.out.println(getCellValue(cell));
	                    break;
	                }
	                System.out.print(" - ");
	            }
	            System.out.println();
	        }
	         
	        workbook.close();
	        inputStream.close();
	}

}
