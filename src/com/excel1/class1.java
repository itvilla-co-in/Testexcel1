package com.excel1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.util.AreaReference;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class class1 {

	public static void main(String[] args) throws FileNotFoundException, IOException 
	{
		//Creating the workbook and the spreadsheet:
		// Tab Name
		Workbook wb = new XSSFWorkbook();
		CreationHelper ch = wb.getCreationHelper();
		String safeName = WorkbookUtil.createSafeSheetName("Data");
		XSSFSheet sheet = (XSSFSheet)wb.createSheet(safeName);
		
		//Reading the CSV file and writing the data. The first row is treated as column headers.
		int rowNum = 0;
		List<String> colNames = null;
		try (InputStream in = new FileInputStream("D:/nk0072025/TECHM/temp.txt");) 
		{
		    CSV csv = new CSV(true, ',', in);
		    if ( csv.hasNext() ) 
		    {
		    	// colname has all fileds of row 1 and stored each field into array of string
		    	colNames = new ArrayList<String>(csv.next());
		    	//System.out.println("Lets see what colnames has " + colNames);
		    	Row row = sheet.createRow((short)0);
		    	for (int i = 0 ; i < colNames.size() ; i++) 
		    	{
		    		String name = colNames.get(i);
		    		row.createCell(i).setCellValue(name);
		    		//System.out.println("Lets see what this has " + name);
		    		//writes each filed value to cell
		    	}
		    }

		    while (csv.hasNext()) {
		    List<String> fields = csv.next();
		    rowNum++;
		    Row row = sheet.createRow((short)rowNum);
		    
		    for (int i = 0 ; i < fields.size() ; i++) {
		        /* Attempt to set as double. If that fails, set as
		         * text. */
		        try {
		        double value = Double.parseDouble(fields.get(i));
		        row.createCell(i).setCellValue(value);
		        } catch(NumberFormatException ex) {
		        String value = fields.get(i);
		        row.createCell(i).setCellValue(value);
		        }
		    }
		    }
		}
		
		//We want the whole of data included in the pivot table. 
		//So we use the following ranges to create cell references
		
		int firstRow = sheet.getFirstRowNum();
		int lastRow = sheet.getLastRowNum();
		int firstCol = sheet.getRow(0).getFirstCellNum();
		int lastCol = sheet.getRow(0).getLastCellNum();
		
		//The cell references specify the top left and the bottom right of the table data.
		
		CellReference topLeft = new CellReference(firstRow, firstCol);
		CellReference botRight = new CellReference(lastRow, lastCol - 1);
		
		//And the area reference which marks out the data table.
		AreaReference aref = new AreaReference(topLeft, botRight);
		
		//We insert the pivot table at this location
		//a couple of rows offset from the top of the sheet and to the right of the table.
		CellReference pos = new CellReference(firstRow + 4, lastCol + 1);
		
		//Finally we create the pivot table from the area reference and the cell position.
		XSSFPivotTable pivotTable = sheet.createPivotTable(aref, pos);

		//In this example, we are grouping data by yearID and teamID. These are columns 0 and 1 respectively. 
		//And for the summary, we choose to sum the salaries over these columns.
		pivotTable.addRowLabel(0);
		pivotTable.addRowLabel(1);
		pivotTable.addColumnLabel(DataConsolidateFunction.SUM,
		              4, "Sum of " + colNames.get(4));
		
		//save
		FileOutputStream fileOut = new FileOutputStream("D:/nk0072025/TECHM/temp.xls");
		wb.write(fileOut);
		fileOut.close();
	}

}
