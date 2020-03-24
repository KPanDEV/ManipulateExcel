package manipulate.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ProcessExcel {
	public String policyNumber;
	public String date;
    Map<String, String> stringData = new HashMap<String, String>();
    Map<String, String> stringCell2Data = new HashMap<String, String>();
	public void processExcelFile()
	{
	try {
		FileInputStream file = new FileInputStream(
				new File("SourceFOlder-Excel-File-Location"));

		// Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		// Get first/desired sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);
        
		// Iterate through each rows one by one
		Iterator<Row> rowIterator = sheet.iterator();

		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();
			// For each row, iterate through all the columns
			Iterator<Cell> cellIterator = row.cellIterator();

			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				String[] splited = cell.getStringCellValue().split("\\s+");
				// Check the cell type and format accordingly
				System.out.println("Printing CellType Value ="+cell.getCellType());
				switch (cell.getCellType()) {
				case Cell.CELL_TYPE_NUMERIC:
					System.out.print(cell.getNumericCellValue() + "t");
					break;
				case Cell.CELL_TYPE_STRING:
					System.out.print("CellValue="+cell.getStringCellValue() + "\t");
				    manipulateExcelData(splited);//call this method to manipulate Data
				    if(cell.getStringCellValue().equals("L") || cell.getStringCellValue().equals("H"))
				    {
				    	stringData.put(this.policyNumber, this.date);
				    	stringCell2Data.put(this.policyNumber,cell.getStringCellValue());
				    }
					break;
				}
			}
			System.out.println("");
		}
		file.close();
		enterDataInCell(stringData,stringCell2Data);
	} catch (Exception e) {
		e.printStackTrace();
	}
	}

	public void manipulateExcelData(String [] splited)
	{
		if(splited.length > 1)
		{
		this.date=splited[0];
		splitPolicyNumber(splited[4]);
		}
	}
	
	public void splitPolicyNumber(String splited)
	{
		int index = splited.lastIndexOf("_");
		String splittedNow = splited.substring(17, index);
		this.policyNumber = splittedNow;
	}
	
	public void enterDataInCell(Map stringData, Map stringData2) throws IOException
	{
		FileOutputStream file = new FileOutputStream(
				new File("SourceFile-Excel-File-Location"));

		XSSFWorkbook workbook = new XSSFWorkbook();
		
		XSSFSheet sheet =  workbook.createSheet("CompanyData");

		System.out.println("StringData2="+stringData2);
		int rowCount = sheet.getLastRowNum();
		Iterator iterate  = stringData.entrySet().iterator();
		Iterator iteratePolicyType = stringData2.entrySet().iterator();
		while(iterate.hasNext() && iteratePolicyType.hasNext())
		{
			Map.Entry mapElementCell1 = (Map.Entry)iterate.next();//For column1
			Map.Entry mapElementCell2 = (Map.Entry)iteratePolicyType.next();//For Column2
			Row row = sheet.createRow(++rowCount);
			 
            int columnCount = 0;
             
            Cell cell ;//= row.createCell(columnCount);
            
            cell = row.createCell(++columnCount);
            cell.setCellValue(mapElementCell1.getKey().toString());//this is for policyNumber
            
            cell = row.createCell(++columnCount);
    		cell.setCellValue(mapElementCell1.getValue().toString());//For date
    		
    		cell = row.createCell(++columnCount);
    		cell.setCellValue(mapElementCell2.getValue().toString());//For type of policy
            
        }
		workbook.write(file);
        workbook.close();
		file.close();
	}
}
