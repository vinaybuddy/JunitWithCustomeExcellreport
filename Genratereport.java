package com.nt.test;


import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.junit.runner.JUnitCore;
import org.junit.runner.Result;
import org.junit.runner.notification.Failure;



public class Genratereport {

	public static void main(String[] args) throws IOException {
		System.out.println("Genratereport.main()");
		Result result = null;
		HSSFWorkbook workbook = null;
		HSSFSheet sheet1 = null;
		HSSFSheet sheet2 = null;
		Map<String, Object[]> data = null;
		Row row = null;
		Row row2=null;
		FileOutputStream stream = null;
		
		Cell cell=null;
		java.util.List<Failure> failureList=null;
		
		result = JUnitCore.runClasses(CalculatorTestSuit.class);
		// create Empty work wook
		workbook = new HSSFWorkbook();
		// create Empty sheet

		sheet1 = workbook.createSheet("TestCase");
		
	   
       /*  // Auto size the column widths
            for(int columnIndex = 0; columnIndex < 10; columnIndex++) {
                 sheet.autoSizeColumn(columnIndex);
            }*/
		// create row data
		CellStyle cellStyle = workbook.createCellStyle();
		CreationHelper createHelper = workbook.getCreationHelper();
		cellStyle.setDataFormat(
		    createHelper.createDataFormat().getFormat("d/m/yy h.mm;@"));
		

		for (int i = 0; i <2; i++) {
			int j = 0;
			if (i == 0) {
				row = sheet1.createRow(i);
				
				row.createCell(j++).setCellValue("Test Class");
				row.createCell(j++).setCellValue("Run Count");
				row.createCell(j++).setCellValue("Failure Count");
				row.createCell(j++).setCellValue("Ignore count");
				row.createCell(j++).setCellValue("Run Time");
				row.createCell(j++).setCellValue("genrate date");
				j = 0;
			} // if
			else {
				row = sheet1.createRow(i);
				
				row.createCell(j++).setCellValue("CalculatorTestSuit");
				row.createCell(j++).setCellValue(result.getRunCount());
				row.createCell(j++).setCellValue(result.getFailureCount());
				row.createCell(j++).setCellValue(result.getIgnoreCount());
				row.createCell(j++).setCellValue(result.getRunTime());
				cell = row.createCell(j++);
				cell.setCellValue(new Date());
				cell.setCellStyle(cellStyle);
				j = 0;

			} // else

		} // for
		
		
		if(result.getFailures().size()>0) {
			
			sheet2 = workbook.createSheet("FailureResone");
			// create row data
					row = sheet2.createRow(0);
					row.createCell(0).setCellValue("Failure Resons");
					failureList=result.getFailures();
					System.out.println(failureList.size());
					int i=1;
					for(Failure failure:failureList) {
						 row2=sheet2.createRow(i++);
						row2.createCell(0).setCellValue(failure.getMessage());
						
						
					}//for
				
							
		}//else

		// locate output stream
		try {
			stream = new FileOutputStream("D:\\excelreport\\UnitTest.xls");
			workbook.write(stream);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		stream.close();

	}// main(-)

}// class
