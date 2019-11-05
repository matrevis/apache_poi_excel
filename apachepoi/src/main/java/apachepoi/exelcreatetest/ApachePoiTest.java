package apachepoi.exelcreatetest;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class ApachePoiTest {
	/**
	 * Logger for this class
	 */
	private static final Logger logger = LoggerFactory.getLogger(ApachePoiTest.class);
	
	private static final String FILE_NAME = "MyExcelTest.xlsx";

	public static void main(String[] args) {
		logger.info("Start..");
		// write code here!

		XSSFWorkbook workbook = new XSSFWorkbook();
		
		// default sheet
		/*
		XSSFSheet sheet = workbook.createSheet("Datatypes in Java");
		Object[][] datatypes = { { "Datatype", "Type", "Size(in bytes)" }, { "int", "Primitive", 2 },
				{ "float", "Primitive", 4 }, { "double", "Primitive", 8 }, { "char", "Primitive", 1 },
				{ "String", "Non-Primitive", "No fixed size" } };

		int rowNum = 0;
		logger.info("Creating excel");

		for (Object[] datatype : datatypes) {
			Row row = sheet.createRow(rowNum++);
			int colNum = 0;
			for (Object field : datatype) {
				Cell cell = row.createCell(colNum++);
				if (field instanceof String) {
					cell.setCellValue((String) field);
				} else if (field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}
			}
		}
		*/
		// end default sheet
		
		XSSFSheet KPIAvailability = workbook.createSheet("KPI availability");
		XSSFSheet KPIPerformance = workbook.createSheet("KPI perfomance");
		
		// KPI 
		
		Object[] header = {
				"HeaderColonna1", "HeaderColonna2", "HeaderColonna3"
		};
		
		Object[] subHeader = {
				"SubHeaderColonna1_1", "SubHeaderColonna1_2",
				"SubHeaderColonna2_1", "SubHeaderColonna2_2", 
				"SubHeaderColonna3_1", "SubHeaderColonna3_2"
		};
		
		// creo Header di KPI Availability
		Row headerRow = KPIAvailability.createRow(1);
		int colNum = 2;
		for (Object child : header) {
			Cell cell = headerRow.createCell(colNum);
			colNum += 2;
			if(child instanceof String) {
				cell.setCellValue( (String) child);
			}
			else {
				//cell.setBlank();
				cell.setCellValue("N\\A");
				logger.warn("Warning: empty cell at colNum: {}, oggetto: {} non convertibile in String..", colNum, child);
			}
		}
		KPIAvailability.addMergedRegion(new CellRangeAddress(1,1,2,3));
		KPIAvailability.addMergedRegion(new CellRangeAddress(1,1,4,5));
		KPIAvailability.addMergedRegion(new CellRangeAddress(1,1,6,7));
		
		// creo SubHeader di KPI Availability
		Row subHeaderRow = KPIAvailability.createRow(2);
		colNum = 2;
		for (Object child : subHeader) {
			Cell cell = subHeaderRow.createCell(colNum);
			colNum += 1;
			if(child instanceof String) {
				cell.setCellValue( (String) child);
			}
			else {
				//cell.setBlank();
				cell.setCellValue("N\\A");
				logger.warn("Warning: empty cell at colNum: {}, oggetto: {} non convertibile in String..", colNum, child);
			}
		}
		
		/*
		Object[][] dataFromJPA = {
				{"misurazione1", 1, 0},
				{"misurazione2", 1, 0}
		};
		
		colNum = 2;
		for (Object child : dataFromJPA) {
			Cell cell = headerRow.createCell(colNum++);
			if(child instanceof String) {
				cell.setCellValue( (String) child);
			}
			else if(child instanceof Integer) {
				cell.setCellValue( (Integer) child);
			}
			else {
				cell.setBlank();
				logger.warn("Warning: empty cell at colNum: {}, oggetto: {} non convertibile in String o Integer..", colNum++, child);
			}
		}
		*/
		
		// autosize interested columns
		logger.info("ApachePOI: resize columns..");
		for(int i=1; i<colNum; i++) {
			KPIAvailability.autoSizeColumn(i);
		}
		try {
			FileOutputStream outputStream = new FileOutputStream(FILE_NAME);
			workbook.write(outputStream);
			workbook.close();
		} catch (FileNotFoundException e) {
			logger.error("Errore: file non trovato o inaccessibile..", e);
//			e.printStackTrace();
		} catch (IOException e) {
			logger.error("Errore: ioExp..", e);
//			e.printStackTrace();
		}

		// end
		logger.info("End..");
	}

}
