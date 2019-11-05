package apachepoi.exelcreatetest;

import java.awt.List;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

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

	private static final String FILE_NAME = "target/MyExcelTest.xlsx";

	private static Integer colMaxNum = 2;

	private static void doHeader(XSSFSheet sheet, Object[] headerKPI) {
		Row headerRow = sheet.createRow(1);
		int colNum = 2;
		for (Object child : headerKPI) {
			Cell cell = headerRow.createCell(colNum);
			sheet.addMergedRegion(new CellRangeAddress(1, 1, colNum, colNum + 1));
			colNum += 2;
			colMaxNum = colNum > colMaxNum ? colNum : colMaxNum;
			if (child instanceof String) {
				cell.setCellValue((String) child);
			} else {
				cell.setCellValue("N\\A");
				logger.warn("Warning: empty cell at colNum: {}, oggetto: {} non convertibile in String..", colNum,
						child);
			}
		}
	}

	private static void doSubHeader(XSSFSheet sheet, Object[] subHeaderKPI) {
		Row subHeaderRow = sheet.createRow(2);
		int colNum = 2;
		for (Object child : subHeaderKPI) {
			Cell cell = subHeaderRow.createCell(colNum);
			colNum += 1;
			if (child instanceof String) {
				cell.setCellValue((String) child);
			} else {
				// cell.setBlank();
				cell.setCellValue("N\\A");
				logger.warn("Warning: empty cell at colNum: {}, oggetto: {} non convertibile in String..", colNum,
						child);
			}
		}
	}
	
	private static void doBody(XSSFSheet sheet, ArrayList<Object[]> data) {
		int colNum = 1;
		int rowNum = 3;
		for (Object[] rowItem: data) {
			colNum = 1;
			Row rowData = sheet.createRow(rowNum);
			rowNum += 1;
			for (Object child : rowItem) {
				Cell cell = rowData.createCell(colNum++);
				if (child instanceof String) {
					cell.setCellValue((String) child);
				} else if (child instanceof Integer) {
					cell.setCellValue((Integer) child);
				} else {
					cell.setBlank();
					logger.warn("Warning: empty cell at colNum: {}, oggetto: {} non convertibile in String o Integer..",
							colNum++, child);
				}
			}
		}
	}

	private static void doResize(XSSFSheet sheet) {
		logger.info("ApachePOI resize columns..");
		for (int i = 1; i < colMaxNum; i++) {
			sheet.autoSizeColumn(i);
		}
	}

	public static void main(String[] args) {
		logger.info("Start..");
		// write code here!

		XSSFWorkbook workbook = new XSSFWorkbook();

		XSSFSheet KPIAvailability = workbook.createSheet("KPI availability");
		XSSFSheet KPIPerformance = workbook.createSheet("KPI perfomance");

		// KPI

		Object[] headerKPIA = { "HeaderColonna1", "HeaderColonna2" };

		Object[] subHeaderKPIA = { "SubHeaderColonna1_1", "SubHeaderColonna1_2", "SubHeaderColonna2_1",
				"SubHeaderColonna2_2" };

		// creo Header di KPI Availability
		doHeader(KPIAvailability, headerKPIA);
		doSubHeader(KPIAvailability, subHeaderKPIA);
		doResize(KPIAvailability);

		// riempimento dati KPIAvailability
		Object[][] dataFromJPA = { { "31-Nov", 100, 0, 99, 1 }, { "01-Dec", 100, 0, 95, 5 } };
		Object[] datoSingolo = { "31-Nov", 100, 0, 99, 1 };
		ArrayList<Object[]> listOfJPAData = new ArrayList<Object[]>();
		for(int i=0; i<100; i++) {
			listOfJPAData.add(datoSingolo);
		}
		
		//doBody(KPIAvailability, dataFromJPA);
		doBody(KPIAvailability, listOfJPAData);

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
