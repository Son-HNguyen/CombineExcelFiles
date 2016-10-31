import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Test {

	// ------------------------------------------------------------------------------------------------------------------------
	// Number of months/sheets that should be processed in the workbook
	public static final int NR_OF_MONTHS = 12;

	// Decides in which order sheets should be processed:
	// [true: from left to right]; [false: from right to left]
	public static final boolean LATEST_MONTH_FIRST = true;

	// Relative path of input XLSX file
	public static final String INPUT_XLSX_PATH = "input/Docks_Consumption.xlsx";

	// Relative path of output XLSX file
	public static final String OUTPUT_XLSX_PATH = "output/Combined_Docks_Consumption.xlsx";
	// ------------------------------------------------------------------------------------------------------------------------

	public static void main(String[] args) {

		try {
			// CREATE NEW WORKBOOK TO STORE COMBINED DATA
			XSSFWorkbook outputWorkbook = new XSSFWorkbook();
			XSSFSheet outputSheet = outputWorkbook.createSheet("Combined");

			// EXTRACT SHEETS FROM AVAILABLE WORKBOOK
			FileInputStream file = new FileInputStream(new File(INPUT_XLSX_PATH));
			XSSFWorkbook inputWorkbook = new XSSFWorkbook(file);

			// Initialize auxiliary boundary indices for the upcoming for loop
			int inputSheetIndBegin, inputSheetIndEnd, inputSheetIndStep;
			if (LATEST_MONTH_FIRST) {
				inputSheetIndBegin = 0;
				inputSheetIndEnd = inputWorkbook.getNumberOfSheets();
				inputSheetIndStep = 1;
			} else {
				inputSheetIndBegin = inputWorkbook.getNumberOfSheets() - 1;
				inputSheetIndEnd = -1;
				inputSheetIndStep = -1;
			}

			int outputRowInd = 0;

			// Loop through rows in every sheet
			for (int inputRowInd = 0; inputRowInd < inputWorkbook.getSheetAt(0).getLastRowNum(); inputRowInd++) {

				// Loop through sheets in workbook
				for (int inputSheetInd = inputSheetIndBegin; inputSheetInd != inputSheetIndEnd; inputSheetInd += inputSheetIndStep) {

					XSSFSheet inputSheet = inputWorkbook.getSheetAt(inputSheetInd);

					// Ignore first row of sheet starting from index 1
					if (inputSheetInd != 0 && inputRowInd == 0) {
						continue;
					}

					Row inputRow = inputSheet.getRow(inputRowInd);
					Row outputRow = outputSheet.createRow(outputRowInd++);

					// Loop through cells in a row
					int inputCellInd = 0;
					for (; inputCellInd < inputRow.getLastCellNum(); inputCellInd++) {
						Cell outputCell = outputRow.createCell(inputCellInd);
						Cell inputCell = inputRow.getCell(inputCellInd);
						// A cell can have String or Numeric value
						try {
							outputCell.setCellValue(inputCell.getNumericCellValue());
						} catch (IllegalStateException e) {
							outputCell.setCellValue(inputCell.getStringCellValue());
						}
					}

					// Create extra columns that stores the current sheet's month and year ["mm_yyyy"]
					String[] monthAndYear = inputSheet.getSheetName().split("_");
					Cell outputMonthCell = outputRow.createCell(inputCellInd++);
					Cell outputYearCell = outputRow.createCell(inputCellInd++);
					if(inputRowInd == 0){
						outputMonthCell.setCellValue("Month");						
						outputYearCell.setCellValue("Year");
					}else{
						outputMonthCell.setCellValue(Integer.parseInt(monthAndYear[0]));
						outputYearCell.setCellValue(Integer.parseInt(monthAndYear[1]));
					}					
				}
			}

			// WRITE OUTPUT WORKBOOK
			FileOutputStream out = new FileOutputStream(new File(OUTPUT_XLSX_PATH));
			outputWorkbook.write(out);
			outputWorkbook.close();
			inputWorkbook.close();
			out.close();
			System.out.println("Excel written successfully..");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
