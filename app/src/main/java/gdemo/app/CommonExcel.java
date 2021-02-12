package gdemo.app;

import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

public class CommonExcel {

	private DataFormatter formatter;
	private FormulaEvaluator evaluator;
	
	
	CommonExcel(Workbook workbook){
		this.evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		this.formatter = new DataFormatter(true);
	}
	
	/**
	 * Called to convert a row of cells into a line of data that can later be output
	 * to the CSV file.
	 *
	 * @param row An instance of either the HSSFRow or XSSFRow classes that
	 *            encapsulates information about a row of cells recovered from an
	 *            Excel workbook.
	 */
	public ArrayList<String> rowToCSV(Row row) {
		Cell cell;
		int lastCellNum;
		ArrayList<String> csvLine = new ArrayList<String>();
		
		
		

		// Check to ensure that a row was recovered from the sheet as it is
		// possible that one or more rows between other populated rows could be
		// missing - blank. If the row does contain cells then...
		if (row != null) {

			// Get the index for the right most cell on the row and then
			// step along the row from left to right recovering the contents
			// of each cell, converting that into a formatted String and
			// then storing the String into the csvLine ArrayList.
			lastCellNum = row.getLastCellNum();
			for (int i = 0; i <= lastCellNum; i++) {
				cell = row.getCell(i);
				if (cell == null) {
					csvLine.add("");
				} else {
					if (cell.getCellType() != CellType.FORMULA) {
						csvLine.add(this.formatter.formatCellValue(cell));
					} else {
						try {
							csvLine.add(this.formatter.formatCellValue(cell, this.evaluator));
						} catch (Exception e) {
							csvLine.add(e.getLocalizedMessage());
						}
					}
				}
			}
//			// Make a note of the index number of the right most cell. This value
//			// will later be used to ensure that the matrix of data in the CSV file
//			// is square.
//			if (lastCellNum > this.maxRowWidth) {
//				this.maxRowWidth = lastCellNum;
//			}
		}
		return csvLine;
	}
}
