package gdemo.app;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import gdemo.app.ToCSV.ExcelFilenameFilter;

public class CompareHeader {
	
	private Workbook workbook;
	private ArrayList<ArrayList<String>> csvData;
	
	private FormulaEvaluator evaluator;
	private String separator;
	private int formattingConvention;
	private DataFormatter formatter;
	
	public void compareHeaderToCSV(String strSource, String strDestination)
			throws FileNotFoundException, IOException, IllegalArgumentException {

		// Simply chain the call to the overloaded convertExcelToCSV(String,
		// String, String, int) method, pass the default separator and ensure
		// that certain embedded characters are escaped in accordance with
		// Excel's formatting conventions
		this.compareHeaderToCSV(strSource, strDestination, ToCSV.DEFAULT_SEPARATOR, ToCSV.EXCEL_STYLE_ESCAPING);
	}
	
	/**
	 * Open an Excel workbook ready for conversion.
	 *
	 * @param file An instance of the File class that encapsulates a handle to a
	 *             valid Excel workbook. Note that the workbook can be in either
	 *             binary (.xls) or SpreadsheetML (.xlsx) format.
	 * @throws java.io.FileNotFoundException Thrown if the file cannot be located.
	 * @throws java.io.IOException           Thrown if a problem occurs in the file
	 *                                       system.
	 */
	private void openWorkbook(File file) throws FileNotFoundException, IOException {
		System.out.println("Opening workbook [" + file.getName() + "]");
		try (FileInputStream fis = new FileInputStream(file)) {

			// Open the workbook and then create the FormulaEvaluator and
			// DataFormatter instances that will be needed to, respectively,
			// force evaluation of forumlae found in cells and create a
			// formatted String encapsulating the cells contents.
			this.workbook = WorkbookFactory.create(fis);
			this.evaluator = this.workbook.getCreationHelper().createFormulaEvaluator();
			this.formatter = new DataFormatter(true);
		}
	}

	public void compareHeaderToCSV(String strSource, String strDestination, String separator, int formattingConvention)
			throws FileNotFoundException, IOException, IllegalArgumentException {
		// Check that the source file/folder exists.
		File source = new File(strSource);
		if (!source.exists()) {
			throw new IllegalArgumentException("The source for the Excel " + "file(s) cannot be found at " + source);
		}

		// Ensure thaat the folder the user has chosen to save the CSV files
		// away into firstly exists and secondly is a folder rather than, for
		// instance, a data file.
		File destination = new File(strDestination);
		if (!destination.exists()) {
			throw new IllegalArgumentException(
					"The destination directory " + destination + " for the " + "converted CSV file(s) does not exist.");
		}
		if (!destination.isDirectory()) {
			throw new IllegalArgumentException(
					"The destination " + destination + " for the CSV " + "file(s) is not a directory/folder.");
		}

		// Ensure the value passed to the formattingConvention parameter is
		// within range.
		if (formattingConvention != ToCSV.EXCEL_STYLE_ESCAPING && formattingConvention != ToCSV.UNIX_STYLE_ESCAPING) {
			throw new IllegalArgumentException("The value passed to the "
					+ "formattingConvention parameter is out of range: " + formattingConvention + ", expecting one of "
					+ ToCSV.EXCEL_STYLE_ESCAPING + " or " + ToCSV.UNIX_STYLE_ESCAPING);
		}

		// Copy the spearator character and formatting convention into local
		// variables for use in other methods.
		this.separator = separator;
		this.formattingConvention = formattingConvention;

		// Check to see if the sourceFolder variable holds a reference to
		// a file or a folder full of files.
		final File[] filesList;
		if (source.isDirectory()) {
			// Get a list of all of the Excel spreadsheet files (workbooks) in
			// the source folder/directory
			filesList = source.listFiles(new ExcelFilenameFilter());
		} else {
			// Assume that it must be a file handle - although there are other
			// options the code should perhaps check - and store the reference
			// into the filesList variable.
			filesList = new File[] { source };
		}

		// Step through each of the files in the source folder and for each
		// open the workbook, convert it's contents to CSV format and then
		// save the resulting file away into the folder specified by the
		// contents of the destination variable. Note that the name of the
		// csv file will be created by taking the name of the Excel file,
		// removing the extension and replacing it with .csv. Note that there
		// is one drawback with this approach; if the folder holding the files
		// contains two workbooks whose names match but one is a binary file
		// (.xls) and the other a SpreadsheetML file (.xlsx), then the names
		// for both CSV files will be identical and one CSV file will,
		// therefore, over-write the other.
		if (filesList != null) {
			for (File excelFile : filesList) {
				// Open the workbook
				this.openWorkbook(excelFile);

				// Convert it's contents into a CSV file
				this.compareHeaderToCSV();

				// Build the name of the csv folder from that of the Excel workbook.
				// Simply replace the .xls or .xlsx file extension with .csv
				String destinationFilename = excelFile.getName();
				destinationFilename = destinationFilename.substring(0, destinationFilename.lastIndexOf('.'))
						+ ToCSV.CSV_FILE_EXTENSION;

				// Save the CSV file away using the newly constricted file name
				// and to the specified directory.
				this.saveCSVFile(new File(destination, destinationFilename));
			}
		}
	}
	
	/**
	 * Called to actually save the data recovered from the Excel workbook as a CSV
	 * file.
	 *
	 * @param file An instance of the File class that encapsulates a handle
	 *             referring to the CSV file.
	 * @throws java.io.FileNotFoundException Thrown if the file cannot be found.
	 * @throws java.io.IOException           Thrown to indicate and error occurred
	 *                                       in the underylying file system.
	 */
	private void saveCSVFile(File file) throws FileNotFoundException, IOException {
		StringBuilder buffer;
		// Open a writer onto the CSV file.
		try (BufferedWriter bw = Files.newBufferedWriter(file.toPath(), StandardCharsets.ISO_8859_1)) {

			System.out.println("Saving the CSV file [" + file.getName() + "]");

			// Step through the elements of the ArrayList that was used to hold
			// all of the data recovered from the Excel workbooks' sheets, rows
			// and cells.
			for (ArrayList<String> oneLine:this.csvData) {
				buffer = new StringBuilder();
				for (String oneCell: oneLine) {
					if (oneCell != null) {
						buffer.append(this.escapeEmbeddedCharacters(oneCell));
					}
					buffer.append(this.separator);
				}
				// Once the line is built, write it away to the CSV file.
				bw.write(buffer.toString().trim());
				bw.newLine();
			}
		}
	}
	
	private String escapeEmbeddedCharacters(String field) {
		StringBuilder buffer;

		// If the fields contents should be formatted to confrom with Excel's
		// convention....
		if (this.formattingConvention == ToCSV.EXCEL_STYLE_ESCAPING) {

			// Firstly, check if there are any speech marks (") in the field;
			// each occurrence must be escaped with another set of spech marks
			// and then the entire field should be enclosed within another
			// set of speech marks. Thus, "Yes" he said would become
			// """Yes"" he said"
			if (field.contains("\"")) {
				buffer = new StringBuilder(field.replaceAll("\"", "\\\"\\\""));
				buffer.insert(0, "\"");
				buffer.append("\"");
			} else {
				// If the field contains either embedded separator or EOL
				// characters, then escape the whole field by surrounding it
				// with speech marks.
				buffer = new StringBuilder(field);
				if ((buffer.indexOf(this.separator)) > -1 || (buffer.indexOf("\n")) > -1) {
					buffer.insert(0, "\"");
					buffer.append("\"");
				}
			}
			return (buffer.toString().trim());
		}
		// The only other formatting convention this class obeys is the UNIX one
		// where any occurrence of the field separator or EOL character will
		// be escaped by preceding it with a backslash.
		else {
			if (field.contains(this.separator)) {
				field = field.replaceAll(this.separator, ("\\\\" + this.separator));
			}
			if (field.contains("\n")) {
				field = field.replaceAll("\n", "\\\\\n");
			}
			return (field);
		}
	}
	
	/**
	 * Called to convert the contents of the currently opened workbook into a CSV
	 * file.
	 */
	private void compareHeaderToCSV() {
		Sheet sheet;
		Row row;
		int lastRowNum;
		this.csvData = new ArrayList<>();
		CommonExcel commonExcel = new CommonExcel(this.workbook);

		System.out.println("Get second row from first worksheet to CSV format.");

		sheet = this.workbook.getSheetAt(0);
		if (sheet.getPhysicalNumberOfRows() > 0) {
			lastRowNum = sheet.getLastRowNum();
			for (int j = 0; j <= lastRowNum; j++) {
				row = sheet.getRow(j);
				this.csvData.add(commonExcel.rowToCSV(row));
			}
		}
	}



}
