package asst.hssf;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Class to hold information about the work book being read.  Workbook
 * and Sheet are interfaces which can hold implementations for both.
 * The plan is to be able to read either type of spread sheet and
 * process them in the same way using the same libraries.  The goal is
 * to be able to handle both xls and xlsx documents.
 * @author Material Gain
 * @2015 01
 */

public class WorkbookManager {
  protected CreationHelper createHelper;
  /** Column map in the current sheet*/
  public Map<String, Integer> columnMap;
  /** The current row being processed*/
  public Row row;
  /** The current work sheet in the work book. */
  public Sheet sheet;
  /** Name of the current sheet */
  public String sheetName;
  /** Error accumulator for all work sheet operations. */
  public StringBuilder sb = new StringBuilder();
  /** This stores a workbook that came either from .xls or .xlsx, but it
   * remembers which it was and writes itself accordingly.*/
  public Workbook wb;
  /** Store the name of the file which is associated with the work book*/
  public String fileName;
  /** True means to update the database if all rows are good */
  public boolean doIt;
  /** True means to generate extra sysout. */
  public boolean verbose;
  /** True means that the work sheet is updating existing rows, false means
   * that the work sheet is creating new rows from spread sheet rows*/
  public boolean updating;

  /** The 0-based index of the current row being processed*/
  public int currentRowNumber;
  protected int firstRowNumber;
  protected int lastRowNumber;
  /** Number of actual rows in the spread sheet*/
  protected int physicalRows;
  protected int sheetIndex;

  /**
   * Default constructor
   */
  public WorkbookManager() { }

  /**
   * Constructor which grabs the first sheet in a work book
   * @param wb the workbook of either kind
   */
  public WorkbookManager(Workbook wb) {
    this.wb = wb;
    createHelper = wb.getCreationHelper();
    sheet = wb.getSheetAt(0);
    sheetCharacteristics();
  }

  /**
   * Constructor which selects or creates a named sheet if asked to do
   * so
   * @param wb the workbook of either kind
   * @param sheetName the name of the desired sheet
   * @param create tells whether to create the sheet if it does not exist
   */
  public WorkbookManager(Workbook wb, String sheetName, boolean create) {
    this.wb = wb;
    createHelper = wb.getCreationHelper();
    if (( (sheet = wb.getSheet(sheetName)) == null) && create) {
      sheet = wb.createSheet(sheetName);
    }
    sheetCharacteristics();
  }

  /**
   * Constructor which selects or creates a numbered sheet and creates
   * one new sheet if asked to do so
   * @param wb the workbook of either kind
   * @param sheetNum the 0-based number of the desired sheet
   * @param create tells whether to create the sheet if it does not exist
   */
  public WorkbookManager(Workbook wb, int sheetNum, boolean create) {
    this.wb = wb;
    createHelper = wb.getCreationHelper();
    if (( (sheet = wb.getSheetAt(sheetNum)) == null) && create) {
      sheet = wb.createSheet();
    }
    sheetCharacteristics();
  }

  /**
   * Selects a specified sheet as the current sheet.  Leaves sheet null
   * if there is no sheet by that name.
   * @param desiredSheetName name of the desired sheet.
   * @return The desired sheet or null.
   */
  public Sheet pickSheet(String desiredSheetName) {
    if ( (sheet = wb.getSheet(desiredSheetName)) == null) { return null; }
    sheetCharacteristics();
    return sheet;
  }

  /**
   * Compute various interesting characteristics of the specified
   * sheet within the work book
   */
  public void sheetCharacteristics() {
    if (sheet == null) { return; }
    sheetName = sheet.getSheetName();
    lastRowNumber = sheet.getLastRowNum();
    firstRowNumber = sheet.getFirstRowNum();
    physicalRows = sheet.getPhysicalNumberOfRows();
    if (physicalRows != 0) {
      row = sheet.getRow(firstRowNumber);
      currentRowNumber = firstRowNumber;
    } else {
      currentRowNumber = -1;
      row = null;
    }
  }

  /**
   * Set the sheet margins to fairly narrow
   */
  public void setSheetMargins() {
    if (sheet == null) { return; }
    sheet.setMargin((short)0, .5d);
    sheet.setMargin((short)1, .5d);
    sheet.setMargin((short)2, .5d);
    sheet.setMargin((short)3, .5d);
  }

  /**
   * To be called after the sheet is filled in.  Set all the columns to their
   * preferred width based on content.
   */
  public void setSheetColumnWidths() {
    if (sheet == null) { return; }
    Row rowX = sheet.getRow(0);
    if (rowX == null) { return; }
    for (int i = 0; i<rowX.getLastCellNum(); i++) {
      sheet.autoSizeColumn(i);
    }
  }

  /**
   * The only row in the sheet is the column headings.  Set the column
   * widths to accommodate the column labels.
   */
  public void setColumnWidthsToLabels() {
    if ((sheet == null) || (row == null)) { return; }
    /* The last cell number is one more than the index of the last
     * existing cell*/
    for (int k = row.getFirstCellNum(); k<row.getLastCellNum(); k++) {
      sheet.autoSizeColumn(k);
    }
  }

  /**
   * Process all the rows based on having explored the sheet characteristics
   * @return the next non-comment row in the current sheet or null if
   * there are no more rows.
   */
  public Row nextRow() {
    if (currentRowNumber >= lastRowNumber) { return null; }
    Cell cell;
    do {
      currentRowNumber++;
      row = sheet.getRow(currentRowNumber);
      if (row != null) {
	cell = row.getCell(0);
	if (cell == null) { continue; }
	if ((cell.getCellType() == CellType.NUMERIC) ||
	    ((cell.getStringCellValue() != null) &&
		!cell.getStringCellValue().startsWith("#"))) { return row; }
      }
    } while (currentRowNumber <= lastRowNumber);
    return null;
  }

  /**
   * @return a new row which was added to the work sheet.
   */
  public Row addRowToSheet() {
    return row = sheet.createRow(sheet.getLastRowNum()+1);
  }

  /**
   * @param cols array of required column names.
   * @return list of missing columns or null if there are no missing
   * columns
   */
  public List<String> missingColumns(String[] cols) {
    List<String> missing = null;
    for (String col : cols) {
      if (!columnMap.containsKey(col)) {
	if (missing == null) {
	  missing = new ArrayList<String>();
	}
	missing.add(col);
      }
    }
    return missing;
  }

  /**
   * Make sure that the current work sheet has at least 2 rows and
   * that the first physical row has a list of column names which have
   * no duplicates and which contain all the required column names
   * which are listed in the input column name array
   * @param requiredColNames required columns
   * @return true if the work sheet is OK, false otherwise.
   */
  public boolean isSheetOK(String[] requiredColNames) {
    return isSheetOK(requiredColNames, 2);
  }

  /**
   * Make sure that the current work sheet has enough rows and
   * that the first physical row has a list of column names which have
   * no duplicates and which contain all the required column names
   * which are listed in the input column name array
   * @param requiredColNames required columns
   * @param minRows Minimum number of rows in the sheet
   * @return true if the work sheet is OK, false otherwise.
   */
  public boolean isSheetOK(String[] requiredColNames, int minRows) {

    if ((row == null) || (physicalRows < minRows)) {
      sb.append("Work sheet " + sheetName +
	  fileName + " must have " + minRows +
	  " or more rows\n");
      return false;
    }
    return true;
  }

  /**
   * Augment an error message string
   * @param why why the list of bad strings is being created
   * @param what list of bad strings
   */
  public void augmentComplaints(String why, List<String> what) {
    sb.append(why + ":");
    for (String w : what) {
      sb.append("\t" + w);
    }
    sb.append("<br>\n");
  }

  /**
   * Add a complaint to the accumulated string
   * @param whinge text of the complaint
   */
  public void whinge(String whinge) {
    whinge(whinge, true);
  }

  /**
   * Add a complaint to the accumulated string
   * @param whinge text of the complaint
   * @param wantRow add sheet name and row number to whinge
   */
  public void whinge(String whinge, boolean wantRow) {
    sb.append((wantRow ? "Sheet " + sheetName + " row " + (currentRowNumber + 1) + " " : "") +
	whinge + "<br>\n");
  }

  /**
   * Set the value of the named cell in the current row.  Does nothing if
   * that cell is null.  This is intended for overriding existing cell
   * values.
   * @param col name of the column
   * @param val value to set in the cell
   */
  public void setColumnTo(String col, String val) {
    int which = columnMap.get(col);
    Cell cell = row.getCell(which);
    if (cell != null) {
      cell.setCellValue(val);
    }
  }

  @Override
  public String toString() {
    return "Sheet " + sheetName + " fr " + firstRowNumber +
	" lr " + lastRowNumber;
  }
}
