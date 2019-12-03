package asst.thatsbiz.poiBug;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import asst.hssf.WorkbookManager;

/** Read an Excel file and modify it to demonstrate a bug
 * @author wataylor
 * @since 2019 12
 *
 */
public class ReaditAndWeepMain {

  /** Array of valid image file extensions.  */
  public static final String[] excelFileExts = {
    "xls", "xlsx" };

  /**
   * Check a file name for ending in an excel file extension
   * @param fileName name to be checked
   * @return the extension or null if it is incorrect.
   */
  public static String excelFileExt(String fileName) {
    int ix;
    String ext;

    if ( (ix = fileName.lastIndexOf(".")) <= 0) { return null; }
    ext = fileName.substring(ix+1);
    for (ix=0; ix<excelFileExts.length; ix++) {
      if (excelFileExts[ix].equalsIgnoreCase(ext)) {
        ext = excelFileExts[ix];
        return ext;
      }
    }
    return null;
  }

  /**
   * Give a file name a new extension or strip off the extension if the new
   * extension is the empty string.
   * 
   * @param fileName
   *          the existing file name
   * @param ext
   *          the new extension
   * @return updated file name
   */
  public static String newFileType(String fileName, String ext) {
    int ix;

    if (!fileName.endsWith("." + ext)) {
      if ((ix = fileName.lastIndexOf(".")) >= 0) {
	fileName = fileName.substring(0, ix);
      }
    }
    if (ext.length() <= 0) {
      return fileName;
    }
    return fileName + "." + ext;
  }

  /**
   * @param args space-separated list of input files.
   */
  public static void main(String[] args) {
    if (args.length <= 0) {
      System.out.println("Must pass files on command line");
      System.exit(-1);
    }
    String ext;
    boolean isXLSX;

    for (String fn : args) {
      File file = new File(fn);
      if (!file.canRead()) {
	System.out.println("Cannot read file " + fn);
	continue;
      }
      if ( (ext = excelFileExt(fn)) == null) {
	System.out.println("Invalid file extension for " + fn +
	    ". Please supply an .xls or .xlsx file.");
	continue;
      }
      InputStream is;
      WorkbookManager wm;
      Row row;
      Cell cell;
      try {
	is = new FileInputStream(file);
      } catch (FileNotFoundException e) {
	System.out.println("Issue with file " + fn + " " + e.getMessage());
	continue;
      }
      try {
	if (ext.endsWith(".xls")) {
	  isXLSX = false;
	  HSSFWorkbook wb = new HSSFWorkbook(is);
	  wm = new WorkbookManager(wb);
	} else {
	  isXLSX = true;
	  XSSFWorkbook wb = new XSSFWorkbook(is);
	  wm = new WorkbookManager(wb);
	}
      } catch (IOException e) {
	System.out.println("Not create WB for file " + fn + " "
	    + e.getMessage());
	continue;
      }
      row = wm.nextRow();
      row = wm.nextRow();
      row.createCell(0).setCellValue("Done Donner Donnest");
      System.out.println("Row " + (wm.currentRowNumber + 1)
	  + " of " + wm.sheetName + " set to " + row.getCell(0).getStringCellValue());
      String value = null;
      try {
	if (isXLSX) {
	  value = newFileType(args[0], "") + "New.xlsx";
	  FileOutputStream out = new FileOutputStream(value);
	  wm.wb.write(out);
	  System.out.println("Wrote new spread sheet " + value);
	} else {
	  value = newFileType(args[0], "") + "New.xls";
	  FileOutputStream out = new FileOutputStream(value);
	  wm.wb.write(out);
	  System.out.println("Wrote new spread sheet " + value);
	}
      } catch (Exception e) {
	System.out.println("Issue writing file " + value + " " + e.getMessage());
	e.printStackTrace();
      }
    }
  }

}
