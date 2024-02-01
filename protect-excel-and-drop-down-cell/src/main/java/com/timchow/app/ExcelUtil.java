/*
 * author: Tim Chow
 * Github: tim-ck
 */
package com.timchow.app;

import java.util.List;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddressList;



public class ExcelUtil {

	/**
	 * protectSheet but unlock list specified col
	 * The sheetName argument must match the sheet name in the workbook argument.
	 * 
	 * if your sheet contain header, the indexOfFirstUnlockRow should set to 1
	 *
	 * 	 *<p>
	 * all row and column index start from 0.
	 * 
	 * @param wb              org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName       name of the sheet
	 * @param listOfUnlockCol list of Column index that need to be editable for
	 *                        user.
	 * @param password        password protection for the locked sheet
	 */
	public static void protectSheet(Workbook wb, String sheetName, int[] listOfUnlockCol, int indexOfFirstUnlockRow,
			String password) {
		Sheet sheet = wb.getSheet(sheetName);
		CellStyle unlockedCellStyle = wb.createCellStyle();
		unlockedCellStyle.setLocked(false);
		// apply style to all cell in col
		for (int col : listOfUnlockCol) {
			for (int i = indexOfFirstUnlockRow; i <= sheet.getLastRowNum(); i++) {
				sheet.getRow(i).getCell(col).setCellStyle(unlockedCellStyle);
			}
		}
		sheet.protectSheet(password);
	}

	/**
	 * Freeze panel in the sheet.
	 * The sheetName argument must match the sheet name in the workbook argument.
	 * row and col argument define number of column/row freeze in the sheet. set to
	 * 0 if no row/ col needed to be freeze.
	 * 
	 * 	 *<p>
	 * all row and column index start from 0.
	 *
	 * @param wb        org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName name of the sheet
	 * @param colSplit  Horizontal position of split.
	 * @param rowSplit  Vertical position of split.
	 */
	public static void freezePanel(Workbook wb, String sheetName, int colSplit, int rowSplit) {
		Sheet sheet = wb.getSheet(sheetName);
		sheet.createFreezePane(colSplit, rowSplit);
		return;
	}

	/**
	 * Creates drop down list cell in sheet with defined position. Last row is set
	 * to last row index in the sheet.
	 *
	 * 	 *<p>
	 * all row and column index start from 0.
	 * 
	 * @param wb                 org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName          name of the sheet
	 * @param firstRow           Start Vertical position of split.
	 * @param firstCol           Start Horizontal position of split.
	 * @param firstCol           End Horizontal position of split.
	 * @param dataValidationList String array of drop down list data
	 */
	public static void createDropDownCell(Workbook wb, String sheetName,
			int firstRow, int firstCol, int lastCol, String[] dataValidationList) {
		Sheet sheet = wb.getSheet(sheetName);
		sheet.getLastRowNum();
		createDropDownCell(wb, sheetName, firstRow, sheet.getLastRowNum(), firstCol, lastCol,
				dataValidationList);
	}

	/**
	 * Creates drop down list cell in sheet with defined position
	 *
	 *<p>
	 * all row and column index start from 0.
	 * 
	 * @param wb                 org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName          name of the sheet
	 * @param firstRow           Start Vertical position of split.
	 * @param lastRow            End Vertical position of split.
	 * @param firstCol           Start Horizontal position of split.
	 * @param lastCol           End Horizontal position of split.
	 * @param dataValidationList String array of drop down list data
	 */
	public static void createDropDownCell(Workbook wb, String sheetName,
			int firstRow, int lastRow, int firstCol, int lastCol, String[] dataValidationList) {
		Sheet sheet = wb.getSheet(sheetName);
		CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		DVConstraint constraint = DVConstraint.createExplicitListConstraint(dataValidationList);
		HSSFDataValidation validation = new HSSFDataValidation(addressList, constraint);
		sheet.addValidationData(validation);
	}

	/**
	 * Creates drop down list cell in sheet with defined position 
	 * using data validation source from reference sheet.
	 * 
	 * <p>
	 * before calling this method, you need to create a reference sheet with data validation list.
	 * 
	 * <p>
	 * all row and column index start from 0.
	 * 
	 * 
	 * @param wb                 org.apache.poi.ss.usermodel.Workbook
	 * @param sheetName          name of the sheet
	 * @param firstRow           Start Vertical position of split.
	 * @param lastRow            End Vertical position of split.
	 * @param firstCol           Start Horizontal position of split.
	 * @param lastCol           End Horizontal position of split.
	 * @param namedCellName      name of the reference sheet
	 * @param referenceSheetName name of the reference sheet
	 * @param referenceSheetStartColumn Start Horizontal position of reference sheet.
	 * @param referenceSheetEndColumn End Horizontal position of reference sheet.
	 * @param referenceSheetStartRow Start Vertical position of reference sheet.
	 * @param referenceSheetEndRow End Vertical position of reference sheet.
	 */
	public static void createDropDownCellUsingDataValidationSourceFromReferenceSheet(Workbook wb, String sheetName,
			int firstRow, int lastRow, int firstCol, int lastCol,
			String namedCellName, String referenceSheetName, int referenceSheetStartColumn, int referenceSheetEndColumn,
			int referenceSheetStartRow, int referenceSheetEndRow) {
		Sheet sheet = wb.getSheet(sheetName);
		Name namedCell = wb.createName();
		namedCell.setNameName(namedCellName);
		String startColumn = Character.toString((char) ('A' + referenceSheetStartColumn));
		String endColumn = Character.toString((char) ('A' + referenceSheetEndColumn));
		namedCell.setRefersToFormula(
				referenceSheetName + "!$" + startColumn + "$" + (referenceSheetStartRow+1) + ":$" + endColumn + "$"
						+ (referenceSheetEndRow+1));
		DVConstraint constraint = DVConstraint.createFormulaListConstraint(namedCellName);
		CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		HSSFDataValidation validation = new HSSFDataValidation(addressList, constraint);
		sheet.addValidationData(validation);
	}

	/**
	 * Creates data validation list in the reference sheet.
	 * 
	 * 	 *<p>
	 * all row and column index start from 0.
	 * 
	 * @param wb                 	org.apache.poi.ss.usermodel.Workbook
	 * @param referenceSheet		reference sheet
	 * @param rowIndex				row index of the reference sheet to save the data validation list
	 * @param dataValidationList	List<String> of drop down list data
	 * @param dataValidationName	name of the data validation list
	 */
	public static void createDataValidationSourceForDropDownCellToReferenceSheet(Workbook wb, Sheet referenceSheet, 
		int rowIndex, List<String> dataValidationList, String dataValidationName){
			Row row = referenceSheet.createRow(rowIndex);
			Cell sheetNameCell = row.createCell(0);
			sheetNameCell.setCellValue(dataValidationName);
			CellStyle style = wb.createCellStyle();
			style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
			style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
			sheetNameCell.setCellStyle(style);
			for (int j = 0; j < dataValidationList.size(); j++) {
				Cell cell2 = row.createCell(j + 1);
				cell2.setCellValue(dataValidationList.get(j));
			}
		}

}