package com.timchow.app.Util;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import com.timchow.app.Entity.CustomeExcelSheet.RecordType;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;

import com.timchow.app.Entity.CustomeExcelSheet;

public class ExcelUtil {


	public static void getLastRowNum(Sheet sheet) {
		System.out.println(sheet.getLastRowNum());
	}

	public static void createExcelTemplate(CustomeExcelSheet customeExcelSheet, String fileName) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		RecordType recordType = customeExcelSheet.getRecordType();
		String title = customeExcelSheet.getTitle();
		String subTitle = customeExcelSheet.getSubTitle();
		Sheet sheet = workbook.createSheet(recordType.sheetName);

		Row titleRow = sheet.createRow(0);
		titleRow.createCell(0).setCellValue(title);

		Row subTitleRow = sheet.createRow(1);
		subTitleRow.createCell(0).setCellValue(subTitle);

		Row headerRow = sheet.createRow(2);
		headerRow.setHeight((short) 1000);
		CellStyle headerCellStyle = workbook.createCellStyle();
		headerCellStyle.setWrapText(true);
		headerCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
		for (int j = 0; j < customeExcelSheet.getRecordType().sheetHeader.length; j++) {
			headerRow.createCell(j).setCellValue(customeExcelSheet.getRecordType().sheetHeader[j]);
			sheet.setColumnWidth(j, 5000);
			headerRow.getCell(j).setCellStyle(headerCellStyle);
		}

		sheet.setColumnWidth(4, 10000);
		
		Sheet referenceSheet = workbook.createSheet("table");
		createReferenceSheet(workbook, referenceSheet, customeExcelSheet.getRecordType().responseTemplate, recordType.recordTypeShortForm);
		referenceSheet.protectSheet("passwordForReferenceSheet");
				
		createDropDownCellByReferenceSheet(workbook, sheet, 3, 1000, 3, 3, recordType.recordTypeShortForm, "table",
				0, 0, 1, customeExcelSheet.getRecordType().responseTemplate.length);

		try (FileOutputStream out = new FileOutputStream(fileName)) {
			workbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}	


	/**
	 * Creates reference sheet with single Validation List.
	 * 
	 * <p>
	 * all row and column index start from 0.
	 */
	public static void createReferenceSheet(Workbook wb, Sheet referenceSheet,
			String[] dataValidationList, String dataValidationName) {
		Row row = referenceSheet.createRow(0);
		Cell sheetNameCell = row.createCell(0);
		sheetNameCell.setCellValue(dataValidationName);
		CellStyle style = wb.createCellStyle();
		style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		sheetNameCell.setCellStyle(style);
		for (int i = 0; i < dataValidationList.length; i++) {
			Row dataRow = referenceSheet.createRow(i + 1);
			dataRow.createCell(0).setCellValue(dataValidationList[i]);
		}
	}


	/**
	 * Creates drop down list cell in sheet with defined position
	 * using data validation source from reference sheet.
	 * 
	 * <p>
	 * before calling this method, you need to create a reference sheet with data
	 * validation list.
	 * all row and column index start from 0.
	 */
	public static void createDropDownCellByReferenceSheet(Workbook wb, Sheet sheet,
			int firstRow, int lastRow, int firstCol, int lastCol,
			String namedCellName, String referenceSheetName, int referenceSheetStartColumn, int referenceSheetEndColumn,
			int referenceSheetStartRow, int referenceSheetEndRow) {
		Name namedCell = wb.createName();
		namedCell.setNameName(namedCellName);
		String startColumn = Character.toString((char) ('A' + referenceSheetStartColumn));
		String endColumn = Character.toString((char) ('A' + referenceSheetEndColumn));
		namedCell.setRefersToFormula(referenceSheetName + "!$" + startColumn + "$" + (referenceSheetStartRow + 1) + ":$"
				+ endColumn + "$" + (referenceSheetEndRow + 1));
		DVConstraint constraint = DVConstraint.createFormulaListConstraint(namedCellName);
		CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
		HSSFDataValidation validation = new HSSFDataValidation(addressList, constraint);
		sheet.addValidationData(validation);
	}

	public static void unlockCellInProtectedSheet(Sheet sheet, int rowIndex, int colIndex) {
		Row row = sheet.getRow(rowIndex);
		Cell cell = row.getCell(colIndex);
		CellStyle unlockedCellStyle = sheet.getWorkbook().createCellStyle();
		unlockedCellStyle.setLocked(false);
		if(cell != null){
			cell.setCellStyle(unlockedCellStyle);
		}
	}

	public static Workbook readExcelFromFilePath(String filePath) {
		Workbook workbook = null;
		try {
			workbook = new HSSFWorkbook(new FileInputStream(filePath));
		} catch (IOException e) {
			e.printStackTrace(); 
		}
		return workbook;
	}
}