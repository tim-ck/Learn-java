/*
 * author: Tim Chow
 * Github: tim-ck
 */
package com.timchow.app;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import com.timchow.app.ExcelUtil;
import org.apache.poi.ss.usermodel.*;

public class ExcelExample {

	public static void main(String[] args) {
		// Example of how it works
		Workbook workbook = new HSSFWorkbook();
		Sheet sheet = workbook.createSheet("Sheet1");
		// create a row with header "Name", "Age", "preference"
		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Name");
		header.createCell(1).setCellValue("Age");
		header.createCell(2).setCellValue("Preference");
		// create a row with data "John", 20, "Blue"
		Row dataRow = sheet.createRow(1);
		dataRow.createCell(0).setCellValue("John");
		dataRow.createCell(1).setCellValue(20);
		dataRow.createCell(2).setCellValue("Blue");

		ExcelUtil.createDropDownCell(workbook, "Sheet1", 1, 2, 2, 2, new String[] { "Red", "Blue", "Green" });
		ExcelUtil.protectSheet(workbook, "Sheet1", new int[] { 2 },0, "password");

		// save excel to project directory
		try {
			FileOutputStream out = new FileOutputStream("workbook.xls");
			workbook.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	
}