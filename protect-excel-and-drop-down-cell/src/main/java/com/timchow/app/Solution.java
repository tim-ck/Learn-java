package com.timchow.app;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;

import com.timchow.app.Entity.*;
import com.timchow.app.Entity.CustomeExcelSheet.RecordType;
import com.timchow.app.Util.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class Solution {
    public static void main(String[] args) {
        CreateExcelTemplateWithDropDownCell.createTemplate();
        List<CustomeRecord> customeRecordList = FakeData.genTypeARecords();

        downloadTypeAExcelForm(customeRecordList);

    }

    private static void downloadTypeAExcelForm(List<CustomeRecord> customeRecordList) {
        Map<RecordType, List<CustomeRecord>> customeRecordsListMap = new HashMap<>();
        for (CustomeRecord record : customeRecordList) {
            List<CustomeRecord> customeRecordsList = customeRecordsListMap.get(record.getRecordType());
            if (customeRecordsList == null) {
                customeRecordsList = new ArrayList<CustomeRecord>();
                customeRecordsListMap.put(record.getRecordType(), customeRecordsList);
            }
            customeRecordsList.add(record);
        }

        for (Map.Entry<RecordType, List<CustomeRecord>> entry : customeRecordsListMap.entrySet()) {
            RecordType recordType = entry.getKey();
            List<CustomeRecord> tmpRecordList = entry.getValue();
            if (tmpRecordList.isEmpty()) {
                // no record
                continue;
            }

            if (recordType == null) {
                //exception case: the record does not have correct language name.
                System.out.println("fail to download some of the record as they hv incorrect type name:");
                for (CustomeRecord record : tmpRecordList) {
                    System.out.println(record.getName());
                }

            }

            try (Workbook template = new HSSFWorkbook(new FileInputStream(recordType.templatePath))) {
                Date today = new Date();
                Sheet sheet = template.getSheet(recordType.sheetName);
                writeRecordToExcel(template, sheet, tmpRecordList, today);
                sheet.protectSheet("password");
                // unlock cell
                for (int i = 3; i <= FakeData.genTypeARecords().size() + 2; i++) {
                    ExcelUtil.unlockCellInProtectedSheet(sheet, i, 3);
                }
                StringBuilder sb = new StringBuilder("Record_"+recordType.recordTypeShortForm+"_");
                DateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HHmmss", Locale.getDefault());
                sb.append(df.format(today));
                sb.append(".xls");

                // create the file
                try (FileOutputStream out = new FileOutputStream(sb.toString())) {
                    template.write(out);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } catch (IOException e) {
                // template Path file not found
                e.printStackTrace();
            }

        }

    }

    public static void writeRecordToExcel(Workbook workbook, Sheet sheet,
            List<CustomeRecord> customeRecordList, Date generateDate) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet is null");
        }
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        cellStyle.setFont(font);
        DateFormat df = new SimpleDateFormat("yyyy.MM.dd.HHmmss", Locale.getDefault());
        Row titleRow = sheet.getRow(0);
        Cell createDate = titleRow.createCell(5);
        createDate.setCellValue("File generated time: " + df.format(generateDate));
        createDate.setCellStyle(cellStyle);

        for (int j = 0; j < customeRecordList.size(); j++) {
            Row dataRow = sheet.createRow(j + 3);
            dataRow.createCell(0).setCellValue(customeRecordList.get(j).getName());
            dataRow.createCell(1).setCellValue(customeRecordList.get(j).getGender());
            dataRow.createCell(2).setCellValue(customeRecordList.get(j).getRecordType().recordTypeShortForm);
            dataRow.createCell(3).setCellValue(customeRecordList.get(j).getStatus());
        }
    }
}
