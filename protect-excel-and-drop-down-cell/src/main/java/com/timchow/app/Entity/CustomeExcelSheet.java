package com.timchow.app.Entity;

import java.util.ArrayList;
import java.util.List;

public class CustomeExcelSheet {
    public enum RecordType {

        typeA("Type A", "sheet name", "a",
                new String[] { "name", "gender", "recordTypeShortForm", "status" },
                new String[] { "Pending", "TypeA - Pass", "TypeA - Fail", "The candidate is not found in our record" },
                "template.xls");

        public final String recordTypeName;
        public final String sheetName;
        public final String recordTypeShortForm;
        public final String[] sheetHeader;
        public final String[] responseTemplate;
        public final String templatePath;

        private RecordType(String recordTypeName, String sheetName, String recordTypeShortForm, String[] sheetHeader,
                String[] responseTemplate, String templatePath) {
            this.recordTypeName = recordTypeName;
            this.sheetName = sheetName;
            this.recordTypeShortForm = recordTypeShortForm;
            this.sheetHeader = sheetHeader;
            this.responseTemplate = responseTemplate;
            this.templatePath = templatePath;
        }

        public static RecordType fromrecordTypeName(String recordTypeName) {
            for (RecordType s : RecordType.values()) {
                if (s.recordTypeName.equals(recordTypeName))
                    return s;
            }
            throw new IllegalArgumentException(
                    "Illegal CatCExamRecord recordTypeName (cannot covert recordTypeName to CatCExamRecord.Subject): "
                            + recordTypeName);
        }

    }

    private RecordType recordType;
    private String title;
    private String subTitle;
    private List<CustomeRecord> examRecords;

    public CustomeExcelSheet() {
        examRecords = new ArrayList<>();
    }

    public CustomeExcelSheet(RecordType recordType, String title, String subTitle, List<CustomeRecord> examRecords) {
        this.recordType = recordType;
        this.title = title;
        this.subTitle = subTitle;
        this.examRecords = examRecords;
    }

    public RecordType getRecordType() {
        return recordType;
    }

    public void setRecordType(RecordType recordType) {
        this.recordType = recordType;
    }

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public String getSubTitle() {
        return subTitle;
    }

    public void setSubTitle(String subTitle) {
        this.subTitle = subTitle;
    }

    public List<CustomeRecord> getExamRecords() {
        return examRecords;
    }

    public void setExamRecords(List<CustomeRecord> examRecords) {
        this.examRecords = examRecords;
    }

}
