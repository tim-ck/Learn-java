package com.timchow.app.Entity;

import java.util.ArrayList;
import java.util.List;

public class CustomeExcelSheet {
    public enum RecordType{
        typeA( "Type A", "sheet name", "a", "typeATemplate.xls");

        public final String recordTypeName;
        public final String sheetName;
        public final String recordTypeShortForm;
        public final String templatePath;
        private RecordType(String recordTypeName, String sheetName, String recordTypeShortForm, String templatePath) {
            this.recordTypeName = recordTypeName;
            this.sheetName = sheetName;
            this.recordTypeShortForm = recordTypeShortForm;
            this.templatePath = templatePath;
        }

        public static RecordType fromrecordTypeName(String recordTypeName){
            for(RecordType s: RecordType.values()){
                if(s.recordTypeName.equals(recordTypeName))
                    return s;
            }
            throw new IllegalArgumentException("Illegal CatCExamRecord recordTypeName (cannot covert recordTypeName to CatCExamRecord.Subject): " + recordTypeName);
        }

    }
    
    private RecordType recordType;
    private String title;
    private String subTitle;
    private List<String> sheetHeader;
    private List<CustomeRecord> examRecords;
    private List<String> responseTemplate;
    

    public CustomeExcelSheet(){
        sheetHeader = new ArrayList<>();
        examRecords = new ArrayList<>();
        responseTemplate = new ArrayList<>();
    }

    public CustomeExcelSheet(RecordType recordType,  String title, String subTitle, List<String> sheetHeader, List<CustomeRecord> examRecords, List<String> responseTemplate) {
        this.recordType = recordType;
        this.title = title;
        this.subTitle = subTitle;
        this.sheetHeader = sheetHeader;
        this.examRecords = examRecords;
        this.responseTemplate = responseTemplate;
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

    public List<String> getSheetHeader() {
        return sheetHeader;
    }

    public void setSheetHeader(List<String> sheetHeader) {
        this.sheetHeader = sheetHeader;
    }

    public List<CustomeRecord> getExamRecords() {
        return examRecords;
    }

    public void setExamRecords(List<CustomeRecord> examRecords) {
        this.examRecords = examRecords;
    }

    public List<String> getResponseTemplate() {
        return responseTemplate;
    }

    public void setResponseTemplate(List<String> responseTemplate) {
        this.responseTemplate = responseTemplate;
    }


}
