package com.timchow.app.Entity;
import com.timchow.app.Entity.CustomeExcelSheet.RecordType;
public class CustomeRecord {
    private String name;
    private String gender;
    private RecordType recordType;
    private String status;

    // Constructor
    public CustomeRecord() {
    }
    
    public CustomeRecord(String name, String gender, String recordTypeName, String status) {
        this.name = name;
        this.gender = gender;
        this.recordType = RecordType.fromrecordTypeName(recordTypeName);
        this.status = status;  
    }

    // Getter
    public String getName() {
        return name;
    }

    public String getGender() {
        return gender;

    }

    public RecordType getRecordType() {
        return recordType;
    }

    public String getStatus(){
        return status;
    }

    // Setter

    public void setName(String name) {
        this.name = name;
    }

    public void setGender(String gender) {
        this.gender = gender;
    }

    public void setRecordType(RecordType recordType) {
        this.recordType = recordType;
    }
    

}
