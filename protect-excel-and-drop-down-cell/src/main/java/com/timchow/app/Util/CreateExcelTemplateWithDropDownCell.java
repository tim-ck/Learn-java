package com.timchow.app.Util;

import com.timchow.app.Entity.CustomeExcelSheet;
import java.util.ArrayList;
import java.util.List;
import com.timchow.app.Entity.CustomeExcelSheet.RecordType;


public class CreateExcelTemplateWithDropDownCell {

    public static void createTemplate(){
        String title = "Title";
        String subTitle = "Subtitle ";
        CustomeExcelSheet typeAExcelSheet = new CustomeExcelSheet(RecordType.typeA,  title, subTitle, null); 
        ExcelUtil.createExcelTemplate(typeAExcelSheet, typeAExcelSheet.getRecordType().templatePath);

    }
    
}
