package com.timchow.app.Util;

import com.timchow.app.Entity.CustomeExcelSheet;
import java.util.ArrayList;
import java.util.List;
import com.timchow.app.Entity.CustomeExcelSheet.RecordType;


public class CreateExcelTemplateWithDropDownCell {

    public static void createTemplate(){
        String title = "Title";
        String subTitle = "Subtitle ";
        CustomeExcelSheet typeAExcelSheet = new CustomeExcelSheet(RecordType.typeA,  title, subTitle,
                getHeader(), null, genTypeATemplate()); 
        ExcelUtil.createCatCExcelTemplate(typeAExcelSheet, typeAExcelSheet.getRecordType().templatePath);

    }


    private static List<String> getHeader() {
        List<String> header = new ArrayList<>();
        header.add("Name");
        header.add("gender");
        header.add("recordTypeShortForm");
        header.add("status");
        return header;
    }

    public static List<String> genTypeATemplate() {
        List<String> typeATemplate = new ArrayList<String>();
        typeATemplate.add("Pending");
        typeATemplate.add("TypeA - Pass");
        typeATemplate.add("TypeA - Fail");
        typeATemplate.add("The candidate is not found in our record");
        return typeATemplate;
    }
}
