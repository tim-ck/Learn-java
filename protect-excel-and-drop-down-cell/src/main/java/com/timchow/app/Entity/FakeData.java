package com.timchow.app.Entity;

import java.util.ArrayList;
import java.util.List;
public class FakeData {
    public static List<CustomeRecord> genTypeARecords() {
            List<CustomeRecord> result = new ArrayList<>();
            result.add(new CustomeRecord("Tim", "M", "Type A", ""));
            result.add(new CustomeRecord("Tom", "M", "Type A", ""));
            result.add(new CustomeRecord("Tina", "F", "Type A", ""));
            return result;
    }

}