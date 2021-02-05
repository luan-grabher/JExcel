package test;

import JExcel.XLSX;
import java.io.File;
import java.util.Calendar;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class testmain {

    public static void main(String[] args) {
        test();       
        
    }
    public static void test(){
        Map<String, Map<String, String>> config = new HashMap<>();

        Map<String, String> data = new HashMap<>();
        data.put("name", "data");
        data.put("collumn", "B");
        data.put("type", "date");
        data.put("collumn", "B");
        data.put("required", "true");
        
        config.put("data", data);
        

        File file = new File("C:\\Users\\Admnistrador\\Desktop\\teste.xlsx");
        List<Map<String, Object>> rows = XLSX.get(file, config);
        
        rows.forEach((row) ->{
            System.out.println(
                    Dates.Dates.getCalendarInThisStringFormat((Calendar) row.get("data"), "dd/MM/yyyy")
            );
        });
    }

}
