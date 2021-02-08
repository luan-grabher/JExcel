package test;

import JExcel.XLSX;
import java.io.File;
import java.math.BigDecimal;
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
        data.put("collumn", "A");
        data.put("type", "date");
        data.put("required", "true");
        
        Map<String, String> historico = new HashMap<>();
        historico.put("name", "historico");
        historico.put("collumn", "D§E");
        historico.put("type", "string");
        historico.put("required", "true");
        historico.put("unifyDown", "D");
        historico.put("replace", "\\.§");
        
        Map<String, String> documento = new HashMap<>();
        documento.put("name", "documento");
        documento.put("collumn", "E");
        documento.put("type", "string");
        documento.put("required", "true");
        documento.put("replace", "[^0-9]+§");
        
        Map<String, String> valor = new HashMap<>();
        valor.put("name", "valor");
        valor.put("collumn", "F");
        valor.put("type", "value");
        valor.put("required", "true");
        valor.put("regex", "[-]?[0-9,.]+");
        valor.put("replace", "[^0-9,-]+§");
        valor.put("forceNegativeIf", "(?i).*[D].*");
        
        
        
        config.put("data", data);
        config.put("historico", historico);
        config.put("documento", documento);
        config.put("valor", valor);
        
        File file = new File("D:\\Downloads\\teste.xlsx");
        List<Map<String, Object>> rows = XLSX.get(file, config);
        
        
        BigDecimal[] total = new BigDecimal[]{new BigDecimal("0.00")};
        
        rows.forEach((row) ->{
            total[0] = total[0].add((BigDecimal) row.get("valor"));
        });
        
        System.out.println(total[0].toPlainString());
    }

}
