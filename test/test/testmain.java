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
        data.put("collumn", "B");
        data.put("type", "date");
        data.put("required", "true");
        
        Map<String, String> historico = new HashMap<>();
        historico.put("name", "historico");
        historico.put("collumn", "D");
        historico.put("type", "string");
        historico.put("required", "true");
        historico.put("unifyDown", null);
        historico.put("replace", "\\.§");
        
        Map<String, String> documento = new HashMap<>();
        documento.put("name", "documento");
        documento.put("collumn", "E");
        documento.put("type", "string");
        documento.put("required", "true");
        documento.put("replace", "[^0-9]+§");       
        
        Map<String, String> entrada = new HashMap<>();
        entrada.put("name", "entrada");
        entrada.put("collumn", "M");
        entrada.put("type", "value");
        entrada.put("required", "true");
        //entrada.put("regex", "[-]?[0-9,.]+");
        //entrada.put("replace", "[^0-9,-]+§");
        //entrada.put("forceNegativeIf", "(?i).*[D].*");
        
        Map<String, String> saida = new HashMap<>();
        saida.put("name", "saida");
        saida.put("collumn", "-O");
        saida.put("type", "value");
        saida.put("required", "true");
        //saida.put("regex", "[-]?[0-9,.]+");
        //saida.put("replace", "[^0-9,-]+§");
        //saida.put("forceNegativeIf", "(?i).*[D].*");
        
        
                
        /*Quando que pode começar a pegar*/
        Map<String, String> startGet = new HashMap<>();
        startGet.put("name", "startGet");
        startGet.put("collumn", "B");
        startGet.put("regex", "(?i).*CAIXA GERAL.*");
        
        /*Quando pode parar de pegar*/
        Map<String, String> endGet = new HashMap<>();
        endGet.put("name", "endGet");
        endGet.put("collumn", "A");
        endGet.put("regex", "(?i).*Conta.*");
        
        
        
        
        config.put("data", data);
        config.put("historico", historico);
        //config.put("documento", documento);
        config.put("entrada", entrada);
        config.put("saida", saida);
        config.put("startGet", startGet);
        config.put("endGet", endGet);
        
        
        
        
        File file = new File("C:\\Users\\User\\Documents\\NetBeansProjects\\Arquivos de teste\\pantano\\CAIXAS E BANCOS.xlsx");
        List<Map<String, Object>> rows = XLSX.get(file, config);
        
        
        BigDecimal[] entradas = new BigDecimal[]{new BigDecimal("0.00")};
        BigDecimal[] saidas = new BigDecimal[]{new BigDecimal("0.00")};
        
        rows.forEach((row) ->{
            entradas[0] = entradas[0].add((BigDecimal) row.get("entrada"));
            saidas[0] = saidas[0].add((BigDecimal) row.get("saida"));
        });
        
        System.out.println("Entradas: " + entradas[0]);
        System.out.println("Saidas: " + saidas[0]);
    }

}
