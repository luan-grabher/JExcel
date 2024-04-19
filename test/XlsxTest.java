import java.io.File;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import JExcel.XLSX;

public class XlsxTest {
    public static void main(String[] args) {
        String filePath = "./test/arquivos/Contas RECEBIDAS CEF 03-2024.xlsx";
        File arquivo = new File(filePath);

        Map<String, String> colunaData = XLSX.getCollumnConfigFromString("data", "-collumn¬H¬-type¬date¬-required¬true");
        Map<String, String> colunaDocumento = XLSX.getCollumnConfigFromString("documento", "-collumn¬J¬-type¬value");
        Map<String, String> colunaHistorico = XLSX.getCollumnConfigFromString("historico", "-collumn¬I§C¬-type¬string¬-required¬true¬-replace¬\\.§");
        Map<String, String> colunaValor = XLSX.getCollumnConfigFromString("valor", "-collumn¬L¬-type¬value¬-required¬true¬-regex¬[-]?[0-9,.]+");

        
        Map<String,Map<String,String>> configuracaoDeLeitura = new HashMap<String,Map<String,String>>(); //Ajustar para ler o arquivo de configuração
        configuracaoDeLeitura.put("data", colunaData);
        configuracaoDeLeitura.put("documento", colunaDocumento);
        configuracaoDeLeitura.put("historico", colunaHistorico);
        configuracaoDeLeitura.put("valor", colunaValor);

        List <Map<String, Object>> linhasArquivo = XLSX.get(arquivo, configuracaoDeLeitura);

        for (Map<String, Object> linha : linhasArquivo) {
            System.out.println("historico: " + linha.get("historico"));
        }
    }
}
