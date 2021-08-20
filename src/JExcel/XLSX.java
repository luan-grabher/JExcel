package JExcel;

import fileManager.Args;
import fileManager.CSV;
import java.io.File;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class XLSX {

    /**
     * Converte um texto definindo as configurações da coluna da mesma forma que
     * argumentos são definidos nos atalhos do Windows, usando "-" na frente e
     * depois o valor, porém ao invés de separar com espaços a separação é com
     * "¬".
     * <br>
     * Os argumentos são os mesmos do method "get" desta classe.
     *
     * @param collumnName O nome da coluna
     * @param iniString O texto com os argumentos declarados.
     */
    public static Map<String, String> getCollumnConfigFromString(String collumnName, String iniString) {
        if (!"".equals(iniString) && iniString != null) {
            Map<String, String> config = new HashMap<>();
            String[] configs = iniString.split("¬", -1);

            config.put("name", collumnName);
            config.put("collumn", Args.get(configs, "collumn"));
            config.put("regex", Args.get(configs, "regex"));
            config.put("replace", Args.get(configs, "replace"));
            config.put("type", Args.get(configs, "type"));
            config.put("dateFormat", Args.get(configs, "dateFormat"));
            config.put("required", Args.get(configs, "required"));
            config.put("requiredBlank", Args.get(configs, "requiredBlank"));
            config.put("unifyDown", Args.get(configs, "unifyDown"));
            config.put("forceNegativeIf", Args.get(configs, "forceNegativeIf"));

            return config;
        } else {
            return null;
        }
    }

    /**
     * Retorna uma lista das linhas de um arquivo xlsx com um mapa com o valores
     * das colunas de cada linha conforme as configurações definidas.
     * <p>
     * Para procurar um regex no meio de algo e ignorar case: "(?i).*" + search
     * + ".*"<br>
     * <p>
     * Configurações de cada coluna:
     * <p>
     * Para colunas Booleanas - utilizar 'true' para verdadeiro e qualquer outra
     * coisa, inclusive null, para false
     * <br>
     * -name<br>
     * -collumn: Caso tenha que unir colunas, separe por §. Caso o que estiver
     * entre os § for um caractere somente, será pego o valor da coluna, se não
     * será adicionado a palavra escrita no resultado.<br>
     * -regex: Se após converter a data e fazer os replaces não for match do
     * regex, não pega<br>
     * -replace: Separa o regex do replace com § por exemplo "aa§bb" para
     * substituir todos "aa" por "bb"<br>
     * -type:Tipo de Objeto: string,value,date<br>
     * -dateFormat:Formato da data: Se tiver data dd/MM/yyyy (BR)<br>
     * -required: Se é Obrigatoria ou não, se for obrigatória e não tiver valor
     * ou nao for válida, não pega a linha<br>
     * -requiredBlank: Tem que estar em branco: bool<br>
     * -unifyDown: UnirColunaAbaixo: Coluna(s) em baixo que vai ser unida no
     * resultado. Para não tiizar deixe em branco ou não declare.<br>
     * -forceNegativeIf: Coloca um "-" na frente se o tipo for valor e se
     * possuir o regex. Utilize regex. <br>
     *
     * @param file Arquivo XLSX
     * @param config Configuração das colunas em mapa
     * @return lista das linhas válidas com mapa de colunas
     *
     */
    public static List<Map<String, Object>> get(File file, Map<String, Map<String, String>> config) {
        List<Map<String, Object>> rows = new ArrayList<>();

        try {
            /*Verifica CSV*/
            if (file.getName().toLowerCase().endsWith(".csv")) {
                //Se for CSV converte para arquivo xslx
                file = getCSV(file);
            }

            XSSFWorkbook wk;
            XSSFSheet sheet;

            wk = new XSSFWorkbook(file);
            sheet = wk.getSheetAt(0);

            //Remove todas config NULL
            while (config.values().remove(null));
            config.forEach((n, c) -> {
                //Remove todas colunas NULL
                while (c.values().remove(null));
            });

            /*Variavel que permite a adição de colunas, por padrão define como true */
            Boolean[] canAdd = new Boolean[]{true};
            //Se tiver a configuração para filtro de start e end do get
            if (config.containsKey("startGet")) {
                //Define o canAdd para false
                canAdd[0] = false;
            }

            for (Row row : sheet) {
                //Cria mapa de colunas
                Map<String, Object> cols = new HashMap<>();
                Boolean[] rowValid = new Boolean[]{true};

                try {
                    //Percorre todas colunas das configurações
                    config.forEach((String name, Map<String, String> col) -> {
                        /*Se o nome da configuração não for a que gerencia o inicio e finalização de gets, poder adicionar colunas ativo e tiver configuração de colunas, verifica as colunas*/
                        if (col != null) {

                            //Pega o objeto da coluna
                            Object colObj = getCollumnVal(row, col);
                                                        
                            //Se o nome for startGet ou EndGet
                            if("startGet".equals(name) || "endGet".equals(name)){
                                //Se tiver conseguido pegar o valor da linha
                                if(colObj != null && !"".equals(colObj)){
                                    //Se for startGet, define o canAdd como true, se for o endGet, define o canAdd como false
                                    canAdd[0] = "startGet".equals(name);
                                    //Vai para a proxima linha
                                    throw new Error(name);
                                }
                            }

                            //Se estiver perrmitido adicionar
                            if(canAdd[0] && !"startGet".equals(name) && !"endGet".equals(name)){
                                //Se for tipo string e tiver que juntar com a proxima linha
                                if ("string".equals(col.getOrDefault("type", "string"))
                                        && !"".equals(col.getOrDefault("unifyDown", ""))) {
                                    //Pega o valor da proxima linha
                                    Object nextRowCol = getNextCollumnVal(sheet.getRow(row.getRowNum() + 1), col);
                                    //Se pelo menos um valor não for null
                                    if (colObj != null || nextRowCol != null) {
                                        //Tansforma os valores null em "" e junta os dois no colObj
                                        colObj = (String) (colObj == null ? "" : colObj.toString());
                                        colObj += (String) (colObj.toString().equals("") ? "" : " " + (nextRowCol == null ? "" : nextRowCol.toString()));
                                    }
                                }

                                //Se for required e não for null OU se não for required
                                if (Boolean.TRUE.equals(Boolean.valueOf(col.get("required")))
                                        && (colObj == null || colObj.equals(""))) {
                                    rowValid[0] = false;
                                } else {
                                    //Se não tiver que ser em branco o tiver que ser em branco e o objeto for null ou
                                    if (Boolean.TRUE.equals(Boolean.valueOf(col.get("requiredBlank")))
                                            && colObj != null && !colObj.equals("")) {
                                        rowValid[0] = false;
                                    } else {
                                        cols.put(name, colObj);
                                    }
                                }
                            }
                        }
                    });

                    if (rowValid[0] == true && cols.size() > 0) {
                        rows.add(cols);
                    }
                } catch (Error e) {
                    //Se o break for para dar end
                    if(e.getMessage().equals("endGet")){
                        //sai do loop pois já pegou tudo;
                        break;
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("");
        }
        return rows;
    }

    /**
     * Retorna o objeto String/BigDecimal/Calendar conforme configuração
     *
     * @param row Linha da sheet
     * @param colConfig Configuração da coluna
     */
    private static Object getCollumnVal(Row row, Map<String, String> colConfig) {
        return getCollumnVal(row, colConfig, "collumn");
    }

    /**
     * Retorna o objeto String/BigDecimal/Calendar conforme configuração
     *
     * @param row Linha da sheet
     * @param colConfig Configuração da coluna
     */
    private static Object getNextCollumnVal(Row row, Map<String, String> colConfig) {
        return getCollumnVal(row, colConfig, "unifyDown");
    }

    /**
     * Retorna o objeto String/BigDecimal/Calendar conforme configuração
     *
     * @param row Linha da sheet
     * @param colConfig Configuração da coluna
     * @param nameMapCollumns nome do vetor que tem as colunas
     */
    private static Object getCollumnVal(Row row, Map<String, String> colConfig, String nameMapCollumns) {

        try {
            String stringVal = getStringOfCols(row, colConfig.getOrDefault(nameMapCollumns, "").split("§"));
            if (!stringVal.equals("")) {
                //Converte data se for tipo data e estiver no formato de numero
                if ("date".equals(colConfig.getOrDefault("type", "string"))
                        && stringVal.matches("[0-9]+[.][0-9]+")) {
                    Integer dateInt = Integer.valueOf(stringVal.split("\\.")[0]);
                    stringVal = JExcel.getStringDate(dateInt);
                } else if ("value".equals(colConfig.getOrDefault("type", "string"))
                        && !stringVal.equals("")
                        && !"".equals(colConfig.getOrDefault("forceNegativeIf", ""))) {
                    if (stringVal.matches(colConfig.get("forceNegativeIf"))) {
                        stringVal = "-" + stringVal;
                    }
                }

                //Aplica replace se tiver replace e nao estiver em branco
                if (!"".equals(colConfig.getOrDefault("replace", ""))) {
                    String[] replaces = colConfig.get("replace").split("§", -1);
                    if (replaces.length == 2) {
                        stringVal = stringVal.replaceAll(replaces[0], replaces[1]);
                    }
                }

                //Continua se nao tiver filtro de regex ou for match do regex
                if ("".equals(colConfig.getOrDefault("regex", ""))
                        || (!"".equals(colConfig.getOrDefault("regex", ""))
                        && stringVal.matches(colConfig.get("regex")))) {
                    String type = colConfig.getOrDefault("type", "string");
                    switch (type) {
                        case "string":
                            //Se for tipo string retorna string
                            return stringVal;
                        case "value":
                            Boolean forceNegative = colConfig.get("collumn").startsWith("-");
                            return getBigDecimalFromCell(stringVal, forceNegative);
                        case "date":
                            return Dates.Dates.getCalendarFromFormat(stringVal, colConfig.getOrDefault("dateFormat", "dd/MM/yyyy"));
                        default:
                            break;
                    }
                }
            }
        } catch (Exception e) {
        }
        return null;
    }

    /**
     * Retorna uma String de todas colunas que tiverem que pegar Para apenas
     * colocar algo entre as colunas, o vetor deve possuir mais de um caractere
     *
     * @param row Linha da Sheet
     * @param cols Colunas a serem pegas ou textos a serem colocados
     */
    private static String getStringOfCols(Row row, String[] cols) {
        StringBuilder result = new StringBuilder("");

        for (String col : cols) {
            if (col.length() == 1 || (col.length() == 2 && col.startsWith("-"))) {
                String prepend = "";

                if (col.length() == 2 && col.startsWith("-")) {
                    col = col.replaceAll("-", "");
                    prepend = "-";
                }

                Cell cel = row.getCell(JExcel.Cell(col));
                if (cel != null) {
                    if (!result.toString().equals("")) {
                        result.append(" ");
                    }
                    result.append(prepend);
                    result.append(JExcel.getStringCell(cel));
                }
            } else if (col.length() > 1) {
                if (!result.toString().equals("")) {
                    result.append(" ");
                }
                result.append(col);
            }
        }

        return result.toString();
    }

    /**
     * Pega bigdecimal de uma celula do excel numerica
     *
     * @param cell Celula que ira pegar numero
     * @param forceNegative Se deve multiplicar por -1 o numero se for positivo
     * @return celula em número BigDecimal
     */
    private static BigDecimal getBigDecimalFromCell(String celString, boolean forceNegative) {
        //Pega texto das celulas
        String valueString = celString != null ? celString : "0.00";

        try {
            valueString = valueString.replaceAll("[^0-9E\\.,-]", "");

            //Se tiver . antes da virgula remove os pontos e coloca ponto no lugar da virgula
            if (valueString.indexOf(".") < valueString.indexOf(",")) {
                valueString = valueString.replaceAll("\\.", "").replaceAll("\\,", ".");
            }
        } catch (Exception e) {
        }

        BigDecimal valueBigDecimal = new BigDecimal(valueString.equals("") ? "0" : valueString);

        //Se a coluna tiver que multiplicar por -1 e o valor encontrado for maior que zero
        if (forceNegative && valueBigDecimal.compareTo(BigDecimal.ZERO) > 0) {
            valueBigDecimal = valueBigDecimal.negate();
        }

        return valueBigDecimal;
    }

    /**
     * Converte String, geralmente em arquivo ini para mapa de configuração de
     * coluna.
     *
     *
     * @param collumnName Nome da coluna
     * @param iniString String com configuração separada por "¬" e definidos por
     * "-" na frente de cada argumento.
     * @return Mapa de configuração de coluna
     */
    public static Map<String, String> convertStringToConfig(String collumnName, String iniString) {
        if (!"".equals(iniString) && iniString != null) {
            Map<String, String> config = new HashMap<>();
            String[] configs = iniString.split("¬", -1);

            config.put("name", collumnName);
            config.put("collumn", Args.get(configs, "collumn"));
            config.put("regex", Args.get(configs, "regex"));
            config.put("replace", Args.get(configs, "replace"));
            config.put("type", Args.get(configs, "type"));
            config.put("dateFormat", Args.get(configs, "dateFormat"));
            config.put("required", Args.get(configs, "required"));
            config.put("requiredBlank", Args.get(configs, "requiredBlank"));
            config.put("unifyDown", Args.get(configs, "unifyDown"));
            config.put("forceNegativeIf", Args.get(configs, "forceNegativeIf"));

            return config;
        } else {
            return null;
        }
    }

    /**
     * Converte csv para xslx
     *
     * @param csv Arquivo csv
     * @return Retorna Arquivo xslx salvo na mesma pasta do csv
     */
    public static File getCSV(File csv) {
        //Get map from csv
        List<Map<String, String>> map = CSV.getMap(csv);

        XSSFWorkbook wk = new XSSFWorkbook();
        XSSFSheet sheet = wk.createSheet("csv");

        //Coloca cabeçalho
        Row header = sheet.createRow(0);
        Integer[] headersAdd = new Integer[]{-1};
        map.get(0).forEach((col, val) -> {
            headersAdd[0]++;
            Cell cell = header.createCell(headersAdd[0]);
            cell.setCellValue(col);
        });

        //Cria linhas
        Integer[] rowsAdd = new Integer[]{0};
        map.forEach((line) -> {
            rowsAdd[0]++;
            Row row = sheet.createRow(rowsAdd[0]);

            Integer[] colsAdd = new Integer[]{-1};
            line.forEach((col, val) -> {
                colsAdd[0]++;
                Cell cell = row.createCell(colsAdd[0]);
                cell.setCellValue(val);
            });
        });

        File xlsxFile = new File(csv.getParent() + "\\" + csv.getName().toLowerCase().replaceAll(".csv", ".xlsx"));
        JExcel.saveWorkbookAs(xlsxFile, wk);
        return xlsxFile;
    }
}
