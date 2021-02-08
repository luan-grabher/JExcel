package JExcel;

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
     * Retorna uma lista das linhas de um arquivo xlsx com um mapa com o valores
     * das colunas de cada linha conforme as configurações definidas.
     * <p>
     * Para procurar um regex no meio de algo e ignorar case:
     * "(?i).*" + search + ".*"<br>
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
            XSSFWorkbook wk;
            XSSFSheet sheet;

            wk = new XSSFWorkbook(file);
            sheet = wk.getSheetAt(0);
            
            //Remove todas config NULL
            while (config.values().remove(null));
            config.forEach((n, c)->{
                //Remove todas colunas NULL
                while (c.values().remove(null));
            });            

            for (Row row : sheet) {
                //Cria mapa de colunas
                Map<String, Object> cols = new HashMap<>();
                Boolean[] rowValid = new Boolean[]{true};

                //Percorre todas colunas das configurações
                config.forEach((name, col) -> {
                    if (col != null) {

                        //Pega o objeto da coluna
                        Object colObj = getCollumnVal(row, col);

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
                });

                if (rowValid[0] == true && cols.size() > 0) {
                    rows.add(cols);
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
                if (col.length() == 2 && col.startsWith("-")) {
                    col = col.replaceAll("-", "");
                }

                Cell cel = row.getCell(JExcel.Cell(col));
                if (cel != null) {
                    if (!result.toString().equals("")) {
                        result.append(" ");
                    }
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
        valueString = valueString.replaceAll("[^0-9\\.,-]", "");

        //Se tiver . antes da virgula remove os pontos e coloca ponto no lugar da virgula
        if (valueString.indexOf(".") < valueString.indexOf(",")) {
            valueString = valueString.replaceAll("\\.", "").replaceAll("\\,", ".");
        }

        BigDecimal valueBigDecimal = new BigDecimal(valueString.equals("") ? "0" : valueString);

        //Se a coluna tiver que multiplicar por -1 e o valor encontrado for maior que zero
        if (forceNegative && valueBigDecimal.compareTo(BigDecimal.ZERO) > 0) {
            valueBigDecimal = valueBigDecimal.multiply(new BigDecimal("-1"));
        }

        return valueBigDecimal;
    }

}
