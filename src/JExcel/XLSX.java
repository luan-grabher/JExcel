package JExcel;

import JExcel.JExcel;
import java.io.File;
import java.math.BigDecimal;
import java.math.BigInteger;
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
     *
     * Configurações de cada coluna:
     *
     * -name
     *
     * -collumn
     *
     * -regex: Filtro Regex
     *
     * -replace: Separa o regex do replace com § por exemplo "aa§bb" para
     * substituir todos "aa" por "bb"
     *
     * -type:Tipo de Objeto: string,value,date
     *
     * -dateFormat:Formato da data: Se tiver data
     *
     * -required: Se é Obrigatoria ou não, se for obrigatória e não tiver valor
     * ou nao for válida, não pega a linha
     *
     * -blank: Tem que estar em branco: bool
     *
     * -unifyDown: UnirColunaAbaixo: Coluna em baixo que vai ser unida no
     * resultado
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
            
            for (Row row : sheet) {
                //Pega colunas
            }
        } catch (Exception e) {
        }
        return rows;
    }

    /**
     * Retorna o objeto String/BigDecimal/Calendar conforme configuração
     *
     *
     */
    private static Object getCollumnVal(Row row, Map<String, String> colConfig) {

        try {
            if (colConfig.containsKey("collumn")) {
                Cell cel = row.getCell(JExcel.Cell(colConfig.get("collumn")));
                //Se a celula da data existir
                if (cel != null) {
                    String stringVal = JExcel.getStringCell(cel);

                    //Aplica replace se tiver replace e nao estiver em branco
                    if (colConfig.containsKey("replace") && !colConfig.get("replace").equals("")) {
                        String[] replaces = colConfig.get("replace").split("§");
                        if (replaces.length == 2) {
                            stringVal = stringVal.replaceAll(replaces[0], replaces[1]);
                        }
                    }

                    //Continua se nao tiver filtro de regex ou for match do regex
                    if (!colConfig.containsKey("regex")
                            || colConfig.get("regex").equals("")
                            || (!colConfig.get("regex").equals("")
                            && stringVal.matches(colConfig.get("regex")))) {
                        String type = colConfig.getOrDefault("type", "string");
                        if (type.equals("string")) {
                            //Se for tipo string retorna string
                            return stringVal;
                        } else if (type.equals("value")) {
                            return new BigDecimal(stringVal);
                        } else if (type.equals("date") && colConfig.containsKey("dateFormat")) {
                            return Dates.Dates.getCalendarFromFormat(stringVal, colConfig.get("dateFormat"));
                        }
                    }
                }
            }
        } catch (Exception e) {
        }
        return null;
    }

    public XLSX(File arquivo) {
        this.arquivo = arquivo;
    }

    /**
     * Adiciona os lançamentos do arquivo Excel com base nas colunas passadas.
     *
     * @param colunaData
     * @param colunaDoc
     * @param colunaPreTexto
     * @param colunaValor
     * @param colunaEntrada
     * @param colunasHistorico
     * @param colunaSaida
     */
    public void setLctos(String colunaData, String colunaDoc, String colunaPreTexto, String colunasHistorico, String colunaEntrada, String colunaSaida, String colunaValor) {
        try {
            System.out.println("Definindo workbook de " + arquivo.getName());
            wk = new XSSFWorkbook(arquivo);
            System.out.println("Definindo Sheet de " + arquivo.getName());
            sheet = wk.getSheetAt(0);

            System.out.println("Iniciando extração em " + arquivo.getName());
            setLctosFromSheet(colunaData, colunaDoc, colunaPreTexto, colunasHistorico, colunaEntrada, colunaSaida, colunaValor);
            wk.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * Adiciona os lançamentos do arquivo Excel com base nas colunas passadas. A
     * sheet e workbook ja devem estar definidos
     *
     * @param colunaData
     * @param colunaDoc
     * @param colunaPreTexto Para definir um pretexto bruto ao invés de uma
     * coluna coloque "#" na frente
     * @param colunasHistorico Coloque as colunas que compoem o historico
     * separados por ";" na ordem em que aparecem. Para configuração avançada do
     * historico separe 3 vetorescom '#', no primeiro vetor coloque a coluna do
     * excel, na segunda o prefixo (pode ficar em branco), na terceira o filtro
     * regex. o prefixo e filtro regex podem ficar em branco
     * @param colunaEntrada coluna com valores de entrada
     * @param colunaSaida coluna com valores de saida, tem que colocar "-" na
     * frente caso no excel os valores apareçam positivos
     * @param colunaValor Coluna que possui valores de entrada e saida(com sinal
     * -)
     */
    private void setLctosFromSheet(String colunaData, String colunaDoc, String colunaPreTexto, String colunasHistorico, String colunaEntrada, String colunaSaida, String colunaValor) {
        if (colunaData != null && !colunaData.isBlank()
                && colunasHistorico != null && !colunasHistorico.isBlank()) {
            //Separa as colunas de historico
            String[] colunasComplemento = colunasHistorico.split(";");
            /*
            for (Row row : sheet) {
                try {
                    Cell celData = row.getCell(JExcel.Cell(colunaData));
                    //Se a celula da data existir
                    if (celData != null) {
                        String celDateValueString = JExcel.getStringCell(celData);
                        Valor data = new Valor(celDateValueString);
                        if (data.éUmaDataValida() || (!celDateValueString.equals("") && JExcel.isDateCell(celData))) {
                            //Converte Data se for data excel
                            if (!data.éUmaDataValida()) {
                                data.setString(JExcel.getStringDate(Integer.valueOf(data.getNumbersList().get(0))));
                            }

                            String doc = "";
                            String preTexto = "";
                            String complemento = "";
                            BigDecimal value;

                            //Define o documento se tiver
                            if (colunaDoc != null && !colunaDoc.equals("")) {
                                doc = JExcel.getStringCell(row.getCell(JExcel.Cell(colunaDoc)));
                            }

                            //Define o pretexto se tiver
                            if (colunaPreTexto != null && !colunaPreTexto.equals("")) {
                                if (colunaPreTexto.contains("#")) {
                                    preTexto = colunaPreTexto.replaceAll("#", "");
                                } else {
                                    Cell cell = row.getCell(JExcel.Cell(colunaPreTexto));
                                    if (cell != null) {
                                        preTexto = JExcel.getStringCell(cell);
                                    }
                                }
                            }

                            //Define o completemento se tiver
                            if (colunasComplemento.length > 0) {
                                //Cria String Builder para fazer o Complemento
                                StringBuilder sbComplemento = new StringBuilder();
                                //Percorre todas colunas que tem
                                for (String colunaComplemento : colunasComplemento) {
                                    //Se existir uma coluna para verificar
                                    if (!colunaComplemento.equals("")) {
                                        //Divide para pegar o prefixo
                                        String[] colunaSplit = colunaComplemento.split("#");
                                        if (colunaSplit.length > 0) {
                                            String coluna = colunaSplit[0];
                                            String prefixo = colunaSplit.length > 1 ? colunaSplit[1] : "";
                                            String regex = colunaSplit.length > 2 ? colunaSplit[2] : "";

                                            //Pega celula da coluna
                                            Cell cell = row.getCell(JExcel.Cell(coluna));

                                            //Se a celula nao for nula
                                            if (cell != null) {
                                                //Pega String da celula
                                                String cellString = JExcel.getStringCell(cell);
                                                //Se nao estiver em branco e o regex estiver em branco ou a string bater com o regex
                                                if (!cellString.equals("") && ("".equals(regex) || cellString.matches(regex))) {
                                                    //Se o stringbuilder nao estiver vazio coloca - para separar
                                                    if (!sbComplemento.toString().equals("")) {
                                                        sbComplemento.append(" - ");
                                                    }
                                                    if (!prefixo.equals("")) {
                                                        sbComplemento.append(prefixo).append("- ");
                                                    }

                                                    //Adiciona a string da celula
                                                    sbComplemento.append(cellString.trim());
                                                }
                                            }
                                        }
                                    }
                                }

                                complemento = sbComplemento.toString();
                            }

                            if (colunaValor == null || colunaValor.equals("")) {
                                //Pega celulas
                                Cell entryCell = row.getCell(JExcel.Cell(colunaEntrada));
                                Cell exitCell = row.getCell(JExcel.Cell(colunaSaida.replaceAll("-", "")));

                                //Cria variavel de valores
                                BigDecimal entryBD = getBigDecimalFromCell(entryCell, false);
                                BigDecimal exitBD = getBigDecimalFromCell(exitCell, colunaSaida.contains("-"));

                                value = entryBD.compareTo(BigDecimal.ZERO) == 0 ? exitBD : entryBD;
                            } else {
                                //Pega celula
                                Cell cell = row.getCell(JExcel.Cell(colunaValor.replaceAll("-", "")));

                                value = getBigDecimalFromCell(cell, colunaValor.contains("-"));
                            }

                            //Se valor for diferente de zero
                            if (value.compareTo(BigDecimal.ZERO) != 0) {
                                lctos.add(new LctoTemplate(data.getString(), doc, preTexto, complemento, new Valor(value)));
                            }
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
             */
        } else {
            throw new Error("A coluna de extração da data e historico não podem ficar em branco!");
        }

    }

    /**
     * Pega bigdecimal de uma celula do excel numerica
     *
     * @param cell CElula que ira pegar numero
     * @param forceNegative Se deve multiplicar por -1 o numero se for positivo
     * @return celula em número BigDecimal
     */
    private BigDecimal getBigDecimalFromCell(Cell cell, boolean forceNegative) {
        //Pega texto das celulas
        String valueString = cell != null ? JExcel.getStringCell(cell) : "0.00";
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
