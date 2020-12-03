package JExcel;

import fileManager.FileManager;
import java.io.File;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.util.Calendar;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class JExcel {

    public static boolean saveWorkbookAs(File saveFile, Workbook workbook) {
        try {
            if (!saveFile.isDirectory()) {
                //Cria novo arquivo editado
                FileOutputStream outFile = new FileOutputStream(saveFile.getAbsolutePath());
                workbook.write(outFile);
                outFile.close();
                return true;
            } else {
                throw new Exception("Para salvar o workbook, o file passado deve ser um arquivo!");
            }
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }

    }

    public static boolean saveSheetAsCsv(File newFile, File arquivoExcelOriginal, Sheet sheet) {
        try {
            String textCSV = JExcel.sheetToCSV(sheet);
            return FileManager.save(newFile.getAbsolutePath(), textCSV);
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }

    }

    public static String sheetToCSV(Sheet sheet) {
        StringBuilder builder = new StringBuilder();

        try {
            Row row = null;
            for (int i = 0; i <= sheet.getLastRowNum(); i++) {
                try {
                    row = sheet.getRow(i);
                    short lastCol = row.getLastCellNum();
                    lastCol--;

                    for (int j = 0; j <= lastCol; j++) {
                        builder.append(row.getCell(j) + (j != lastCol ? ";" : ""));
                    }
                    if (i != sheet.getLastRowNum()) {
                        builder.append("\r\n");
                    }
                } catch (Exception e) {
                }
            }
        } catch (Exception e) {
            System.out.println("Erro ao converter Sheet para CSV: " + e);
            e.printStackTrace();
        }

        return builder.toString();
    }

    public static String getStringCell(Cell cel) {
        try {
            if(cel != null){
                String tipo = cel.getCellType().name();

                switch (tipo) {
                    case "STRING":
                        return cel.getStringCellValue();
                    case "NUMERIC":
                        return String.valueOf(cel.getNumericCellValue());
                    default:
                        return "";
                }
            }else{
                throw new Exception("Célula inexistente ou sem nada");
            }
        } catch (Exception e) {
            e.printStackTrace();
            return "";
        }

    }
    
        /**
     * Pega a String da célula
     * @param cel
     * @return String da célula
     */
    public static String getCellString(Cell cel) {
        try {
            CellType type = cel.getCellType();

            switch (type.name()) {
                case "STRING":
                    return cel.getStringCellValue();
                case "NUMERIC":
                    return String.valueOf(new BigDecimal(cel.getNumericCellValue()).toPlainString());
                default:
                    return "";
            }
        } catch (Exception e) {
            return "";
        }
    }

    public static String getStringDate(int daysAfter1900) {
        Calendar cal = Calendar.getInstance();
        cal.set(1900, 0, 1);
        cal.add(Calendar.DAY_OF_MONTH, (daysAfter1900 - 2));

        return cal.get(Calendar.DAY_OF_MONTH) + "/" + (cal.get(Calendar.MONTH) + 1) + "/" + (cal.get(Calendar.YEAR));
    }

    public static void removeRows(Sheet sheet, int firstRow, int lastRow) {
        try {
            for (int i = lastRow; i >= firstRow; i--) {
                if (i <= sheet.getLastRowNum()) {
                    Row row = sheet.getRow(i);
                    // sheet.removeRow(row);    NO NEED FOR THIS LINE
                    sheet.shiftRows(row.getRowNum() + 1, sheet.getLastRowNum() + 1, -1);
                } else if (sheet.getLastRowNum() >= firstRow) {
                    i = sheet.getLastRowNum() + 1;
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static int Cell(String letter) {
        int start = 97;
        int end = 122;

        String collumStr = letter.toLowerCase();
        try {
            int charAt0 = collumStr.charAt(0);
            if (charAt0 >= start && charAt0 <= end) {
                return charAt0 - start;
            } else {
                return 0;
            }
        } catch (Exception e) {
            return 0;
        }
    }

    public static boolean isDateCell(Cell cell) {
        try {
            if (DateUtil.isCellDateFormatted(cell)) {
                return true;
            } else {
                return false;
            }
        } catch (Exception e) {
            return false;
        }
    }
    
    public void setCellBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
    }
    
    
    /**
     * Pega bigdecimal de uma celula do excel numerica
     * 
     * @param cell CElula que ira pegar numero
     * @param forceNegative Se deve multiplicar por -1 o numero se for positivo
     * @return celula em número BigDecimal
     */
    public static BigDecimal getBigDecimalFromCell(Cell cell, boolean forceNegative) {
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
