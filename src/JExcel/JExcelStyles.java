package JExcel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class JExcelStyles {

    public static void setCellBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
    }

    public static void removeCellBorders(CellStyle style) {
        style.setBorderTop(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);
    }

    public static void setStylesInRange(XSSFSheet sheet, CellStyle estilo,
            int primeira_linha, int ultima_linha,
            int primeira_coluna, int ultima_coluna) {
        for (int linha = primeira_linha; linha <= ultima_linha; linha++) {
            for (int coluna = primeira_coluna; coluna <= ultima_coluna; coluna++) {
                sheet.getRow(linha).getCell(coluna).setCellStyle(estilo);
            }
        }
    }
}
