package JExcel;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class Sheet {

    public XSSFSheet sheet;

    public Sheet(XSSFSheet sheet) {
        this.sheet = sheet;
    }

    public XSSFCell getRange(String rangeStr) {
        if (rangeStr.length() == 2 && rangeStr.matches("[a-zA-Z][0-9]+")) {
            return getCell(Integer.valueOf(rangeStr.replaceAll("[^0-9]+", "")), JExcel.Cell(rangeStr.replaceAll("[0-9]+", "")));
        } else {
            return getCell(0, 0);
        }
    }

    /**
     * Cria ou retorna a celula
     *
     * @param row numero da linha
     * @param cell numero da celula
     * @return retorna celula se existente ou cria uma nova
     */
    public XSSFCell getCell(int row, int cell) {
        XSSFRow rowXSSF = getRow(row);

        return rowXSSF.getCell(cell) != null ? rowXSSF.getCell(cell) : rowXSSF.createCell(cell);
    }

    /**
     * Cria ou retorna a linha
     *
     * @param row numero da linha
     * @return retorna a linha se existir ou cria uma nova
     */
    public XSSFRow getRow(int row) {
        return sheet.getRow(row) != null ? sheet.getRow(row) : sheet.createRow(row);
    }
}
