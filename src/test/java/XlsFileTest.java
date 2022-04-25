import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;

public class XlsFileTest {

    private static final String FILE_NAME = "src/test/resources/sites.xlsx";

    public static SimpleDateFormat dateFormater = new SimpleDateFormat("yyyy/MM/dd");

    /**
     * Read data from cell and do some check
     */
    @Test
    void checkCellDataWithStringValue() throws IOException {
        FileInputStream fis = new FileInputStream("src/test/resources/sites.xlsx");
        Workbook xlsWorkBook = new XSSFWorkbook(fis);
        String result = xlsWorkBook.getSheetAt(0).getRow(2).getCell(3).getStringCellValue();
        doSomeCheck(result);
        fis.close();
    }

    /**
     * Read data from row and do some check
     */
    @Test
    void checkRowData() throws IOException {
        FileInputStream fis = new FileInputStream(FILE_NAME);
        Workbook xlsWorkBook = new XSSFWorkbook(fis);
        for (Cell cell : xlsWorkBook.getSheetAt(0).getRow(2)) {
            doSomeCheck(getCellData(cell));
        }
        fis.close();
    }

    /**
     * Read data from file and do some check
     */
    @Test
    void checkXlsFileData() throws IOException {
        FileInputStream fis = new FileInputStream(FILE_NAME);
        Workbook xlsWorkBook = new XSSFWorkbook(fis);
        for (Row row : xlsWorkBook.getSheetAt(0)) {
            for (Cell cell : row) {
                getCellNumber(row,cell);
                doSomeCheck(getCellData(cell));
            }
        }
        fis.close();
    }

    /**
     * Read data from cell another implementation
     */
    @Test
    void readFromExcelCell() throws IOException {
        // File for reading
        File file = new File("src/test/resources/sites.xlsx");
        // Open file for reading
        FileInputStream excelFileStream = new FileInputStream(file);
        // Read xls work book
        Workbook workBook = new XSSFWorkbook(excelFileStream);
        // Get sheet by name
        Sheet sheet = workBook.getSheet("infoSite");
        // Read row 1
        XSSFRow row = (XSSFRow) sheet.getRow(0);
        // Read cell 1
        XSSFCell cell = row.getCell(0);
        // Get cell object
        String value = cell.getStringCellValue();
        System.out.println(value);
        // Clos input stream
        excelFileStream.close();
        // Close xls work book
        workBook.close();
    }

    /**
     * Read from cells with different cell types
     */
    public static String getCellData(Cell cell) {

        String result = "";

        switch (cell.getCellType()) {
            case STRING:
                result = cell.getRichStringCellValue().getString();
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    result = dateFormater.format(cell.getDateCellValue());
                } else {
                    result = Double.toString(cell.getNumericCellValue());
                }
                break;
            case BOOLEAN:
                result = Boolean.toString(cell.getBooleanCellValue());
                break;
            case FORMULA:
                result = cell.getCellFormula();
                break;
            case BLANK:
                System.out.println();
                break;
            default:
                break;
        }
        return result;
    }

    /**
     * Some checks
     */
    void doSomeCheck(String result) {
        System.out.println(result);
    }

    /**
     * Get cell number
     */
    void getCellNumber(Row row, Cell cell){
        CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
        System.out.print(cellRef.formatAsString());
        System.out.print(" - ");
    }
}
