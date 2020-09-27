package io.avec.excel;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;

public class ExcelReader {

    private static final String file = "salary.xlsx";

    public static void main(String[] args) throws IOException {
        Workbook workbook = WorkbookFactory.create(new FileInputStream(file));

        System.out.println("Workbook " + file + " has " + workbook.getNumberOfSheets() + " sheets.");

        // find all sheet names
        System.out.println("Retrieving Sheets using for-each loop");
        for(Sheet sheet : workbook) {
            System.out.println("=> " + sheet.getSheetName());
        }


        // find latest (newest) sheet
        int sheetCount = workbook.getNumberOfSheets();
        Sheet lastSheet = workbook.getSheetAt(sheetCount-1); // First sheet is 0, last is count-1
        System.out.println("Last sheet is " + lastSheet.getSheetName());

        // last sheet content
        for(Row row : lastSheet) {
            for(Cell cell : row) {
                printCellValue(cell);
            }
            System.out.println();
        }
    }

    // print cell content as correct type
    private static void printCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case BOOLEAN:
                System.out.print(cell.getBooleanCellValue());
                break;
            case STRING:
                System.out.print(cell.getRichStringCellValue().getString());
                break;
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    System.out.print(cell.getDateCellValue());
                } else {
                    System.out.print(cell.getNumericCellValue());
                }
                break;
            case FORMULA:
                System.out.print(cell.getCellFormula());
                break;
            default:
                System.out.print("");
        }

        System.out.print("\t");
    }
}
