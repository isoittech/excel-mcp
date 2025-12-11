/**
 * Command line tool that copies a rectangular range of cells to another
 * location, possibly on another sheet. Implements "copy_range" from
 * README.JA.md.
 */
package jp.isoittech;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CopyRangeTool {

    /**
     * Arguments:
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>sourceStart</li>
     *     <li>sourceEnd</li>
     *     <li>targetStart</li>
     *     <li>targetSheet (optional, when omitted the same sheet is used)</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length < 5 || args.length > 6) {
            System.err.println("Usage: CopyRangeTool <filePath> <sheetName> <sourceStart> <sourceEnd> <targetStart> [targetSheet]");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String sourceStart = args[2];
        String sourceEnd = args[3];
        String targetStart = args[4];
        String targetSheetName = args.length == 6 ? args[5] : sheetName;

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sourceSheet = workbook.getSheet(sheetName);
            if (sourceSheet == null) {
                throw new IllegalArgumentException("Source sheet not found: " + sheetName);
            }

            Sheet targetSheet = ExcelUtils.getOrCreateSheet(workbook, targetSheetName);

            CellRangeAddress sourceRange = ExcelRangeUtils.parseRange(sourceStart + ":" + sourceEnd);
            CellAddress targetStartAddr = ExcelRangeUtils.parseCellAddress(targetStart);

            int rowOffset = targetStartAddr.getRow() - sourceRange.getFirstRow();
            int colOffset = targetStartAddr.getColumn() - sourceRange.getFirstColumn();

            for (int r = sourceRange.getFirstRow(); r <= sourceRange.getLastRow(); r++) {
                Row srcRow = sourceSheet.getRow(r);
                Row tgtRow = targetSheet.getRow(r + rowOffset);
                if (tgtRow == null) {
                    tgtRow = targetSheet.createRow(r + rowOffset);
                }
                for (int c = sourceRange.getFirstColumn(); c <= sourceRange.getLastColumn(); c++) {
                    Cell srcCell = srcRow != null ? srcRow.getCell(c) : null;
                    Cell tgtCell = tgtRow.getCell(c + colOffset);
                    if (tgtCell == null) {
                        tgtCell = tgtRow.createCell(c + colOffset);
                    }
                    if (srcCell != null) {
                        copyCellValue(srcCell, tgtCell);
                    } else {
                        tgtCell.setBlank();
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * Copies the value and basic type from one cell to another.
     */
    private static void copyCellValue(Cell src, Cell dest) {
        switch (src.getCellType()) {
            case STRING:
                dest.setCellValue(src.getStringCellValue());
                break;
            case NUMERIC:
                dest.setCellValue(src.getNumericCellValue());
                break;
            case BOOLEAN:
                dest.setCellValue(src.getBooleanCellValue());
                break;
            case FORMULA:
                dest.setCellFormula(src.getCellFormula());
                break;
            default:
                dest.setBlank();
        }
    }
}
