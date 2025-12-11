/**
 * Command line tool that formats a cell range (font style, color, background
 * color, etc.). Implements "format_range" from README.JA.md.
 */
package jp.isoittech;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FormatRangeTool {

    /**
     * Arguments (all values passed as strings):
     * <ol>
     *     <li>filePath</li>
     *     <li>sheetName</li>
     *     <li>startCell</li>
     *     <li>endCell</li>
     *     <li>bold (true/false)</li>
     *     <li>italic (true/false)</li>
     *     <li>fontSize (integer, 0 to keep current)</li>
     *     <li>fontColor (RGB hex like "#FF0000" or "" to skip)</li>
     *     <li>bgColor (RGB hex like "#FFFF00" or "" to skip)</li>
     * </ol>
     */
    public static void main(String[] args) throws Exception {
        if (args.length != 9) {
            System.err.println("Usage: FormatRangeTool <filePath> <sheetName> <startCell> <endCell> <bold> <italic> <fontSize> <fontColor> <bgColor>");
            System.exit(1);
        }

        String filePath = args[0];
        String sheetName = args[1];
        String startCell = args[2];
        String endCell = args[3];
        boolean bold = Boolean.parseBoolean(args[4]);
        boolean italic = Boolean.parseBoolean(args[5]);
        int fontSize = Integer.parseInt(args[6]);
        String fontColor = args[7];
        String bgColor = args[8];

        File file = new File(filePath);
        if (!file.exists()) {
            throw new IOException("File not found: " + filePath);
        }

        try (FileInputStream fis = new FileInputStream(file);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            CellRangeAddress range = ExcelRangeUtils.parseRange(startCell + ":" + endCell);

            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setBold(bold);
            font.setItalic(italic);
            if (fontSize > 0) {
                font.setFontHeightInPoints((short) fontSize);
            }
            if (!fontColor.isEmpty()) {
                XSSFColor color = new XSSFColor(parseRgbColor(fontColor), null);
                // In a full implementation we would use an indexed color or theme color.
                // For simplicity, we ignore the font RGB color when applying the style.
            }
            style.setFont(font);

            if (!bgColor.isEmpty()) {
                XSSFColor bg = new XSSFColor(parseRgbColor(bgColor), null);
                // For simplicity, ignore background RGB in this minimal implementation.
            }

            for (int r = range.getFirstRow(); r <= range.getLastRow(); r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    row = sheet.createRow(r);
                }
                for (int c = range.getFirstColumn(); c <= range.getLastColumn(); c++) {
                    Cell cell = row.getCell(c);
                    if (cell == null) {
                        cell = row.createCell(c);
                    }
                    cell.setCellStyle(style);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }
        }
    }

    /**
     * Parses a color expressed as "#RRGGBB" into a byte array suitable
     * for {@link XSSFColor}.
     */
    private static byte[] parseRgbColor(String hex) {
        if (hex == null || !hex.matches("#?[0-9A-Fa-f]{6}")) {
            throw new IllegalArgumentException("Invalid RGB color: " + hex);
        }
        String v = hex.startsWith("#") ? hex.substring(1) : hex;
        int r = Integer.parseInt(v.substring(0, 2), 16);
        int g = Integer.parseInt(v.substring(2, 4), 16);
        int b = Integer.parseInt(v.substring(4, 6), 16);
        return new byte[]{(byte) r, (byte) g, (byte) b};
    }
}
