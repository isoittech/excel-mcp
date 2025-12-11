package msi.aiproc.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.List;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import com.google.gson.Gson;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class UpdateExcelFile {
    // クラスロガー
    private static Logger logger = null;

    public static boolean fileExists(String filePath) {
        File file = new File(filePath);
        return file.exists();
    }

    public static void main(String[] args) {
        // 引数を確認する
        if (args.length != 4) {
            System.err.println("引数の数が不適切: " + args.length);
            System.exit(1);
        }

        String loggerXmlFilePath = args[0];
        String inputExcelFilePath = args[1];
        String jsonFilePath = args[2];
        String outputExcelFilePath = args[3];

        if (!fileExists(loggerXmlFilePath)) {
            logger.error("ログの設定ファイルが参照できない: " + loggerXmlFilePath);
            System.exit(1);
        }

        // ロガーを用意する
        System.setProperty("log4j2.configurationFile", loggerXmlFilePath);
        logger = LogManager.getLogger(UpdateExcelFile.class);
        logger.info("[START PROCESSING]");

        logger.info("loggerXmlFilePath: " + loggerXmlFilePath);
        logger.info("inputExcelFilePath: " + inputExcelFilePath);
        logger.info("jsonFilePath: " + jsonFilePath);
        logger.info("outputExcelFilePath: " + outputExcelFilePath);

        // 処理対象ファイルを確認する
        if (!fileExists(inputExcelFilePath)) {
            logger.error("ファイルが存在しない: " + inputExcelFilePath);
            System.exit(1);
        }
        if (!fileExists(jsonFilePath)) {
            logger.error("ファイルが存在しない: " + jsonFilePath);
            System.exit(1);
        }
        
        // jsonファイルを読み込む
        JsonData jsonData = ReadJsonFile.read(jsonFilePath);
        
        try (FileInputStream fis = new FileInputStream(new File(inputExcelFilePath));
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(new File(outputExcelFilePath))) {

            // シートを取得
            updateSheet(workbook, "表紙", jsonData.getCoverPage());
            updateSheet(workbook, "明細", jsonData.getDetails());
            updateSheet(workbook, "表紙 (バイヤ)", jsonData.getCoverPageBuyer());
            updateSheet(workbook, "明細 (バイヤ)", jsonData.getDetailsBuyer());
            // 変更内容を新しいファイルに書き込む
            workbook.write(fos);
        }
        catch (FileNotFoundException e) {
            logger.error("指定された入力ファイルがない: " + inputExcelFilePath, e);
            System.exit(1);
        }
        catch (IOException e) {
            logger.error("ファイルの読み書きエラー", e);
            System.exit(1);
        }
        catch (Exception e) {
            logger.error("処理失敗", e);
            System.exit(1);
        }

        // 正常終了
        logger.info("[SUCCESSFUL TERMINATION]");
        System.exit(0);
    }

    private static void updateSheet(Workbook workbook,
                                    String sheetName,
                                    List<CellInfo> cellInfoList) {
        // デバッグ表示用
        //printSection(sheetName, cellInfoList);
        logger.info("sheetName: " + sheetName);
        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            logger.error("指定されたシートがない: " + sheetName);
            System.exit(1);
        }
        
        if (cellInfoList != null && !cellInfoList.isEmpty()) {
            for (CellInfo cellInfo : cellInfoList) {
                // 行を特定する
                Row row = sheet.getRow(cellInfo.getRow());
                if (row == null) {
                    logger.error("指定された行が存在しない: " + cellInfo.getRow());
                    System.exit(1);
                }

                // セルを特定する
                Cell cell = row.getCell(cellInfo.getColumn());
                if (cell == null) {
                    logger.error("指定されたセルが存在しない: " + cellInfo.getRow() + "行, " + cellInfo.getColumn() + "列");
                    System.exit(1);
                }

                switch (cellInfo.getType()) {
                case "string":
                    logger.info("setStringValue row:" + cellInfo.getRow() + " column:" + cellInfo.getColumn() + " value:" + cellInfo.getStringValue());
                    cell.setCellValue(cellInfo.getStringValue());
                    break;
                case "integer":
                    logger.info("setIntegerValue row:" + cellInfo.getRow() + " column:" + cellInfo.getColumn() + " value:" + cellInfo.getIntValue());
                    cell.setCellValue(cellInfo.getIntValue());
                    break;
                case "double":
                    logger.info("setDoubleValue row:" + cellInfo.getRow() + " column:" + cellInfo.getColumn() + " value:" + cellInfo.getDoubleValue());
                    cell.setCellValue(cellInfo.getDoubleValue());
                    break;
                }
            }
        }
    }

    // 各セクションのデータを表示するヘルパーメソッド
    private static void printSection(String sectionName, List<CellInfo> cellInfoList) {
        System.out.println("--- " + sectionName + " ---");
        if (cellInfoList != null && !cellInfoList.isEmpty()) {
            for (CellInfo cellInfo : cellInfoList) {
                System.out.println(cellInfo);
            }
        } else {
            System.out.println("データなし");
        }
    }
}
