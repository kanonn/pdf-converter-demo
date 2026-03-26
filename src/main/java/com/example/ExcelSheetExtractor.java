package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

/**
 * Extracts a specific sheet from an Excel file.
 */
public class ExcelSheetExtractor {

    public static void extractSheet(String inputPath, String sheetName, String outputPath) throws Exception {
        System.out.println("    Target sheet: " + sheetName);

        File inputFile = new File(inputPath);
        File outputFile = new File(outputPath);
        Files.copy(inputFile.toPath(), outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

        try (FileInputStream fis = new FileInputStream(outputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            int targetIndex = workbook.getSheetIndex(sheetName);
            if (targetIndex == -1) {
                System.out.println("    ERROR: Sheet not found: " + sheetName);
                System.out.println("    Available sheets:");
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    System.out.println("      - " + workbook.getSheetName(i));
                }
                outputFile.delete();
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            List<Integer> sheetsToRemove = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (i != targetIndex) {
                    sheetsToRemove.add(i);
                }
            }

            for (int i = sheetsToRemove.size() - 1; i >= 0; i--) {
                workbook.removeSheetAt(sheetsToRemove.get(i));
            }

            workbook.setActiveSheet(0);

            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            System.out.println("    Sheet extracted: " + outputPath);
        }
    }
}
