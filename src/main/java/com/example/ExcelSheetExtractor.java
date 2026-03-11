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
 *
 * Method: Copy the original file and delete unwanted sheets.
 * This preserves all formatting including column widths, row heights,
 * shapes, images, and styles.
 */
public class ExcelSheetExtractor {

    /**
     * Extract a specific sheet from Excel file.
     * Creates a new file containing only the specified sheet,
     * preserving all original formatting.
     *
     * @param inputPath  Input Excel file path
     * @param sheetName  Name of the sheet to extract
     * @param outputPath Output Excel file path
     * @throws Exception if extraction fails
     */
    public static void extractSheet(String inputPath, String sheetName, String outputPath) throws Exception {
        System.out.println("    Target sheet: " + sheetName);

        // Copy original file
        File inputFile = new File(inputPath);
        File outputFile = new File(outputPath);
        Files.copy(inputFile.toPath(), outputFile.toPath(), StandardCopyOption.REPLACE_EXISTING);

        // Open copied file and delete unwanted sheets
        try (FileInputStream fis = new FileInputStream(outputFile);
             Workbook workbook = new XSSFWorkbook(fis)) {

            // Find target sheet
            int targetIndex = workbook.getSheetIndex(sheetName);
            if (targetIndex == -1) {
                // Show available sheet names
                System.out.println("    ERROR: Sheet not found: " + sheetName);
                System.out.println("    Available sheets:");
                for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                    System.out.println("      - " + workbook.getSheetName(i));
                }
                outputFile.delete();
                throw new IllegalArgumentException("Sheet not found: " + sheetName);
            }

            // Collect indices of sheets to remove (in reverse order)
            List<Integer> sheetsToRemove = new ArrayList<>();
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                if (i != targetIndex) {
                    sheetsToRemove.add(i);
                }
            }

            // Remove sheets from end to avoid index shifting
            for (int i = sheetsToRemove.size() - 1; i >= 0; i--) {
                workbook.removeSheetAt(sheetsToRemove.get(i));
            }

            // Set remaining sheet as active
            workbook.setActiveSheet(0);

            // Save
            try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                workbook.write(fos);
            }

            System.out.println("    Sheet extracted (format preserved): " + outputPath);
        }
    }
}
