package com.example;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * PDF Converter Main Class
 *
 * Downloads files from S3, converts Excel/Word/TIFF to PDF,
 * merges all PDFs into one, and uploads the result back to S3.
 */
public class PdfConverterMain {

    // Local working directories
    private static final String INPUT_DIR = "input";
    private static final String OUTPUT_DIR = "output";
    private static final String TEMP_DIR = "temp";

    // Configuration - can be overridden by environment variables or command line args
    private static String s3BucketName = "your-bucket-name";
    private static String s3Region = "ap-northeast-1";
    private static String s3InputPrefix = "input/";
    private static String s3OutputPrefix = "output/";
    private static String targetExcelSheet = "Sheet1";

    public static void main(String[] args) {
        System.out.println("========================================");
        System.out.println("  PDF Converter - Start");
        System.out.println("========================================");

        // Parse command line arguments
        parseArguments(args);

        S3ClientUtil s3Client = null;

        try {
            // Setup directories
            setupDirectories();

            // Initialize S3 client
            s3Client = new S3ClientUtil(s3BucketName, s3Region);

            // Step 1: Download files from S3
            System.out.println("\n[1/5] Downloading files from S3...");
            downloadFilesFromS3(s3Client);

            List<String> pdfFiles = new ArrayList<>();

            // Step 2: Convert Excel to PDF (specific sheet only)
            System.out.println("\n[2/5] Converting Excel to PDF...");
            String excelPdf = convertExcel();
            if (excelPdf != null) {
                pdfFiles.add(excelPdf);
                System.out.println("  OK: " + excelPdf);
            }

            // Step 3: Convert Word to PDF
            System.out.println("\n[3/5] Converting Word to PDF...");
            String wordPdf = convertWord();
            if (wordPdf != null) {
                pdfFiles.add(wordPdf);
                System.out.println("  OK: " + wordPdf);
            }

            // Step 4: Convert TIFF to PDF (with rotation if needed)
            System.out.println("\n[4/5] Converting TIFF to PDF...");
            List<String> tiffPdfs = convertTiffs();
            pdfFiles.addAll(tiffPdfs);
            for (String pdf : tiffPdfs) {
                System.out.println("  OK: " + pdf);
            }

            // Step 5: Merge all PDFs
            System.out.println("\n[5/5] Merging PDFs...");
            if (pdfFiles.isEmpty()) {
                System.out.println("  No PDF files to merge");
            } else {
                String mergedPdf = OUTPUT_DIR + "/merged_document.pdf";
                PdfMerger.mergePdfs(pdfFiles, mergedPdf);
                System.out.println("  OK: " + mergedPdf);

                // Upload merged PDF to S3
                System.out.println("\n[6/6] Uploading result to S3...");
                uploadResultToS3(s3Client, mergedPdf);
            }

            System.out.println("\n========================================");
            System.out.println("  PDF Converter - Complete");
            System.out.println("========================================");
            System.out.println("\nOutput files:");
            listOutputFiles();

        } catch (Exception e) {
            System.err.println("\nError: " + e.getMessage());
            e.printStackTrace();
        } finally {
            if (s3Client != null) {
                s3Client.close();
            }
        }
    }

    /**
     * Parse command line arguments.
     */
    private static void parseArguments(String[] args) {
        for (int i = 0; i < args.length; i++) {
            switch (args[i]) {
                case "--bucket":
                    if (i + 1 < args.length) s3BucketName = args[++i];
                    break;
                case "--region":
                    if (i + 1 < args.length) s3Region = args[++i];
                    break;
                case "--input-prefix":
                    if (i + 1 < args.length) s3InputPrefix = args[++i];
                    break;
                case "--output-prefix":
                    if (i + 1 < args.length) s3OutputPrefix = args[++i];
                    break;
                case "--sheet":
                    if (i + 1 < args.length) targetExcelSheet = args[++i];
                    break;
                case "--help":
                    printUsage();
                    System.exit(0);
                    break;
            }
        }

        // Also check environment variables
        String envBucket = System.getenv("S3_BUCKET_NAME");
        if (envBucket != null && !envBucket.isEmpty()) {
            s3BucketName = envBucket;
        }

        String envRegion = System.getenv("AWS_REGION");
        if (envRegion != null && !envRegion.isEmpty()) {
            s3Region = envRegion;
        }

        String envInputPrefix = System.getenv("S3_INPUT_PREFIX");
        if (envInputPrefix != null && !envInputPrefix.isEmpty()) {
            s3InputPrefix = envInputPrefix;
        }

        String envOutputPrefix = System.getenv("S3_OUTPUT_PREFIX");
        if (envOutputPrefix != null && !envOutputPrefix.isEmpty()) {
            s3OutputPrefix = envOutputPrefix;
        }

        String envSheet = System.getenv("TARGET_EXCEL_SHEET");
        if (envSheet != null && !envSheet.isEmpty()) {
            targetExcelSheet = envSheet;
        }

        System.out.println("Configuration:");
        System.out.println("  Bucket: " + s3BucketName);
        System.out.println("  Region: " + s3Region);
        System.out.println("  Input prefix: " + s3InputPrefix);
        System.out.println("  Output prefix: " + s3OutputPrefix);
        System.out.println("  Target sheet: " + targetExcelSheet);
    }

    /**
     * Print usage information.
     */
    private static void printUsage() {
        System.out.println("Usage: java -jar pdf-converter-s3.jar [options]");
        System.out.println();
        System.out.println("Options:");
        System.out.println("  --bucket <name>        S3 bucket name");
        System.out.println("  --region <region>      AWS region (default: ap-northeast-1)");
        System.out.println("  --input-prefix <path>  S3 input prefix (default: input/)");
        System.out.println("  --output-prefix <path> S3 output prefix (default: output/)");
        System.out.println("  --sheet <name>         Excel sheet name to convert (default: Sheet1)");
        System.out.println("  --help                 Show this help message");
        System.out.println();
        System.out.println("Environment variables:");
        System.out.println("  S3_BUCKET_NAME         S3 bucket name");
        System.out.println("  AWS_REGION             AWS region");
        System.out.println("  S3_INPUT_PREFIX        S3 input prefix");
        System.out.println("  S3_OUTPUT_PREFIX       S3 output prefix");
        System.out.println("  TARGET_EXCEL_SHEET     Excel sheet name to convert");
    }

    /**
     * Setup local working directories.
     */
    private static void setupDirectories() throws Exception {
        Files.createDirectories(Paths.get(INPUT_DIR));
        Files.createDirectories(Paths.get(OUTPUT_DIR));
        Files.createDirectories(Paths.get(TEMP_DIR));
    }

    /**
     * Download files from S3 input prefix to local input directory.
     */
    private static void downloadFilesFromS3(S3ClientUtil s3Client) throws Exception {
        List<String> s3Keys = s3Client.listObjects(s3InputPrefix);

        if (s3Keys.isEmpty()) {
            System.out.println("  No files found in s3://" + s3BucketName + "/" + s3InputPrefix);
            return;
        }

        System.out.println("  Found " + s3Keys.size() + " file(s)");

        for (String s3Key : s3Keys) {
            // Extract filename from S3 key
            String fileName = s3Key.substring(s3Key.lastIndexOf('/') + 1);
            Path localPath = Paths.get(INPUT_DIR, fileName);

            s3Client.downloadFile(s3Key, localPath);
        }
    }

    /**
     * Upload result PDF to S3.
     */
    private static void uploadResultToS3(S3ClientUtil s3Client, String localPdfPath) {
        Path localPath = Paths.get(localPdfPath);
        String s3Key = s3OutputPrefix + localPath.getFileName().toString();

        s3Client.uploadFile(localPath, s3Key, "application/pdf");
    }

    /**
     * Convert Excel file to PDF (specific sheet only).
     */
    private static String convertExcel() throws Exception {
        File inputDir = new File(INPUT_DIR);
        File[] excelFiles = inputDir.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".xlsx") || name.toLowerCase().endsWith(".xls"));

        if (excelFiles == null || excelFiles.length == 0) {
            System.out.println("  No Excel files found");
            return null;
        }

        // Process first Excel file found
        File excelFile = excelFiles[0];
        String tempExcel = TEMP_DIR + "/excel_single_sheet.xlsx";
        String outputPdf = OUTPUT_DIR + "/excel_output.pdf";

        // Extract specific sheet only
        ExcelSheetExtractor.extractSheet(excelFile.getPath(), targetExcelSheet, tempExcel);

        // Convert to PDF using LibreOffice
        LibreOfficeConverter.convertToPdf(tempExcel, OUTPUT_DIR);

        // Rename output file
        File generatedPdf = new File(OUTPUT_DIR + "/excel_single_sheet.pdf");
        if (generatedPdf.exists()) {
            generatedPdf.renameTo(new File(outputPdf));
        }

        return outputPdf;
    }

    /**
     * Convert Word file to PDF.
     */
    private static String convertWord() throws Exception {
        File inputDir = new File(INPUT_DIR);
        File[] wordFiles = inputDir.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".docx") || name.toLowerCase().endsWith(".doc"));

        if (wordFiles == null || wordFiles.length == 0) {
            System.out.println("  No Word files found");
            return null;
        }

        // Process first Word file found
        File wordFile = wordFiles[0];
        String outputPdf = OUTPUT_DIR + "/word_output.pdf";

        // Convert to PDF using LibreOffice
        LibreOfficeConverter.convertToPdf(wordFile.getPath(), OUTPUT_DIR);

        // Rename output file
        String baseName = wordFile.getName().replaceAll("\\.[^.]+$", "");
        File generatedPdf = new File(OUTPUT_DIR + "/" + baseName + ".pdf");
        if (generatedPdf.exists()) {
            generatedPdf.renameTo(new File(outputPdf));
        }

        return outputPdf;
    }

    /**
     * Convert TIFF files to PDF (rotate portrait images 90 degrees).
     */
    private static List<String> convertTiffs() throws Exception {
        List<String> pdfFiles = new ArrayList<>();
        File inputDir = new File(INPUT_DIR);

        File[] tiffFiles = inputDir.listFiles((dir, name) ->
                name.toLowerCase().endsWith(".tif") || name.toLowerCase().endsWith(".tiff"));

        if (tiffFiles == null || tiffFiles.length == 0) {
            System.out.println("  No TIFF files found");
            return pdfFiles;
        }

        for (File tiffFile : tiffFiles) {
            String baseName = tiffFile.getName().replaceAll("\\.(tif|tiff)$", "");
            String outputPdf = OUTPUT_DIR + "/" + baseName + ".pdf";

            TiffToPdfConverter.convert(tiffFile.getPath(), outputPdf);
            pdfFiles.add(outputPdf);
        }

        return pdfFiles;
    }

    /**
     * List output files.
     */
    private static void listOutputFiles() {
        File outputDir = new File(OUTPUT_DIR);
        File[] files = outputDir.listFiles();
        if (files != null) {
            for (File file : files) {
                long sizeKB = file.length() / 1024;
                System.out.println("  - " + file.getName() + " (" + sizeKB + " KB)");
            }
        }
    }
}
