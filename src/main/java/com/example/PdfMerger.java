package com.example;

import org.apache.pdfbox.multipdf.PDFMergerUtility;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDDocumentInformation;

import java.io.File;
import java.util.Calendar;
import java.util.List;

/**
 * Merges multiple PDF files into a single PDF.
 */
public class PdfMerger {

    /**
     * Merge multiple PDF files into one.
     *
     * @param pdfPaths   List of PDF file paths to merge
     * @param outputPath Output PDF file path
     * @throws Exception if merge fails
     */
    public static void mergePdfs(List<String> pdfPaths, String outputPath) throws Exception {
        if (pdfPaths == null || pdfPaths.isEmpty()) {
            throw new IllegalArgumentException("No PDF files specified for merging");
        }

        System.out.println("    Merging " + pdfPaths.size() + " file(s)");

        PDFMergerUtility merger = new PDFMergerUtility();

        // Add files to merge
        for (String pdfPath : pdfPaths) {
            File file = new File(pdfPath);
            if (!file.exists()) {
                System.out.println("    Skip (file not found): " + pdfPath);
                continue;
            }
            merger.addSource(file);
            System.out.println("    + " + file.getName());
        }

        // Set output destination
        merger.setDestinationFileName(outputPath);

        // Execute merge
        merger.mergeDocuments(null);

        // Add metadata (optional)
        addMetadata(outputPath);

        System.out.println("    Merge complete: " + outputPath);
    }

    /**
     * Add metadata to PDF.
     */
    private static void addMetadata(String pdfPath) {
        try (PDDocument document = PDDocument.load(new File(pdfPath))) {
            PDDocumentInformation info = document.getDocumentInformation();
            info.setTitle("Merged Document");
            info.setAuthor("PDF Converter");
            info.setSubject("Merged PDF from Excel, Word, and TIFF files");
            info.setCreator("pdf-converter-s3");
            info.setCreationDate(Calendar.getInstance());
            info.setModificationDate(Calendar.getInstance());

            document.save(pdfPath);
        } catch (Exception e) {
            // Skip metadata addition on failure
            System.out.println("    Note: Metadata addition skipped: " + e.getMessage());
        }
    }
}
