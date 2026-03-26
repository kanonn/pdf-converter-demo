package com.example;

import org.apache.pdfbox.multipdf.PDFMergerUtility;

import java.io.File;
import java.util.List;

/**
 * Merges multiple PDF files into a single PDF.
 */
public class PdfMerger {

    public static void mergePdfs(List<String> pdfPaths, String outputPath) throws Exception {
        if (pdfPaths == null || pdfPaths.isEmpty()) {
            throw new IllegalArgumentException("No PDF files specified for merging");
        }

        System.out.println("    Merging " + pdfPaths.size() + " file(s)");

        PDFMergerUtility merger = new PDFMergerUtility();

        for (String pdfPath : pdfPaths) {
            File file = new File(pdfPath);
            if (!file.exists()) {
                System.out.println("    Skip (file not found): " + pdfPath);
                continue;
            }
            merger.addSource(file);
            System.out.println("    + " + file.getName());
        }

        merger.setDestinationFileName(outputPath);
        merger.mergeDocuments(null);

        System.out.println("    Merge complete: " + outputPath);
    }
}
