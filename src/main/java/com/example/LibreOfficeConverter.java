package com.example;

import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.util.concurrent.TimeUnit;

/**
 * Converts Excel/Word files to PDF using LibreOffice headless mode.
 */
public class LibreOfficeConverter {

    private static final int TIMEOUT_SECONDS = 60;

    /**
     * Convert a file to PDF format.
     *
     * @param inputFile Input file path (Excel or Word)
     * @param outputDir Output directory path
     * @throws Exception if conversion fails
     */
    public static void convertToPdf(String inputFile, String outputDir) throws Exception {
        File input = new File(inputFile);
        if (!input.exists()) {
            throw new IllegalArgumentException("Input file does not exist: " + inputFile);
        }

        File outDir = new File(outputDir);
        if (!outDir.exists()) {
            outDir.mkdirs();
        }

        // Build LibreOffice command
        ProcessBuilder pb = new ProcessBuilder(
                "libreoffice",
                "--headless",
                "--invisible",
                "--nologo",
                "--nofirststartwizard",
                "--convert-to", "pdf",
                "--outdir", outDir.getAbsolutePath(),
                input.getAbsolutePath()
        );

        pb.redirectErrorStream(true);

        System.out.println("    Executing: libreoffice --headless --convert-to pdf ...");

        Process process = pb.start();

        // Read output for debugging
        StringBuilder output = new StringBuilder();
        try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
            String line;
            while ((line = reader.readLine()) != null) {
                output.append(line).append("\n");
            }
        }

        boolean finished = process.waitFor(TIMEOUT_SECONDS, TimeUnit.SECONDS);

        if (!finished) {
            process.destroyForcibly();
            throw new RuntimeException("LibreOffice conversion timed out (" + TIMEOUT_SECONDS + " seconds)");
        }

        int exitCode = process.exitValue();
        if (exitCode != 0) {
            throw new RuntimeException("LibreOffice conversion failed. Exit code: " + exitCode + "\nOutput: " + output);
        }

        // Verify output file exists
        String baseName = input.getName().replaceAll("\\.[^.]+$", "");
        File outputPdf = new File(outputDir, baseName + ".pdf");
        if (!outputPdf.exists()) {
            throw new RuntimeException("PDF file was not generated: " + outputPdf.getPath());
        }
    }

    /**
     * Check if LibreOffice is installed.
     *
     * @return true if LibreOffice is available
     */
    public static boolean isLibreOfficeAvailable() {
        try {
            ProcessBuilder pb = new ProcessBuilder("libreoffice", "--version");
            Process process = pb.start();
            boolean finished = process.waitFor(10, TimeUnit.SECONDS);
            if (finished && process.exitValue() == 0) {
                try (BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream()))) {
                    String version = reader.readLine();
                    System.out.println("    LibreOffice detected: " + version);
                }
                return true;
            }
        } catch (Exception e) {
            // LibreOffice not found
        }
        return false;
    }
}
