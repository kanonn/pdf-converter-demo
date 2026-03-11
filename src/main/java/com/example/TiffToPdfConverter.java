package com.example;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;

import javax.imageio.ImageIO;
import javax.imageio.ImageReader;
import javax.imageio.stream.ImageInputStream;
import java.awt.*;
import java.awt.geom.AffineTransform;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.util.Iterator;

/**
 * Converts TIFF files to PDF.
 * Portrait images (height > width) are rotated 90 degrees clockwise.
 */
public class TiffToPdfConverter {

    /**
     * Convert a TIFF file to PDF.
     *
     * @param tiffPath Input TIFF file path
     * @param pdfPath  Output PDF file path
     * @throws Exception if conversion fails
     */
    public static void convert(String tiffPath, String pdfPath) throws Exception {
        File tiffFile = new File(tiffPath);
        if (!tiffFile.exists()) {
            throw new IllegalArgumentException("TIFF file does not exist: " + tiffPath);
        }

        System.out.println("    Processing: " + tiffFile.getName());

        try (PDDocument document = new PDDocument()) {
            // Handle multi-page TIFF
            try (ImageInputStream iis = ImageIO.createImageInputStream(tiffFile)) {
                Iterator<ImageReader> readers = ImageIO.getImageReaders(iis);

                if (!readers.hasNext()) {
                    throw new RuntimeException("TIFF reader not found. Check imageio-tiff library.");
                }

                ImageReader reader = readers.next();
                reader.setInput(iis);

                int pageCount = reader.getNumImages(true);
                System.out.println("    Page count: " + pageCount);

                for (int i = 0; i < pageCount; i++) {
                    BufferedImage image = reader.read(i);
                    addImageToPdf(document, image, i + 1);
                }

                reader.dispose();
            }

            document.save(pdfPath);
            System.out.println("    Saved: " + pdfPath);
        }
    }

    /**
     * Add an image to PDF as a new page.
     * Portrait images are rotated 90 degrees clockwise.
     */
    private static void addImageToPdf(PDDocument document, BufferedImage originalImage, int pageNum) throws Exception {
        int width = originalImage.getWidth();
        int height = originalImage.getHeight();

        boolean needsRotation = height > width;  // Portrait image needs rotation

        BufferedImage finalImage;
        if (needsRotation) {
            System.out.println("    Page " + pageNum + ": Portrait detected (" + width + "x" + height + ") -> Rotating 90 degrees");
            finalImage = rotateImage(originalImage, 90);
        } else {
            System.out.println("    Page " + pageNum + ": Landscape (" + width + "x" + height + ") -> No rotation");
            finalImage = originalImage;
        }

        // Set PDF page size to match image
        float pdfWidth = finalImage.getWidth();
        float pdfHeight = finalImage.getHeight();

        // Scale to fit A4 if necessary
        float maxWidth = PDRectangle.A4.getWidth();
        float maxHeight = PDRectangle.A4.getHeight();
        float scale = Math.min(maxWidth / pdfWidth, maxHeight / pdfHeight);
        if (scale < 1) {
            pdfWidth *= scale;
            pdfHeight *= scale;
        }

        PDPage page = new PDPage(new PDRectangle(pdfWidth, pdfHeight));
        document.addPage(page);

        // Write image to memory as PNG
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(finalImage, "PNG", baos);
        byte[] imageBytes = baos.toByteArray();

        // Add image to PDF
        PDImageXObject pdImage = PDImageXObject.createFromByteArray(document, imageBytes, "image");

        try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
            contentStream.drawImage(pdImage, 0, 0, pdfWidth, pdfHeight);
        }
    }

    /**
     * Rotate an image by the specified degrees.
     *
     * @param image   Original image
     * @param degrees Rotation angle (clockwise, positive value)
     * @return Rotated image
     */
    private static BufferedImage rotateImage(BufferedImage image, int degrees) {
        int width = image.getWidth();
        int height = image.getHeight();

        // For 90 degree rotation, width and height are swapped
        int newWidth = height;
        int newHeight = width;

        BufferedImage rotated = new BufferedImage(newWidth, newHeight, image.getType());
        Graphics2D g2d = rotated.createGraphics();

        // Configure rotation transform
        AffineTransform transform = new AffineTransform();

        // Rotate 90 degrees clockwise
        // 1. Translate to center
        // 2. Rotate
        // 3. Adjust position
        transform.translate(newWidth / 2.0, newHeight / 2.0);
        transform.rotate(Math.toRadians(degrees));
        transform.translate(-width / 2.0, -height / 2.0);

        // Set anti-aliasing
        g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);

        g2d.drawImage(image, transform, null);
        g2d.dispose();

        return rotated;
    }
}
