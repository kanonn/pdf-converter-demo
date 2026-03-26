package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;

import java.io.*;
import java.nio.file.*;
import java.util.*;

/**
 * Replaces Windows fonts with Linux fonts in Excel files.
 * Handles: cells, shapes, text boxes (best effort for drawings).
 */
public class FontReplacer {

    // Font mapping: Windows font -> Linux font
    private static final Map<String, String> FONT_MAP = new HashMap<>();
    
    static {
        // Gothic family
        FONT_MAP.put("MS Gothic", "Noto Sans CJK JP");
        FONT_MAP.put("MS PGothic", "Noto Sans CJK JP");
        FONT_MAP.put("MS UI Gothic", "Noto Sans CJK JP");
        FONT_MAP.put("ＭＳ ゴシック", "Noto Sans CJK JP");
        FONT_MAP.put("ＭＳ Ｐゴシック", "Noto Sans CJK JP");
        
        // Mincho family
        FONT_MAP.put("MS Mincho", "Noto Sans CJK JP");
        FONT_MAP.put("MS PMincho", "Noto Sans CJK JP");
        FONT_MAP.put("ＭＳ 明朝", "Noto Sans CJK JP");
        FONT_MAP.put("ＭＳ Ｐ明朝", "Noto Sans CJK JP");
        
        // Meiryo
        FONT_MAP.put("Meiryo", "Noto Sans CJK JP");
        FONT_MAP.put("Meiryo UI", "Noto Sans CJK JP");
        FONT_MAP.put("メイリオ", "Noto Sans CJK JP");
        
        // Yu Gothic/Mincho
        FONT_MAP.put("Yu Gothic", "Noto Sans CJK JP");
        FONT_MAP.put("Yu Gothic UI", "Noto Sans CJK JP");
        FONT_MAP.put("Yu Mincho", "Noto Sans CJK JP");
        FONT_MAP.put("游ゴシック", "Noto Sans CJK JP");
        FONT_MAP.put("游明朝", "Noto Sans CJK JP");
        
        // HG fonts
        FONT_MAP.put("HGGothicB", "Noto Sans CJK JP");
        FONT_MAP.put("HGPGothicB", "Noto Sans CJK JP");
        FONT_MAP.put("HGSGothicB", "Noto Sans CJK JP");
        FONT_MAP.put("HGMinchoB", "Noto Sans CJK JP");
        
        // Chinese
        FONT_MAP.put("SimSun", "Noto Sans CJK JP");
        FONT_MAP.put("宋体", "Noto Sans CJK JP");
        
        // Western fonts
        FONT_MAP.put("Arial", "DejaVu Sans");
        FONT_MAP.put("Times New Roman", "DejaVu Serif");
        FONT_MAP.put("Calibri", "DejaVu Sans");
        FONT_MAP.put("Cambria", "DejaVu Serif");
    }

    /**
     * Replace fonts in Excel file and save to new file.
     *
     * @param inputPath  Input Excel file path
     * @param outputPath Output Excel file path
     * @return Statistics of replacements
     */
    public static ReplacementStats replaceAllFonts(String inputPath, String outputPath) throws Exception {
        System.out.println("    Font replacement started: " + inputPath);
        
        ReplacementStats stats = new ReplacementStats();
        
        try (FileInputStream fis = new FileInputStream(inputPath);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {
            
            // 1. Replace cell fonts
            replaceCellFonts(workbook, stats);
            
            // 2. Replace drawing/shape fonts
            replaceDrawingFonts(workbook, stats);
            
            // Save
            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                workbook.write(fos);
            }
        }
        
        System.out.println("    Font replacement complete:");
        System.out.println("      Cells processed: " + stats.cellsProcessed);
        System.out.println("      Fonts replaced: " + stats.fontsReplaced);
        System.out.println("      Shapes processed: " + stats.shapesProcessed);
        System.out.println("      Shape texts replaced: " + stats.shapeTextsReplaced);
        
        return stats;
    }

    /**
     * Replace fonts in all cells.
     */
    private static void replaceCellFonts(XSSFWorkbook workbook, ReplacementStats stats) {
        // Create replacement fonts cache
        Map<String, XSSFFont> fontCache = new HashMap<>();
        
        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    stats.cellsProcessed++;
                    
                    try {
                        CellStyle style = cell.getCellStyle();
                        if (style == null) continue;
                        
                        Font font = workbook.getFontAt(style.getFontIndexAsInt());
                        if (font == null) continue;
                        
                        String fontName = font.getFontName();
                        String replacementName = getReplacementFont(fontName);
                        
                        if (replacementName != null && !replacementName.equals(fontName)) {
                            // Get or create replacement font
                            XSSFFont newFont = getOrCreateFont(workbook, (XSSFFont) font, replacementName, fontCache);
                            
                            // Create new style with replaced font
                            CellStyle newStyle = workbook.createCellStyle();
                            newStyle.cloneStyleFrom(style);
                            newStyle.setFont(newFont);
                            cell.setCellStyle(newStyle);
                            
                            stats.fontsReplaced++;
                        }
                    } catch (Exception e) {
                        // Skip problematic cells
                    }
                }
            }
        }
    }

    /**
     * Replace fonts in drawings (shapes, text boxes, etc.).
     */
    private static void replaceDrawingFonts(XSSFWorkbook workbook, ReplacementStats stats) {
        for (Sheet sheet : workbook) {
            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            
            if (drawing == null) continue;
            
            for (XSSFShape shape : drawing.getShapes()) {
                stats.shapesProcessed++;
                
                try {
                    if (shape instanceof XSSFSimpleShape) {
                        replaceSimpleShapeFonts((XSSFSimpleShape) shape, stats);
                    } else if (shape instanceof XSSFShapeGroup) {
                        replaceGroupFonts((XSSFShapeGroup) shape, stats);
                    }
                } catch (Exception e) {
                    System.out.println("      Warning: Could not process shape: " + e.getMessage());
                }
            }
        }
    }

    /**
     * Replace fonts in a simple shape using POI's high-level API.
     */
    private static void replaceSimpleShapeFonts(XSSFSimpleShape shape, ReplacementStats stats) {
        try {
            // Use POI's TextParagraph API instead of low-level CT classes
            for (XSSFTextParagraph paragraph : shape.getTextParagraphs()) {
                for (XSSFTextRun run : paragraph.getTextRuns()) {
                    String fontFamily = run.getFontFamily();
                    if (fontFamily != null) {
                        String replacement = getReplacementFont(fontFamily);
                        if (replacement != null) {
                            run.setFontFamily(replacement);
                            stats.shapeTextsReplaced++;
                        }
                    }
                }
            }
        } catch (Exception e) {
            // Try low-level approach as fallback
            try {
                replaceLowLevelShapeFonts(shape, stats);
            } catch (Exception ex) {
                // Skip
            }
        }
    }

    /**
     * Low-level font replacement using XML directly.
     */
    private static void replaceLowLevelShapeFonts(XSSFSimpleShape shape, ReplacementStats stats) {
        try {
            CTShape ctShape = shape.getCTShape();
            if (ctShape == null) return;
            
            // Access txBody via XML
            org.apache.xmlbeans.XmlObject[] txBodyArr = ctShape.selectPath(
                "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:txBody");
            
            if (txBodyArr == null || txBodyArr.length == 0) return;
            
            for (org.apache.xmlbeans.XmlObject txBodyObj : txBodyArr) {
                // Find all font references
                org.apache.xmlbeans.XmlObject[] latinFonts = txBodyObj.selectPath(
                    "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:latin");
                org.apache.xmlbeans.XmlObject[] eaFonts = txBodyObj.selectPath(
                    "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:ea");
                
                replaceFontsInXmlObjects(latinFonts, stats);
                replaceFontsInXmlObjects(eaFonts, stats);
            }
        } catch (Exception e) {
            // Skip
        }
    }

    /**
     * Replace fonts in XML objects.
     */
    private static void replaceFontsInXmlObjects(org.apache.xmlbeans.XmlObject[] fonts, ReplacementStats stats) {
        if (fonts == null) return;
        
        for (org.apache.xmlbeans.XmlObject fontObj : fonts) {
            try {
                org.apache.xmlbeans.XmlCursor cursor = fontObj.newCursor();
                String typeface = cursor.getAttributeText(new javax.xml.namespace.QName("typeface"));
                if (typeface != null) {
                    String replacement = getReplacementFont(typeface);
                    if (replacement != null) {
                        cursor.setAttributeText(new javax.xml.namespace.QName("typeface"), replacement);
                        stats.shapeTextsReplaced++;
                    }
                }
                cursor.dispose();
            } catch (Exception e) {
                // Skip
            }
        }
    }

    /**
     * Replace fonts in grouped shapes.
     */
    private static void replaceGroupFonts(XSSFShapeGroup group, ReplacementStats stats) {
        // Iterate through shapes in the group
        try {
            for (XSSFShape shape : group) {
                if (shape instanceof XSSFSimpleShape) {
                    replaceSimpleShapeFonts((XSSFSimpleShape) shape, stats);
                }
            }
        } catch (Exception e) {
            // Skip
        }
    }

    /**
     * Get replacement font name.
     */
    private static String getReplacementFont(String fontName) {
        if (fontName == null) return null;
        return FONT_MAP.get(fontName);
    }

    /**
     * Get or create a replacement font.
     */
    private static XSSFFont getOrCreateFont(XSSFWorkbook workbook, XSSFFont originalFont, 
                                            String newFontName, Map<String, XSSFFont> cache) {
        String cacheKey = newFontName + "_" + originalFont.getFontHeightInPoints() + "_" + 
                         originalFont.getBold() + "_" + originalFont.getItalic();
        
        if (cache.containsKey(cacheKey)) {
            return cache.get(cacheKey);
        }
        
        XSSFFont newFont = workbook.createFont();
        newFont.setFontName(newFontName);
        newFont.setFontHeightInPoints(originalFont.getFontHeightInPoints());
        newFont.setBold(originalFont.getBold());
        newFont.setItalic(originalFont.getItalic());
        newFont.setUnderline(originalFont.getUnderline());
        newFont.setStrikeout(originalFont.getStrikeout());
        newFont.setColor(originalFont.getColor());
        
        cache.put(cacheKey, newFont);
        return newFont;
    }

    /**
     * Statistics of font replacements.
     */
    public static class ReplacementStats {
        public int cellsProcessed = 0;
        public int fontsReplaced = 0;
        public int shapesProcessed = 0;
        public int shapeTextsReplaced = 0;
    }
}
