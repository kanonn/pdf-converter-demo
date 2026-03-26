package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;

import java.io.*;
import java.nio.file.*;
import java.util.*;
import java.util.List;

/**
 * Replaces Windows fonts with Linux fonts in Excel files.
 * Handles: cells, shapes, text boxes (best effort for drawings).
 */
public class FontReplacer {

    // Font mapping: Windows font -> Linux font
    private static final Map<String, String> FONT_MAP = new HashMap<>();
    
    static {
        // Gothic family - English names
        FONT_MAP.put("MS Gothic", "Noto Sans CJK JP");
        FONT_MAP.put("MS PGothic", "Noto Sans CJK JP");
        FONT_MAP.put("MS UI Gothic", "Noto Sans CJK JP");
        
        // Gothic family - Japanese names (全角)
        FONT_MAP.put("ＭＳ ゴシック", "Noto Sans CJK JP");
        FONT_MAP.put("ＭＳ Ｐゴシック", "Noto Sans CJK JP");
        
        // Gothic family - Mixed (半角MS + 全角ゴシック) ← これが重要！
        FONT_MAP.put("MS ゴシック", "Noto Sans CJK JP");
        FONT_MAP.put("MS Pゴシック", "Noto Sans CJK JP");
        FONT_MAP.put("MS UIゴシック", "Noto Sans CJK JP");
        
        // Mincho family - English names
        FONT_MAP.put("MS Mincho", "Noto Sans CJK JP");
        FONT_MAP.put("MS PMincho", "Noto Sans CJK JP");
        
        // Mincho family - Japanese names (全角)
        FONT_MAP.put("ＭＳ 明朝", "Noto Sans CJK JP");
        FONT_MAP.put("ＭＳ Ｐ明朝", "Noto Sans CJK JP");
        
        // Mincho family - Mixed (半角MS + 全角明朝)
        FONT_MAP.put("MS 明朝", "Noto Sans CJK JP");
        FONT_MAP.put("MS P明朝", "Noto Sans CJK JP");
        
        // Meiryo - English
        FONT_MAP.put("Meiryo", "Noto Sans CJK JP");
        FONT_MAP.put("Meiryo UI", "Noto Sans CJK JP");
        
        // Meiryo - Japanese
        FONT_MAP.put("メイリオ", "Noto Sans CJK JP");
        
        // Yu Gothic/Mincho - English
        FONT_MAP.put("Yu Gothic", "Noto Sans CJK JP");
        FONT_MAP.put("Yu Gothic UI", "Noto Sans CJK JP");
        FONT_MAP.put("Yu Mincho", "Noto Sans CJK JP");
        
        // Yu Gothic/Mincho - Japanese
        FONT_MAP.put("游ゴシック", "Noto Sans CJK JP");
        FONT_MAP.put("游ゴシック体", "Noto Sans CJK JP");
        FONT_MAP.put("游明朝", "Noto Sans CJK JP");
        FONT_MAP.put("游明朝体", "Noto Sans CJK JP");
        
        // HG fonts
        FONT_MAP.put("HGGothicB", "Noto Sans CJK JP");
        FONT_MAP.put("HGPGothicB", "Noto Sans CJK JP");
        FONT_MAP.put("HGSGothicB", "Noto Sans CJK JP");
        FONT_MAP.put("HGMinchoB", "Noto Sans CJK JP");
        FONT_MAP.put("HGPMinchoB", "Noto Sans CJK JP");
        FONT_MAP.put("HGSMinchoB", "Noto Sans CJK JP");
        FONT_MAP.put("HG丸ｺﾞｼｯｸM-PRO", "Noto Sans CJK JP");
        FONT_MAP.put("HGP創英角ｺﾞｼｯｸUB", "Noto Sans CJK JP");
        FONT_MAP.put("HGS創英角ｺﾞｼｯｸUB", "Noto Sans CJK JP");
        FONT_MAP.put("HG創英角ｺﾞｼｯｸUB", "Noto Sans CJK JP");
        
        // Chinese
        FONT_MAP.put("SimSun", "Noto Sans CJK JP");
        FONT_MAP.put("宋体", "Noto Sans CJK JP");
        FONT_MAP.put("SimHei", "Noto Sans CJK JP");
        FONT_MAP.put("黑体", "Noto Sans CJK JP");
        
        // Western fonts
        FONT_MAP.put("Arial", "DejaVu Sans");
        FONT_MAP.put("Times New Roman", "DejaVu Serif");
        FONT_MAP.put("Calibri", "DejaVu Sans");
        FONT_MAP.put("Cambria", "DejaVu Serif");
        FONT_MAP.put("Century", "DejaVu Serif");
        FONT_MAP.put("Verdana", "DejaVu Sans");
        FONT_MAP.put("Tahoma", "DejaVu Sans");
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
            
            if (drawing == null) {
                System.out.println("      Sheet [" + sheet.getSheetName() + "]: No drawings found");
                continue;
            }
            
            System.out.println("      Sheet [" + sheet.getSheetName() + "]: Found " + drawing.getShapes().size() + " shapes");
            
            int shapeIndex = 0;
            for (XSSFShape shape : drawing.getShapes()) {
                shapeIndex++;
                stats.shapesProcessed++;
                
                try {
                    String shapeType = shape.getClass().getSimpleName();
                    String shapeName = shape.getShapeName();
                    
                    if (shape instanceof XSSFSimpleShape) {
                        XSSFSimpleShape simpleShape = (XSSFSimpleShape) shape;
                        int textCount = 0;
                        try {
                            textCount = simpleShape.getTextParagraphs().size();
                        } catch (Exception e) {
                            // ignore
                        }
                        
                        System.out.println("        [" + shapeIndex + "] " + shapeType + 
                            " name=\"" + shapeName + "\" textParagraphs=" + textCount);
                        
                        int replaced = replaceSimpleShapeFonts(simpleShape, stats);
                        if (replaced > 0) {
                            System.out.println("            → Replaced " + replaced + " font(s)");
                        } else {
                            System.out.println("            → No fonts to replace (or no text)");
                        }
                        
                    } else if (shape instanceof XSSFShapeGroup) {
                        System.out.println("        [" + shapeIndex + "] " + shapeType + 
                            " name=\"" + shapeName + "\" (group)");
                        replaceGroupFonts((XSSFShapeGroup) shape, stats);
                        
                    } else {
                        System.out.println("        [" + shapeIndex + "] " + shapeType + 
                            " name=\"" + shapeName + "\" (unsupported type)");
                    }
                } catch (Exception e) {
                    System.out.println("        [" + shapeIndex + "] ERROR: " + e.getMessage());
                }
            }
        }
    }

    /**
     * Replace fonts in a simple shape using POI's high-level API.
     * Returns the number of fonts replaced.
     */
    private static int replaceSimpleShapeFonts(XSSFSimpleShape shape, ReplacementStats stats) {
        int replacedCount = 0;
        String shapeName = shape.getShapeName();
        
        try {
            // Use POI's TextParagraph API
            List<XSSFTextParagraph> paragraphs = shape.getTextParagraphs();
            
            for (int pIdx = 0; pIdx < paragraphs.size(); pIdx++) {
                XSSFTextParagraph paragraph = paragraphs.get(pIdx);
                List<XSSFTextRun> runs = paragraph.getTextRuns();
                
                for (int rIdx = 0; rIdx < runs.size(); rIdx++) {
                    XSSFTextRun run = runs.get(rIdx);
                    String text = run.getText();
                    String fontFamily = run.getFontFamily();
                    String textPreview = (text != null ? text.substring(0, Math.min(text.length(), 20)) : "");
                    
                    // Build single log line
                    StringBuilder logLine = new StringBuilder();
                    logLine.append("SHAPE[").append(shapeName).append("] ");
                    logLine.append("P").append(pIdx).append("R").append(rIdx).append(": ");
                    logLine.append("font=\"").append(fontFamily).append("\" ");
                    logLine.append("text=\"").append(textPreview).append("\" ");
                    
                    if (fontFamily != null) {
                        String replacement = getReplacementFont(fontFamily);
                        if (replacement != null) {
                            try {
                                run.setFontFamily(replacement, (byte)0, (byte)0, false);
                                
                                // Verify the change
                                String afterFont = run.getFontFamily();
                                if (replacement.equals(afterFont)) {
                                    stats.shapeTextsReplaced++;
                                    replacedCount++;
                                    logLine.append("=> OK: ").append(replacement);
                                } else {
                                    logLine.append("=> SET BUT NOT SAVED: set=").append(replacement);
                                    logLine.append(" actual=").append(afterFont);
                                }
                            } catch (Exception setEx) {
                                logLine.append("=> FAILED: ").append(setEx.getClass().getSimpleName());
                                logLine.append(": ").append(setEx.getMessage());
                            }
                        } else {
                            logLine.append("=> SKIP: no mapping");
                        }
                    } else {
                        logLine.append("=> SKIP: font is null");
                    }
                    
                    System.out.println(logLine.toString());
                }
            }
        } catch (Exception e) {
            System.out.println("SHAPE[" + shapeName + "] ERROR: " + e.getClass().getSimpleName() + ": " + e.getMessage());
            // Try low-level approach as fallback
            try {
                replacedCount = replaceLowLevelShapeFonts(shape, stats);
            } catch (Exception ex) {
                System.out.println("SHAPE[" + shapeName + "] LOW-LEVEL ERROR: " + ex.getClass().getSimpleName() + ": " + ex.getMessage());
            }
        }
        
        return replacedCount;
    }

    /**
     * Low-level font replacement using XML directly.
     * Returns the number of fonts replaced.
     */
    private static int replaceLowLevelShapeFonts(XSSFSimpleShape shape, ReplacementStats stats) {
        int replacedCount = 0;
        
        try {
            CTShape ctShape = shape.getCTShape();
            if (ctShape == null) return 0;
            
            // Access txBody via XML
            org.apache.xmlbeans.XmlObject[] txBodyArr = ctShape.selectPath(
                "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:txBody");
            
            if (txBodyArr == null || txBodyArr.length == 0) {
                System.out.println("            Low-level: No txBody found");
                return 0;
            }
            
            for (org.apache.xmlbeans.XmlObject txBodyObj : txBodyArr) {
                // Find all font references
                org.apache.xmlbeans.XmlObject[] latinFonts = txBodyObj.selectPath(
                    "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:latin");
                org.apache.xmlbeans.XmlObject[] eaFonts = txBodyObj.selectPath(
                    "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:ea");
                
                System.out.println("            Low-level: Found " + 
                    (latinFonts != null ? latinFonts.length : 0) + " latin, " +
                    (eaFonts != null ? eaFonts.length : 0) + " ea fonts");
                
                replacedCount += replaceFontsInXmlObjects(latinFonts, stats);
                replacedCount += replaceFontsInXmlObjects(eaFonts, stats);
            }
        } catch (Exception e) {
            System.out.println("            Low-level error: " + e.getMessage());
        }
        
        return replacedCount;
    }

    /**
     * Replace fonts in XML objects.
     * Returns the number of fonts replaced.
     */
    private static int replaceFontsInXmlObjects(org.apache.xmlbeans.XmlObject[] fonts, ReplacementStats stats) {
        int replacedCount = 0;
        if (fonts == null) return 0;
        
        for (org.apache.xmlbeans.XmlObject fontObj : fonts) {
            try {
                org.apache.xmlbeans.XmlCursor cursor = fontObj.newCursor();
                String typeface = cursor.getAttributeText(new javax.xml.namespace.QName("typeface"));
                if (typeface != null) {
                    String replacement = getReplacementFont(typeface);
                    if (replacement != null) {
                        cursor.setAttributeText(new javax.xml.namespace.QName("typeface"), replacement);
                        stats.shapeTextsReplaced++;
                        replacedCount++;
                    }
                }
                cursor.dispose();
            } catch (Exception e) {
                // Skip
            }
        }
        
        return replacedCount;
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
