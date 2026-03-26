package com.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.*;

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
                        replaceShapeFonts((XSSFSimpleShape) shape, stats);
                    } else if (shape instanceof XSSFShapeGroup) {
                        // Handle grouped shapes
                        replaceGroupFonts((XSSFShapeGroup) shape, stats);
                    }
                } catch (Exception e) {
                    // Skip problematic shapes
                    System.out.println("      Warning: Could not process shape: " + e.getMessage());
                }
            }
        }
    }

    /**
     * Replace fonts in a simple shape.
     */
    private static void replaceShapeFonts(XSSFSimpleShape shape, ReplacementStats stats) {
        try {
            CTShape ctShape = shape.getCTShape();
            if (ctShape == null || !ctShape.isSetTxBody()) return;
            
            CTTextBody txBody = ctShape.getTxBody();
            if (txBody == null) return;
            
            for (CTTextParagraph paragraph : txBody.getPList()) {
                for (CTRegularTextRun run : paragraph.getRList()) {
                    if (run.isSetRPr()) {
                        CTTextCharacterProperties props = run.getRPr();
                        replaceFontInTextProps(props, stats);
                    }
                }
                
                // Default paragraph run properties
                if (paragraph.isSetPPr() && paragraph.getPPr().isSetDefRPr()) {
                    replaceFontInTextProps(paragraph.getPPr().getDefRPr(), stats);
                }
            }
            
            // Body default properties
            if (txBody.isSetBodyPr()) {
                // Body properties don't contain font info directly
            }
            
            // List style
            if (txBody.isSetLstStyle()) {
                CTTextListStyle lstStyle = txBody.getLstStyle();
                replaceListStyleFonts(lstStyle, stats);
            }
            
        } catch (Exception e) {
            // Skip
        }
    }

    /**
     * Replace fonts in grouped shapes.
     */
    private static void replaceGroupFonts(XSSFShapeGroup group, ReplacementStats stats) {
        // POI doesn't expose group children easily, skip for now
    }

    /**
     * Replace font in text character properties.
     */
    private static void replaceFontInTextProps(CTTextCharacterProperties props, ReplacementStats stats) {
        if (props == null) return;
        
        try {
            // Latin font
            if (props.isSetLatin()) {
                CTTextFont latin = props.getLatin();
                String fontName = latin.getTypeface();
                String replacement = getReplacementFont(fontName);
                if (replacement != null) {
                    latin.setTypeface(replacement);
                    stats.shapeTextsReplaced++;
                }
            }
            
            // East Asian font
            if (props.isSetEa()) {
                CTTextFont ea = props.getEa();
                String fontName = ea.getTypeface();
                String replacement = getReplacementFont(fontName);
                if (replacement != null) {
                    ea.setTypeface(replacement);
                    stats.shapeTextsReplaced++;
                }
            }
            
            // Complex script font
            if (props.isSetCs()) {
                CTTextFont cs = props.getCs();
                String fontName = cs.getTypeface();
                String replacement = getReplacementFont(fontName);
                if (replacement != null) {
                    cs.setTypeface(replacement);
                    stats.shapeTextsReplaced++;
                }
            }
        } catch (Exception e) {
            // Skip
        }
    }

    /**
     * Replace fonts in list style.
     */
    private static void replaceListStyleFonts(CTTextListStyle lstStyle, ReplacementStats stats) {
        // Handle different levels
        if (lstStyle.isSetDefPPr() && lstStyle.getDefPPr().isSetDefRPr()) {
            replaceFontInTextProps(lstStyle.getDefPPr().getDefRPr(), stats);
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
