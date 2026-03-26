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
     * Also fixes textlink shapes by converting them to static text.
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
            
            // 1. Fix textlink shapes (convert dynamic text to static)
            fixTextLinkShapes(workbook, stats);
            
            // 2. Replace cell fonts
            replaceCellFonts(workbook, stats);
            
            // 3. Replace drawing/shape fonts
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
        System.out.println("      TextLinks fixed: " + stats.textLinksFixed);
        
        return stats;
    }
    
    /**
     * Fix shapes with textlink attribute by converting them to static text.
     * LibreOffice cannot render dynamic linked text properly.
     */
    private static void fixTextLinkShapes(XSSFWorkbook workbook, ReplacementStats stats) {
        System.out.println("    Fixing textlink shapes...");
        
        for (Sheet sheet : workbook) {
            XSSFSheet xssfSheet = (XSSFSheet) sheet;
            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            
            if (drawing == null) continue;
            
            for (XSSFShape shape : drawing.getShapes()) {
                try {
                    if (shape instanceof XSSFSimpleShape) {
                        XSSFSimpleShape simpleShape = (XSSFSimpleShape) shape;
                        fixTextLink(simpleShape, workbook, stats);
                    }
                } catch (Exception e) {
                    System.out.println("      Warning: Could not fix textlink for shape: " + e.getMessage());
                }
            }
        }
    }
    
    /**
     * Fix a single shape by removing problematic attributes and extensions.
     * LibreOffice cannot render shapes with textlink/macro/extLst properly.
     */
    private static void fixTextLink(XSSFSimpleShape shape, XSSFWorkbook workbook, ReplacementStats stats) {
        try {
            CTShape ctShape = shape.getCTShape();
            if (ctShape == null) return;
            
            String shapeName = shape.getShapeName();
            String xmlBefore = ctShape.xmlText();
            
            boolean hasTextlink = xmlBefore.contains("textlink=");
            boolean hasMacro = xmlBefore.contains("macro=");
            boolean hasExtLst = xmlBefore.contains("<a:extLst") || xmlBefore.contains("a16:creationId");
            
            if (!hasTextlink && !hasMacro && !hasExtLst) {
                return;
            }
            
            System.out.println("      SHAPE[" + shapeName + "] has textlink=" + hasTextlink + ", macro=" + hasMacro + ", extLst=" + hasExtLst);
            
            // Try to remove attributes using XmlCursor on the shape itself
            org.apache.xmlbeans.XmlCursor cursor = ctShape.newCursor();
            
            // Remove textlink attribute
            if (hasTextlink) {
                if (cursor.removeAttribute(new javax.xml.namespace.QName("textlink"))) {
                    System.out.println("        => Removed textlink from sp element");
                    stats.textLinksFixed++;
                }
            }
            
            // Remove macro attribute
            if (hasMacro) {
                if (cursor.removeAttribute(new javax.xml.namespace.QName("macro"))) {
                    System.out.println("        => Removed macro from sp element");
                    stats.textLinksFixed++;
                }
            }
            
            cursor.dispose();
            
            // Remove extLst elements using XPath
            if (hasExtLst) {
                try {
                    // Find and remove all extLst elements
                    org.apache.xmlbeans.XmlObject[] extLstArr = ctShape.selectPath(
                        "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' .//a:extLst");
                    
                    if (extLstArr != null && extLstArr.length > 0) {
                        for (org.apache.xmlbeans.XmlObject extLst : extLstArr) {
                            org.apache.xmlbeans.XmlCursor c = extLst.newCursor();
                            c.removeXml();
                            c.dispose();
                            System.out.println("        => Removed a:extLst element");
                            stats.textLinksFixed++;
                        }
                    }
                } catch (Exception e) {
                    System.out.println("        => Failed to remove extLst: " + e.getMessage());
                }
            }
            
            // Try to remove macro from nested cNvPr elements using XPath
            if (hasMacro) {
                try {
                    org.apache.xmlbeans.XmlObject[] cNvPrArr = ctShape.selectPath(
                        "declare namespace xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' .//xdr:cNvPr");
                    
                    if (cNvPrArr != null) {
                        for (org.apache.xmlbeans.XmlObject cNvPrObj : cNvPrArr) {
                            org.apache.xmlbeans.XmlCursor c = cNvPrObj.newCursor();
                            if (c.removeAttribute(new javax.xml.namespace.QName("macro"))) {
                                System.out.println("        => Removed macro from cNvPr element");
                                stats.textLinksFixed++;
                            }
                            c.dispose();
                        }
                    }
                } catch (Exception e) {
                    // XPath failed
                }
            }
            
            // Check if attributes are still present
            String xmlAfter = ctShape.xmlText();
            boolean stillHasProblems = xmlAfter.contains("textlink=") || 
                                       xmlAfter.contains("macro=") || 
                                       xmlAfter.contains("<a:extLst") ||
                                       xmlAfter.contains("a16:creationId");
            
            if (stillHasProblems) {
                System.out.println("        => Some attributes still present, trying clearText/setText");
                
                // Get existing text and re-set it
                String existingText = null;
                try {
                    existingText = shape.getText();
                } catch (Exception e) {
                    // ignore
                }
                
                if (existingText != null && !existingText.trim().isEmpty()) {
                    System.out.println("        => Existing text: \"" + existingText.substring(0, Math.min(existingText.length(), 30)) + "\"");
                    try {
                        shape.clearText();
                        shape.setText(existingText);
                        System.out.println("        => Re-set text SUCCESS");
                    } catch (Exception e) {
                        System.out.println("        => Re-set text FAILED: " + e.getMessage());
                    }
                }
            } else {
                System.out.println("        => All problematic attributes removed successfully");
            }
            
        } catch (Exception e) {
            System.out.println("      Warning: fixTextLink error: " + e.getClass().getSimpleName() + ": " + e.getMessage());
        }
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
     * Replace fonts in a simple shape.
     * Uses low-level XML manipulation since POI's high-level API doesn't persist changes.
     */
    private static int replaceSimpleShapeFonts(XSSFSimpleShape shape, ReplacementStats stats) {
        String shapeName = shape.getShapeName();
        int replacedCount = 0;
        
        try {
            CTShape ctShape = shape.getCTShape();
            if (ctShape == null) {
                System.out.println("SHAPE[" + shapeName + "] => SKIP: no CTShape");
                return 0;
            }
            
            // Get text body directly from XML
            if (!ctShape.isSetTxBody()) {
                System.out.println("SHAPE[" + shapeName + "] => SKIP: no txBody");
                return 0;
            }
            
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody txBody = ctShape.getTxBody();
            if (txBody == null) {
                System.out.println("SHAPE[" + shapeName + "] => SKIP: txBody is null");
                return 0;
            }
            
            // Iterate through paragraphs
            for (org.openxmlformats.schemas.drawingml.x2006.main.CTTextParagraph para : txBody.getPList()) {
                // Process runs
                for (org.openxmlformats.schemas.drawingml.x2006.main.CTRegularTextRun run : para.getRList()) {
                    String text = run.getT();
                    
                    if (run.isSetRPr()) {
                        org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties rPr = run.getRPr();
                        replacedCount += replaceCharacterFonts(shapeName, text, rPr, stats);
                    }
                }
                
                // Process default run properties
                if (para.isSetPPr() && para.getPPr().isSetDefRPr()) {
                    org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties defRPr = para.getPPr().getDefRPr();
                    replacedCount += replaceCharacterFonts(shapeName, "(default)", defRPr, stats);
                }
            }
            
            // Process body default properties
            if (txBody.isSetLstStyle()) {
                org.openxmlformats.schemas.drawingml.x2006.main.CTTextListStyle lstStyle = txBody.getLstStyle();
                if (lstStyle.isSetDefPPr() && lstStyle.getDefPPr().isSetDefRPr()) {
                    replacedCount += replaceCharacterFonts(shapeName, "(listStyle)", lstStyle.getDefPPr().getDefRPr(), stats);
                }
            }
            
            if (replacedCount > 0) {
                System.out.println("SHAPE[" + shapeName + "] => Total replaced: " + replacedCount);
            } else {
                System.out.println("SHAPE[" + shapeName + "] => No fonts replaced");
            }
            
        } catch (Exception e) {
            System.out.println("SHAPE[" + shapeName + "] ERROR: " + e.getClass().getSimpleName() + ": " + e.getMessage());
        }
        
        return replacedCount;
    }
    
    /**
     * Replace fonts in character properties (low-level XML).
     * Ensures both a:latin and a:ea are set for proper CJK rendering.
     */
    private static int replaceCharacterFonts(String shapeName, String text, 
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties rPr,
            ReplacementStats stats) {
        int count = 0;
        String textPreview = text != null ? text.substring(0, Math.min(text.length(), 15)) : "";
        
        String targetFont = "Noto Sans CJK JP";
        
        // Latin font (a:latin) - replace or add
        if (rPr.isSetLatin()) {
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont latin = rPr.getLatin();
            String fontName = latin.getTypeface();
            String replacement = getReplacementFont(fontName);
            if (replacement != null) {
                latin.setTypeface(replacement);
                count++;
                stats.shapeTextsReplaced++;
                System.out.println("SHAPE[" + shapeName + "] text=\"" + textPreview + "\" latin=\"" + fontName + "\" => " + replacement);
            }
        } else {
            // Add a:latin if missing
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont latin = rPr.addNewLatin();
            latin.setTypeface(targetFont);
            count++;
            stats.shapeTextsReplaced++;
            System.out.println("SHAPE[" + shapeName + "] text=\"" + textPreview + "\" => ADDED latin: " + targetFont);
        }
        
        // East Asian font (a:ea) - replace or add (CRITICAL for Japanese text!)
        if (rPr.isSetEa()) {
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont ea = rPr.getEa();
            String fontName = ea.getTypeface();
            String replacement = getReplacementFont(fontName);
            if (replacement != null) {
                ea.setTypeface(replacement);
                count++;
                stats.shapeTextsReplaced++;
                System.out.println("SHAPE[" + shapeName + "] text=\"" + textPreview + "\" ea=\"" + fontName + "\" => " + replacement);
            }
        } else {
            // Add a:ea if missing - CRITICAL for Japanese text rendering!
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont ea = rPr.addNewEa();
            ea.setTypeface(targetFont);
            count++;
            stats.shapeTextsReplaced++;
            System.out.println("SHAPE[" + shapeName + "] text=\"" + textPreview + "\" => ADDED ea: " + targetFont);
        }
        
        // Complex script font (a:cs) - replace if exists
        if (rPr.isSetCs()) {
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont cs = rPr.getCs();
            String fontName = cs.getTypeface();
            String replacement = getReplacementFont(fontName);
            if (replacement != null) {
                cs.setTypeface(replacement);
                count++;
                stats.shapeTextsReplaced++;
                System.out.println("SHAPE[" + shapeName + "] text=\"" + textPreview + "\" cs=\"" + fontName + "\" => " + replacement);
            }
        }
        
        return count;
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
        public int textLinksFixed = 0;
    }
}
