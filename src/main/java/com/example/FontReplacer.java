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
     * Fix a single shape's textlink by resolving the cell reference and setting static text.
     */
    private static void fixTextLink(XSSFSimpleShape shape, XSSFWorkbook workbook, ReplacementStats stats) {
        try {
            CTShape ctShape = shape.getCTShape();
            if (ctShape == null) return;
            
            // Get shape name for logging
            String shapeName = shape.getShapeName();
            
            // Check if this shape has a textlink by examining XML
            String xmlString = ctShape.xmlText();
            
            if (!xmlString.contains("textlink=")) {
                return;
            }
            
            // Extract textlink value
            int startIdx = xmlString.indexOf("textlink=\"");
            if (startIdx == -1) return;
            
            startIdx += 10; // length of 'textlink="'
            int endIdx = xmlString.indexOf("\"", startIdx);
            if (endIdx == -1) return;
            
            String textLink = xmlString.substring(startIdx, endIdx);
            System.out.println("      SHAPE[" + shapeName + "] has textlink: " + textLink);
            
            // Try to resolve the cell reference
            String cellValue = resolveCellReference(workbook, textLink);
            System.out.println("        Resolved cell value: " + (cellValue == null ? "NULL" : "\"" + cellValue + "\""));
            
            // Get existing text from shape
            String existingText = null;
            try {
                existingText = shape.getText();
            } catch (Exception e) {
                System.out.println("        Could not get existing text: " + e.getMessage());
            }
            System.out.println("        Existing text: " + (existingText == null ? "NULL" : "\"" + existingText + "\""));
            
            // Set static text if we have a cell value
            if (cellValue != null && !cellValue.isEmpty()) {
                try {
                    shape.setText(cellValue);
                    System.out.println("        => Set static text SUCCESS");
                    stats.textLinksFixed++;
                } catch (Exception e) {
                    System.out.println("        => Set static text FAILED: " + e.getMessage());
                }
            } else {
                System.out.println("        => Skipped: no cell value to set");
            }
            
        } catch (Exception e) {
            System.out.println("      Warning: fixTextLink error: " + e.getClass().getSimpleName() + ": " + e.getMessage());
        }
    }
    
    /**
     * Resolve a cell reference like "'Sheet1'!$A$1" to its value.
     */
    private static String resolveCellReference(XSSFWorkbook workbook, String textLink) {
        System.out.println("        Resolving: " + textLink);
        
        try {
            // Parse the textlink format: 'SheetName'!$A$1 or SheetName!$A$1 or just $A$1
            String sheetName = null;
            String cellRef = textLink;
            
            if (textLink.contains("!")) {
                int exclamIdx = textLink.lastIndexOf("!");
                sheetName = textLink.substring(0, exclamIdx);
                cellRef = textLink.substring(exclamIdx + 1);
                
                // Remove quotes from sheet name
                sheetName = sheetName.replace("'", "").trim();
            }
            
            // Remove $ signs from cell reference
            cellRef = cellRef.replace("$", "");
            
            System.out.println("          Sheet: " + (sheetName == null ? "(default)" : sheetName));
            System.out.println("          Cell: " + cellRef);
            
            // Get the sheet
            Sheet sheet;
            if (sheetName != null && !sheetName.isEmpty()) {
                sheet = workbook.getSheet(sheetName);
                if (sheet == null) {
                    System.out.println("          ERROR: Sheet not found: " + sheetName);
                    // List available sheets
                    System.out.println("          Available sheets:");
                    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                        System.out.println("            - " + workbook.getSheetName(i));
                    }
                    return null;
                }
            } else {
                sheet = workbook.getSheetAt(0);
            }
            
            System.out.println("          Found sheet: " + sheet.getSheetName());
            
            // Parse cell reference (e.g., "A1" -> row 0, col 0)
            org.apache.poi.ss.util.CellReference ref = new org.apache.poi.ss.util.CellReference(cellRef);
            System.out.println("          Row: " + ref.getRow() + ", Col: " + ref.getCol());
            
            Row row = sheet.getRow(ref.getRow());
            if (row == null) {
                System.out.println("          ERROR: Row not found");
                return null;
            }
            
            Cell cell = row.getCell(ref.getCol());
            if (cell == null) {
                System.out.println("          ERROR: Cell not found");
                return null;
            }
            
            System.out.println("          Cell type: " + cell.getCellType());
            
            // Get cell value as string
            switch (cell.getCellType()) {
                case STRING:
                    return cell.getStringCellValue();
                case NUMERIC:
                    return String.valueOf(cell.getNumericCellValue());
                case BOOLEAN:
                    return String.valueOf(cell.getBooleanCellValue());
                case FORMULA:
                    try {
                        return cell.getStringCellValue();
                    } catch (Exception e) {
                        return String.valueOf(cell.getNumericCellValue());
                    }
                default:
                    System.out.println("          ERROR: Unknown cell type");
                    return null;
            }
        } catch (Exception e) {
            System.out.println("          ERROR: " + e.getClass().getSimpleName() + ": " + e.getMessage());
            return null;
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
     */
    private static int replaceCharacterFonts(String shapeName, String text, 
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextCharacterProperties rPr,
            ReplacementStats stats) {
        int count = 0;
        String textPreview = text != null ? text.substring(0, Math.min(text.length(), 15)) : "";
        
        // Latin font (a:latin) - replace with Noto Sans CJK JP
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
        }
        
        // East Asian font (a:ea) - REMOVE it instead of replacing
        // LibreOffice seems to have issues when both a:latin and a:ea are present
        if (rPr.isSetEa()) {
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont ea = rPr.getEa();
            String fontName = ea.getTypeface();
            String replacement = getReplacementFont(fontName);
            if (replacement != null) {
                // Remove a:ea and let a:latin handle all text
                rPr.unsetEa();
                count++;
                stats.shapeTextsReplaced++;
                System.out.println("SHAPE[" + shapeName + "] text=\"" + textPreview + "\" ea=\"" + fontName + "\" => REMOVED (let latin handle)");
            }
        }
        
        // Complex script font (a:cs) - also remove
        if (rPr.isSetCs()) {
            org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont cs = rPr.getCs();
            String fontName = cs.getTypeface();
            String replacement = getReplacementFont(fontName);
            if (replacement != null) {
                rPr.unsetCs();
                count++;
                stats.shapeTextsReplaced++;
                System.out.println("SHAPE[" + shapeName + "] text=\"" + textPreview + "\" cs=\"" + fontName + "\" => REMOVED");
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
