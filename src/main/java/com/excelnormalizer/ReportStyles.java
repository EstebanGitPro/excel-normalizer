package com.excelnormalizer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Cell styles copied from the reference report (Informes Contables).
 * Styles are workbook-level objects: build once, reuse across cells.
 */
final class ReportStyles {

    static final String INTEGER_FORMAT = "#,##0;[Red]\\(#,##0\\)";
    private static final byte[] BLUE_ACCOUNT = {(byte) 0x00, (byte) 0x00, (byte) 0xFF};

    /** Styles for one row kind: text cells and numeric cells. */
    record Pair(CellStyle text, CellStyle number) {}

    private final XSSFWorkbook workbook;

    ReportStyles(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    /** Merged title rows: bold 12pt centered, matching the reference report banner. */
    CellStyle title() {
        CellStyle style = newStyle();
        Font font = workbook.createFont();
        font.setFontName("Albertus Extra Bold");
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /** Column header shared by every sheet: bold Arial 10 on grey, medium borders. */
    CellStyle header() {
        CellStyle style = newStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        borders(style, BorderStyle.MEDIUM);
        style.setFont(arial(10, true, null));
        return style;
    }

    /** Free-form data cell style; numberFormat and fontName may be null for defaults. */
    CellStyle custom(String fontName, int points, String numberFormat, boolean borders) {
        CellStyle style = newStyle();
        XSSFFont font = workbook.createFont();
        font.setFontName(fontName == null ? "Arial" : fontName);
        font.setFontHeightInPoints((short) points);
        style.setFont(font);
        if (borders) {
            thinBorders(style);
        }
        if (numberFormat != null) {
            style.setDataFormat(workbook.createDataFormat().getFormat(numberFormat));
        }
        return style;
    }

    Pair red(String numberFormat) {
        XSSFFont font = arial(8, true, null);
        font.setColor(IndexedColors.RED.getIndex());
        return pair(font, numberFormat);
    }

    Pair blue(String numberFormat) {
        XSSFFont font = arial(8, true, null);
        font.setColor(new XSSFColor(BLUE_ACCOUNT, null));
        return pair(font, numberFormat);
    }

    Pair bold(String numberFormat) {
        return pair(arial(8, true, null), numberFormat);
    }

    Pair plain(String numberFormat) {
        return pair(arial(8, false, null), numberFormat);
    }

    /** Control rows at the bottom of MAYOR: Arial 10, no borders. */
    Pair control() {
        XSSFFont font = arial(10, false, null);
        CellStyle text = newStyle();
        text.setFont(font);
        CellStyle number = newStyle();
        number.setFont(font);
        number.setDataFormat(workbook.createDataFormat().getFormat(INTEGER_FORMAT));
        return new Pair(text, number);
    }

    private Pair pair(Font font, String numberFormat) {
        CellStyle text = newStyle();
        text.setFont(font);
        thinBorders(text);
        CellStyle number = newStyle();
        number.setFont(font);
        thinBorders(number);
        number.setDataFormat(workbook.createDataFormat().getFormat(numberFormat));
        return new Pair(text, number);
    }

    private XSSFFont arial(int points, boolean bold, XSSFColor color) {
        XSSFFont font = workbook.createFont();
        font.setFontName("Arial");
        font.setFontHeightInPoints((short) points);
        font.setBold(bold);
        if (color != null) {
            font.setColor(color);
        }
        return font;
    }

    private static void thinBorders(CellStyle style) {
        borders(style, BorderStyle.THIN);
    }

    private static void borders(CellStyle style, BorderStyle border) {
        style.setBorderTop(border);
        style.setBorderBottom(border);
        style.setBorderLeft(border);
        style.setBorderRight(border);
    }

    /** Every report cell is centered, so every style starts from here. */
    private CellStyle newStyle() {
        CellStyle style = workbook.createCellStyle();
        centered(style);
        return style;
    }

    private static void centered(CellStyle style) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    }
}
