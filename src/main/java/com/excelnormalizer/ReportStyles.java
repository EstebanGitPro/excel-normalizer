package com.excelnormalizer;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Cell styles copied from the reference report (Informes Contables).
 * Styles are workbook-level objects: build once, reuse across cells.
 */
final class ReportStyles {

    static final String INTEGER_FORMAT = "#,##0;[Red]\\(#,##0\\)";
    static final String DECIMAL_FORMAT = "#,##0.00";
    private static final byte[] BLUE_ACCOUNT = {(byte) 0x00, (byte) 0x70, (byte) 0xC0};

    /** Styles for one row kind: text cells and numeric cells. */
    record Pair(CellStyle text, CellStyle number) {}

    private final XSSFWorkbook workbook;

    ReportStyles(XSSFWorkbook workbook) {
        this.workbook = workbook;
    }

    /** Merged title rows: bold 12pt centered, matching the reference report banner. */
    CellStyle title() {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("Albertus Extra Bold");
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
    }

    /** Grey column header (fill indexed 22, thin borders, Arial 10). */
    CellStyle header(boolean bold) {
        return header(bold, true);
    }

    CellStyle header(boolean bold, boolean borders) {
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        if (borders) {
            thinBorders(style);
        }
        style.setFont(arial(10, bold, null));
        return style;
    }

    /** Free-form data cell style; numberFormat and fontName may be null for defaults. */
    CellStyle custom(String fontName, int points, String numberFormat, boolean borders) {
        CellStyle style = workbook.createCellStyle();
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

    /** Header cells outside the grey band (derived columns): plain Arial 10. */
    CellStyle plainHeader(boolean bold) {
        CellStyle style = workbook.createCellStyle();
        style.setFont(arial(10, bold, null));
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
        CellStyle text = workbook.createCellStyle();
        text.setFont(font);
        CellStyle number = workbook.createCellStyle();
        number.setFont(font);
        number.setDataFormat(workbook.createDataFormat().getFormat(INTEGER_FORMAT));
        return new Pair(text, number);
    }

    private Pair pair(Font font, String numberFormat) {
        CellStyle text = workbook.createCellStyle();
        text.setFont(font);
        thinBorders(text);
        CellStyle number = workbook.createCellStyle();
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
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }
}
