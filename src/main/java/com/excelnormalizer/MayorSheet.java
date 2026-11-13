package com.excelnormalizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * Builds the MAYOR sheet from the frm_con_Libro_Mayor export: title banner,
 * derived hierarchy columns, kept data columns, and accounting control rows.
 */
final class MayorSheet {

    static final List<String> COLUMNS_TO_KEEP = List.of(
            "CODCTA", "Tnomcta", "ANTERIOR", "DEBITOS", "CREDITOS", "SALDO_FINAL");
    private static final List<String> DERIVED_COLUMNS = List.of("LARGO", "N1", "N2", "N3");
    private static final int TITLE_ROWS = 4;

    private static final Map<String, String> SPANISH_MONTHS = Map.ofEntries(
            Map.entry("ene", "ENERO"), Map.entry("feb", "FEBRERO"), Map.entry("mar", "MARZO"),
            Map.entry("abr", "ABRIL"), Map.entry("may", "MAYO"), Map.entry("jun", "JUNIO"),
            Map.entry("jul", "JULIO"), Map.entry("ago", "AGOSTO"), Map.entry("sep", "SEPTIEMBRE"),
            Map.entry("oct", "OCTUBRE"), Map.entry("nov", "NOVIEMBRE"), Map.entry("dic", "DICIEMBRE"));

    private final ReportStyles styles;

    MayorSheet(ReportStyles styles) {
        this.styles = styles;
    }

    void write(Sheet sourceSheet, XSSFWorkbook workbook) {
        Map<String, Integer> headerIndex = Cells.headerIndex(sourceSheet);
        List<String> missing = COLUMNS_TO_KEEP.stream()
                .filter(c -> !headerIndex.containsKey(c))
                .toList();
        if (!missing.isEmpty()) {
            throw new IllegalArgumentException(
                    "Columns not found in Libro Mayor file: " + missing);
        }

        Sheet sheet = workbook.createSheet("MAYOR");
        sheet.createFreezePane(0, TITLE_ROWS + 1);
        writeTitleRows(sheet, sourceSheet, headerIndex);
        writeData(sheet, sourceSheet, headerIndex);
        writeControlRows(sheet);
        setColumnWidths(sheet);
    }

    /** Company banner rows 1-4, merged across the data columns, from the export metadata. */
    private void writeTitleRows(Sheet sheet, Sheet sourceSheet, Map<String, Integer> headerIndex) {
        Row firstData = sourceSheet.getRow(sourceSheet.getFirstRowNum() + 1);
        String company = metadata(firstData, headerIndex, "lblNomCia");
        String nit = metadata(firstData, headerIndex, "Nit");
        String reportTitle = metadata(firstData, headerIndex, "TITULO");
        String dateRange = formatDate(metadata(firstData, headerIndex, "dFecIni"))
                + " a " + formatDate(metadata(firstData, headerIndex, "dFecFin"));

        CellStyle titleStyle = styles.title();
        String[] titles = {company, nit, reportTitle, dateRange};
        int firstCol = DERIVED_COLUMNS.size();
        int lastCol = firstCol + COLUMNS_TO_KEEP.size() - 1;
        for (int i = 0; i < titles.length; i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(firstCol);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(titleStyle);
            sheet.addMergedRegion(new CellRangeAddress(i, i, firstCol, lastCol));
        }
    }

    private static String metadata(Row row, Map<String, Integer> headerIndex, String column) {
        Integer index = headerIndex.get(column);
        String value = index == null ? null : Cells.stringValue(row.getCell(index));
        return value == null ? "" : value.trim();
    }

    /** Turns the export date format ("01-abr.-2026") into the report one ("01-ABRIL-2026"). */
    private static String formatDate(String exportDate) {
        String[] parts = exportDate.split("-");
        if (parts.length != 3) {
            return exportDate;
        }
        String month = parts[1].replace(".", "").toLowerCase();
        return parts[0] + "-" + SPANISH_MONTHS.getOrDefault(month, parts[1].toUpperCase())
                + "-" + parts[2];
    }

    private void writeData(Sheet sheet, Sheet sourceSheet, Map<String, Integer> headerIndex) {
        List<Integer> sourceIndexes = new ArrayList<>();
        for (String column : COLUMNS_TO_KEEP) {
            sourceIndexes.add(headerIndex.get(column));
        }
        int derivedCount = DERIVED_COLUMNS.size();

        CellStyle headerStyle = styles.header();
        Map<Integer, ReportStyles.Pair> levelStyles = buildLevelStyles();

        int targetRowNum = TITLE_ROWS;
        for (int r = sourceSheet.getFirstRowNum(); r <= sourceSheet.getLastRowNum(); r++) {
            Row sourceRow = sourceSheet.getRow(r);
            if (sourceRow == null) {
                continue;
            }
            Row targetRow = sheet.createRow(targetRowNum++);
            if (r == sourceSheet.getFirstRowNum()) {
                writeHeaderRow(targetRow, headerStyle);
                continue;
            }
            String codcta = Cells.stringValue(sourceRow.getCell(sourceIndexes.get(0)));
            ReportStyles.Pair style = styleForCode(levelStyles, codcta);
            writeDerivedFormulas(targetRow, style.text());
            for (int c = 0; c < sourceIndexes.size(); c++) {
                Cell targetCell = targetRow.createCell(c + derivedCount);
                targetCell.setCellStyle(c < 2 ? style.text() : style.number());
                Cell sourceCell = sourceRow.getCell(sourceIndexes.get(c));
                if (sourceCell != null) {
                    Cells.copyValue(sourceCell, targetCell);
                }
            }
        }
    }

    private void writeHeaderRow(Row row, CellStyle headerStyle) {
        for (int c = 0; c < DERIVED_COLUMNS.size(); c++) {
            Cell cell = row.createCell(c);
            cell.setCellValue(DERIVED_COLUMNS.get(c));
            cell.setCellStyle(headerStyle);
        }
        for (int c = 0; c < COLUMNS_TO_KEEP.size(); c++) {
            Cell cell = row.createCell(c + DERIVED_COLUMNS.size());
            cell.setCellValue(COLUMNS_TO_KEEP.get(c));
            cell.setCellStyle(headerStyle);
        }
    }

    /** LARGO, N1, N2, N3 mirror the reference report: LEN/LEFT over the CODCTA cell (column E). */
    private static void writeDerivedFormulas(Row row, CellStyle style) {
        int excelRow = row.getRowNum() + 1;
        String codctaRef = "E" + excelRow;
        String[] formulas = {
                "LEN(" + codctaRef + ")",
                "LEFT(" + codctaRef + ",1)",
                "LEFT(" + codctaRef + ",2)",
                "LEFT(" + codctaRef + ",4)"};
        for (int c = 0; c < formulas.length; c++) {
            Cell cell = row.createCell(c);
            cell.setCellFormula(formulas[c]);
            cell.setCellStyle(style);
        }
    }

    /**
     * Level color scheme from the reference report, keyed by CODCTA length:
     * 1-2 digits (class/group) bold red, 4 digits (account) bold blue,
     * 6 digits (sub-account) bold black, 8+ digits (auxiliary) plain black.
     */
    private Map<Integer, ReportStyles.Pair> buildLevelStyles() {
        Map<Integer, ReportStyles.Pair> map = new LinkedHashMap<>();
        ReportStyles.Pair red = styles.red(ReportStyles.INTEGER_FORMAT);
        map.put(1, red);
        map.put(2, red);
        map.put(4, styles.blue(ReportStyles.INTEGER_FORMAT));
        map.put(6, styles.bold(ReportStyles.INTEGER_FORMAT));
        map.put(8, styles.plain(ReportStyles.INTEGER_FORMAT));
        return map;
    }

    private static ReportStyles.Pair styleForCode(Map<Integer, ReportStyles.Pair> styles,
                                                  String codcta) {
        int length = codcta == null ? 0 : codcta.trim().length();
        ReportStyles.Pair match = styles.get(length);
        return match != null ? match : styles.get(8);
    }

    /**
     * Appends the accounting control rows (PYG, PATRIMONIO, EC CONTABLE, PARTIDA DOBLE)
     * found in the reference report. Formulas reference the class-level rows (single-digit
     * CODCTA) resolved by account code, so they survive changes in row positions.
     */
    private void writeControlRows(Sheet sheet) {
        int codctaCol = DERIVED_COLUMNS.size();
        Map<String, Integer> classRow = new LinkedHashMap<>();
        for (int r = TITLE_ROWS + 1; r <= sheet.getLastRowNum(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            Cell cell = row.getCell(codctaCol);
            if (cell != null && cell.getCellType() == CellType.STRING) {
                String code = cell.getStringCellValue().trim();
                if (code.length() == 1) {
                    classRow.put(code, r + 1);
                }
            }
        }

        int anterior = codctaCol + 2;
        String debitosCol = CellReference.convertNumToColString(anterior + 1);
        String creditosCol = CellReference.convertNumToColString(anterior + 2);
        String saldoCol = CellReference.convertNumToColString(anterior + 3);
        int labelCellIndex = anterior + 2;
        int formulaCellIndex = anterior + 3;

        ReportStyles.Pair controlStyle = styles.control();
        int rowNum = sheet.getLastRowNum() + 3;

        String pyg = balanceFormula(saldoCol, classRow, List.of("4"), List.of("5", "6", "7"));
        int pygRow = rowNum + 1;
        addControlRow(sheet, rowNum++, labelCellIndex, formulaCellIndex, "PYG", pyg, controlStyle);

        String patrimonio = classRow.containsKey("3")
                ? "+" + saldoCol + classRow.get("3") + "+" + saldoCol + pygRow
                : "+" + saldoCol + pygRow;
        int patrimonioRow = rowNum + 1;
        addControlRow(sheet, rowNum++, labelCellIndex, formulaCellIndex, "PATRIMONIO", patrimonio,
                controlStyle);

        String ecContable = balanceFormula(saldoCol, classRow, List.of("1"), List.of("2"))
                + "-" + saldoCol + patrimonioRow;
        addControlRow(sheet, rowNum++, labelCellIndex, formulaCellIndex, "EC CONTABLE", ecContable,
                controlStyle);

        StringBuilder partidaDoble = new StringBuilder();
        for (int r : classRow.values()) {
            partidaDoble.append("+").append(debitosCol).append(r);
        }
        for (int r : classRow.values()) {
            partidaDoble.append("-").append(creditosCol).append(r);
        }
        addControlRow(sheet, rowNum, labelCellIndex, formulaCellIndex, "PARTIDA DOBLE",
                partidaDoble.toString(), controlStyle);
    }

    private static String balanceFormula(String col, Map<String, Integer> classRow,
                                         List<String> plus, List<String> minus) {
        StringBuilder formula = new StringBuilder();
        for (String code : plus) {
            if (classRow.containsKey(code)) {
                formula.append("+").append(col).append(classRow.get(code));
            }
        }
        for (String code : minus) {
            if (classRow.containsKey(code)) {
                formula.append("-").append(col).append(classRow.get(code));
            }
        }
        return formula.toString();
    }

    private static void addControlRow(Sheet sheet, int rowNum, int labelCellIndex,
                                      int formulaCellIndex, String label, String formula,
                                      ReportStyles.Pair style) {
        Row row = sheet.createRow(rowNum);
        Cell labelCell = row.createCell(labelCellIndex);
        labelCell.setCellValue(label);
        labelCell.setCellStyle(style.text());
        Cell formulaCell = row.createCell(formulaCellIndex);
        formulaCell.setCellFormula(formula);
        formulaCell.setCellStyle(style.number());
    }

    /** Column widths copied from the reference report (Excel width units). */
    private static void setColumnWidths(Sheet sheet) {
        double[] widths = {4.16, 8.43, 8.43, 6.33, 9.66, 27.66, 16.5, 13.5, 15.16, 16.5};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, (int) (widths[c] * 256));
        }
    }
}
