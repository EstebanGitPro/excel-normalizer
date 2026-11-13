package com.excelnormalizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;

/**
 * Builds the AUXILIAR sheet from the frm_con_Movimiento_Contabilidad export:
 * one row per accounting movement. Prepends N1/N2/N3 hierarchy formulas and
 * appends the NETO column (DEBITO - CREDITO), following the reference report.
 * Also appends the "Meses" column (first day of each row's month), which the
 * TD pivot sheet uses as its column field.
 */
final class AuxiliarSheet {

    static final String SHEET_NAME = "AUXILIAR";
    static final String MONTH_COLUMN = "Meses";
    static final String MONTH_FORMAT = "[$-240A]mmm-yy";

    static final List<String> COLUMNS_TO_KEEP = List.of(
            "codcta", "nomcta", "FECHA", "FUENTE", "COMPROBANT", "DEBITO", "CREDITO",
            "DOCUMENTO", "NIT_CEDULA", "cNOMNIT", "CONCEPTO", "COD_COSTO", "TIPO_MOV",
            "BASE", "Codigo_Ciudad", "SEDE_LUGAR", "NOVEDAD", "Fecha_Registro", "REG", "porc");
    private static final List<String> DERIVED_COLUMNS = List.of("N1", "N2", "N3");
    private static final String NETO_COLUMN = "NETO";
    private static final String DATE_COLUMN = "FECHA";

    /** Excel serial numbers for 1900-01-01 and 9999-12-31: anything outside is not a date. */
    private static final double MIN_EXCEL_DATE = 1;
    private static final double MAX_EXCEL_DATE = 2958465;
    private static final int MAX_ROWS_IN_ERROR = 20;

    /** Number format per kept column; null means General. */
    private static final Map<String, String> NUMBER_FORMATS = Map.of(
            "FECHA", "d-mmm-yy",
            "DEBITO", "#,##0.00",
            "CREDITO", "#,##0.00",
            "BASE", "#,##0.00",
            "REG", "#",
            "porc", "0.00%");

    /** Columns rendered in Arial Narrow in the reference report. */
    private static final List<String> NARROW_COLUMNS = List.of("nomcta", "cNOMNIT");

    private final ReportStyles styles;

    AuxiliarSheet(ReportStyles styles) {
        this.styles = styles;
    }

    void write(Sheet sourceSheet, XSSFWorkbook workbook) {
        Map<String, Integer> headerIndex = Cells.headerIndex(sourceSheet);
        List<String> missing = COLUMNS_TO_KEEP.stream()
                .filter(c -> !headerIndex.containsKey(c))
                .toList();
        if (!missing.isEmpty()) {
            throw new IllegalArgumentException(
                    "Columns not found in Movimiento Contabilidad file: " + missing);
        }

        List<Integer> sourceIndexes = new ArrayList<>();
        List<CellStyle> columnStyles = new ArrayList<>();
        for (String column : COLUMNS_TO_KEEP) {
            sourceIndexes.add(headerIndex.get(column));
            String fontName = NARROW_COLUMNS.contains(column) ? "Arial Narrow" : "Arial";
            columnStyles.add(styles.custom(fontName, 8, NUMBER_FORMATS.get(column), true));
        }
        int derivedCount = DERIVED_COLUMNS.size();
        CellStyle derivedStyle = styles.custom("Arial", 10, null, false);
        CellStyle netoStyle = styles.custom("Arial", 10, ReportStyles.INTEGER_FORMAT, false);
        CellStyle monthStyle = styles.custom("Arial", 10, MONTH_FORMAT, false);
        int netoIndex = derivedCount + COLUMNS_TO_KEEP.size();
        int monthIndex = netoIndex + 1;
        int dateSourceIndex = headerIndex.get(DATE_COLUMN);

        Sheet sheet = workbook.createSheet(SHEET_NAME);
        sheet.createFreezePane(0, 1);
        CellStyle headerStyle = styles.header();

        List<Integer> invalidDateRows = new ArrayList<>();
        int targetRowNum = 0;
        for (int r = sourceSheet.getFirstRowNum(); r <= sourceSheet.getLastRowNum(); r++) {
            Row sourceRow = sourceSheet.getRow(r);
            if (sourceRow == null) {
                continue;
            }
            Row targetRow = sheet.createRow(targetRowNum++);
            if (r == sourceSheet.getFirstRowNum()) {
                writeHeaderRow(targetRow, headerStyle, netoIndex, monthIndex);
                continue;
            }
            writeDerivedFormulas(targetRow, derivedStyle);
            for (int c = 0; c < sourceIndexes.size(); c++) {
                Cell targetCell = targetRow.createCell(c + derivedCount);
                targetCell.setCellStyle(columnStyles.get(c));
                Cell sourceCell = sourceRow.getCell(sourceIndexes.get(c));
                if (sourceCell != null) {
                    Cells.copyValue(sourceCell, targetCell);
                }
            }
            int excelRow = targetRow.getRowNum() + 1;
            Cell neto = targetRow.createCell(netoIndex);
            neto.setCellFormula("I" + excelRow + "-J" + excelRow);
            neto.setCellStyle(netoStyle);

            Optional<LocalDate> month = firstOfMonth(sourceRow.getCell(dateSourceIndex));
            if (month.isPresent()) {
                Cell monthCell = targetRow.createCell(monthIndex);
                monthCell.setCellValue(month.get());
                monthCell.setCellStyle(monthStyle);
            } else {
                invalidDateRows.add(r + 1); // row number as the user sees it in Excel
            }
        }

        if (targetRowNum <= 1) {
            throw new IllegalArgumentException("Movimiento Contabilidad file has no data rows");
        }
        if (!invalidDateRows.isEmpty()) {
            throw new IllegalArgumentException(invalidDatesMessage(invalidDateRows));
        }
        setColumnWidths(sheet);
    }

    /**
     * First day of the month of a FECHA cell, or empty when the cell is not a real
     * Excel date: blank, text, boolean, error, or a number outside the date range
     * (0, or a yyyymmdd number such as 20260531).
     */
    static Optional<LocalDate> firstOfMonth(Cell cell) {
        if (cell == null) {
            return Optional.empty();
        }
        boolean numeric = cell.getCellType() == CellType.NUMERIC
                || (cell.getCellType() == CellType.FORMULA
                && cell.getCachedFormulaResultType() == CellType.NUMERIC);
        if (!numeric) {
            return Optional.empty();
        }
        double value = cell.getNumericCellValue();
        if (value < MIN_EXCEL_DATE || value > MAX_EXCEL_DATE) {
            return Optional.empty();
        }
        return Optional.of(DateUtil.getLocalDateTime(value).toLocalDate().withDayOfMonth(1));
    }

    private static String invalidDatesMessage(List<Integer> rows) {
        String listed = rows.stream()
                .limit(MAX_ROWS_IN_ERROR)
                .map(String::valueOf)
                .collect(Collectors.joining(", "));
        String more = rows.size() > MAX_ROWS_IN_ERROR
                ? " and " + (rows.size() - MAX_ROWS_IN_ERROR) + " more"
                : "";
        return "FECHA is not a valid date in Movimiento Contabilidad file, rows: " + listed + more;
    }

    private void writeHeaderRow(Row row, CellStyle headerStyle, int netoIndex, int monthIndex) {
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

        Cell neto = row.createCell(netoIndex);
        neto.setCellValue(NETO_COLUMN);
        neto.setCellStyle(headerStyle);

        Cell month = row.createCell(monthIndex);
        month.setCellValue(MONTH_COLUMN);
        month.setCellStyle(headerStyle);
    }

    /** N1/N2/N3 mirror the reference report: 1-, 2- and 4-digit prefixes of codcta (column D). */
    private static void writeDerivedFormulas(Row row, CellStyle style) {
        int excelRow = row.getRowNum() + 1;
        String codeRef = "D" + excelRow;
        String[] formulas = {
                "LEFT(" + codeRef + ",1)",
                "LEFT(" + codeRef + ",2)",
                "LEFT(" + codeRef + ",4)"};
        for (int c = 0; c < formulas.length; c++) {
            Cell cell = row.createCell(c);
            cell.setCellFormula(formulas[c]);
            cell.setCellStyle(style);
        }
    }

    /** Column widths copied from the reference report (Excel width units), plus "Meses". */
    private static void setColumnWidths(Sheet sheet) {
        double[] widths = {11.5, 8.43, 8.43, 9.7, 32.0, 12.7, 5.7, 10.5, 12.7, 8.43,
                14.2, 8.43, 31.7, 8.43, 6.2, 3.7, 12.7, 7.7, 8.43, 8.43, 15.5, 11.2, 7.7, 13.7,
                9.0};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, (int) (widths[c] * 256));
        }
    }
}
