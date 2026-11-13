package com.excelnormalizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * Builds the SALDO TERCERO sheet from the frm_con_Saldo_Terceros_Documento export:
 * account balances broken down by third party (NIT). Layout and styling follow the
 * reference report; rows without a NIT are account subtotals and render bold red.
 */
final class SaldoTerceroSheet {

    static final List<String> COLUMNS_TO_KEEP = List.of(
            "Código Cuenta", "Nombre del Concepto Contable", "NIT_CEDULA",
            "Nombre de la Razon Social", "SALDO_ANTERIOR", "DEBITOS", "CREDITOS", "SALDO_FINAL");
    private static final List<String> DERIVED_COLUMNS = List.of("N1", "N2", "N3");
    private static final int NIT_POSITION = 2;
    private static final int FIRST_NUMERIC_POSITION = 4;

    private final ReportStyles styles;

    SaldoTerceroSheet(ReportStyles styles) {
        this.styles = styles;
    }

    void write(Sheet sourceSheet, XSSFWorkbook workbook) {
        Map<String, Integer> headerIndex = Cells.headerIndex(sourceSheet);
        List<String> missing = COLUMNS_TO_KEEP.stream()
                .filter(c -> !headerIndex.containsKey(c))
                .toList();
        if (!missing.isEmpty()) {
            throw new IllegalArgumentException(
                    "Columns not found in Saldo Terceros file: " + missing);
        }

        List<Integer> sourceIndexes = new ArrayList<>();
        for (String column : COLUMNS_TO_KEEP) {
            sourceIndexes.add(headerIndex.get(column));
        }
        int derivedCount = DERIVED_COLUMNS.size();

        Sheet sheet = workbook.createSheet("SALDO TERCERO");
        sheet.createFreezePane(0, 1);
        CellStyle headerStyle = styles.header();
        ReportStyles.Pair subtotal = styles.red(ReportStyles.INTEGER_FORMAT);
        ReportStyles.Pair detail = styles.plain(ReportStyles.INTEGER_FORMAT);

        int targetRowNum = 0;
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
            String nit = Cells.stringValue(sourceRow.getCell(sourceIndexes.get(NIT_POSITION)));
            ReportStyles.Pair style = (nit == null || nit.isBlank()) ? subtotal : detail;
            writeDerivedFormulas(targetRow, style.text());
            for (int c = 0; c < sourceIndexes.size(); c++) {
                Cell targetCell = targetRow.createCell(c + derivedCount);
                targetCell.setCellStyle(c < FIRST_NUMERIC_POSITION ? style.text() : style.number());
                Cell sourceCell = sourceRow.getCell(sourceIndexes.get(c));
                if (sourceCell != null) {
                    Cells.copyValue(sourceCell, targetCell);
                }
            }
        }
        setColumnWidths(sheet);
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

    /**
     * N1/N2/N3 are the 1-, 2- and 4-digit prefixes of the account code (column D),
     * styled like the rest of the row.
     */
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

    /** Column widths copied from the reference report (Excel width units). */
    private static void setColumnWidths(Sheet sheet) {
        double[] widths = {5.8, 8.43, 8.43, 9.7, 35.0, 14.2, 31.7, 12.7, 23.8, 12.7, 12.7};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, (int) (widths[c] * 256));
        }
    }
}
