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
 * Builds the AUXILIAR sheet from the frm_con_Movimiento_Contabilidad export:
 * one row per accounting movement. Prepends N1/N2/N3 hierarchy formulas and
 * appends the NETO column (DEBITO - CREDITO), following the reference report.
 * The TD anual pivot in the reference workbook feeds off this sheet.
 */
final class AuxiliarSheet {

    private static final List<String> COLUMNS_TO_KEEP = List.of(
            "codcta", "nomcta", "FECHA", "FUENTE", "COMPROBANT", "DEBITO", "CREDITO",
            "DOCUMENTO", "NIT_CEDULA", "cNOMNIT", "CONCEPTO", "COD_COSTO", "TIPO_MOV",
            "BASE", "Codigo_Ciudad", "SEDE_LUGAR", "NOVEDAD", "Fecha_Registro", "REG", "porc");
    private static final List<String> DERIVED_COLUMNS = List.of("N1", "N2", "N3");
    private static final String NETO_COLUMN = "NETO";

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
        int netoIndex = derivedCount + COLUMNS_TO_KEEP.size();

        Sheet sheet = workbook.createSheet("AUXILIAR");
        CellStyle headerStyle = styles.header(false);
        CellStyle headerNoBorder = styles.header(false, false);
        CellStyle derivedHeaderStyle = styles.plainHeader(false);

        int targetRowNum = 0;
        for (int r = sourceSheet.getFirstRowNum(); r <= sourceSheet.getLastRowNum(); r++) {
            Row sourceRow = sourceSheet.getRow(r);
            if (sourceRow == null) {
                continue;
            }
            Row targetRow = sheet.createRow(targetRowNum++);
            if (r == sourceSheet.getFirstRowNum()) {
                writeHeaderRow(targetRow, headerStyle, headerNoBorder, derivedHeaderStyle,
                        netoIndex);
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
        }
        setColumnWidths(sheet);
    }

    private void writeHeaderRow(Row row, CellStyle headerStyle, CellStyle headerNoBorder,
                                CellStyle derivedHeaderStyle, int netoIndex) {
        for (int c = 0; c < DERIVED_COLUMNS.size(); c++) {
            Cell cell = row.createCell(c);
            cell.setCellValue(DERIVED_COLUMNS.get(c));
            cell.setCellStyle(derivedHeaderStyle);
        }
        for (int c = 0; c < COLUMNS_TO_KEEP.size(); c++) {
            Cell cell = row.createCell(c + DERIVED_COLUMNS.size());
            cell.setCellValue(COLUMNS_TO_KEEP.get(c));
            cell.setCellStyle(headerStyle);
        }
        Cell neto = row.createCell(netoIndex);
        neto.setCellValue(NETO_COLUMN);
        neto.setCellStyle(headerNoBorder);
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

    /** Column widths copied from the reference report (Excel width units). */
    private static void setColumnWidths(Sheet sheet) {
        double[] widths = {11.5, 8.43, 8.43, 9.7, 32.0, 12.7, 5.7, 10.5, 12.7, 8.43,
                14.2, 8.43, 31.7, 8.43, 6.2, 3.7, 12.7, 7.7, 8.43, 8.43, 15.5, 11.2, 7.7, 13.7};
        for (int c = 0; c < widths.length; c++) {
            sheet.setColumnWidth(c, (int) (widths[c] * 256));
        }
    }
}
