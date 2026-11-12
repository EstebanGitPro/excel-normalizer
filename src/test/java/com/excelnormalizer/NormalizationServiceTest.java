package com.excelnormalizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCacheSource;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotCacheDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataConsolidateFunction;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFieldSortType;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class NormalizationServiceTest {

    @TempDir
    Path dir;

    // ---------------------------------------------------------------- AUXILIAR "Meses"

    @Test
    void auxiliarHasMonthColumnWithFirstOfMonth() throws IOException {
        try (XSSFWorkbook out = normalize(movement("53959596", LocalDateTime.of(2026, 5, 31, 10, 37)))) {
            Sheet auxiliar = out.getSheet("AUXILIAR");
            Cell header = auxiliar.getRow(0).getCell(24);
            Cell month = auxiliar.getRow(1).getCell(24);

            assertEquals("Meses", header.getStringCellValue());
            assertEquals(LocalDate.of(2026, 5, 1), month.getLocalDateTimeCellValue().toLocalDate());
            assertEquals("[$-240A]mmm-yy", month.getCellStyle().getDataFormatString());
        }
    }

    @Test
    void monthsOfDifferentYearsStayDistinct() throws IOException {
        try (XSSFWorkbook out = normalize(
                movement("110505", LocalDateTime.of(2025, 12, 15, 0, 0)),
                movement("110505", LocalDateTime.of(2026, 1, 10, 0, 0)))) {
            Sheet auxiliar = out.getSheet("AUXILIAR");

            assertEquals(LocalDate.of(2025, 12, 1), monthAt(auxiliar, 1));
            assertEquals(LocalDate.of(2026, 1, 1), monthAt(auxiliar, 2));
        }
    }

    // ---------------------------------------------------------------- FECHA validation

    @Test
    void invalidDatesCancelWithoutWritingFile() throws IOException {
        Path output = dir.resolve("out.xlsx");
        writeSources(List.of(
                movement("110505", LocalDateTime.of(2026, 2, 1, 0, 0)), // source row 2: valid
                movement("110505", "03/04/2026"),                         // row 3: text
                movement("110505", LocalDateTime.of(2026, 2, 2, 0, 0)), // row 4: valid
                movement("110505", null),                                 // row 5: blank
                movement("110505", 0.0),                                  // row 6: zero
                movement("110505", 20260531.0)));                         // row 7: yyyymmdd

        IllegalArgumentException error = assertThrows(IllegalArgumentException.class,
                () -> new NormalizationService().normalize(sourcePaths(), output));

        assertEquals("FECHA is not a valid date in Movimiento Contabilidad file, rows: 3, 5, 6, 7",
                error.getMessage());
        assertFalse(Files.exists(output));
    }

    @Test
    void moreThanTwentyInvalidDatesAreTruncated() throws IOException {
        List<Map<String, Object>> movements = new ArrayList<>();
        for (int i = 0; i < 25; i++) {
            movements.add(movement("110505", "not a date"));
        }
        writeSources(movements);

        IllegalArgumentException error = assertThrows(IllegalArgumentException.class,
                () -> new NormalizationService().normalize(sourcePaths(), dir.resolve("out.xlsx")));

        assertTrue(error.getMessage().startsWith(
                "FECHA is not a valid date in Movimiento Contabilidad file, rows: 2, 3, 4,"));
        assertTrue(error.getMessage().endsWith(", 21 and 5 more"), error.getMessage());
    }

    @Test
    void emptyMovimientosCancels() throws IOException {
        writeSources(List.of());

        IllegalArgumentException error = assertThrows(IllegalArgumentException.class,
                () -> new NormalizationService().normalize(sourcePaths(), dir.resolve("out.xlsx")));

        assertEquals("Movimiento Contabilidad file has no data rows", error.getMessage());
    }

    // ---------------------------------------------------------------- formula results

    @Test
    void formulasHaveCachedValues() throws IOException {
        try (XSSFWorkbook out = normalize(movement("53959596", LocalDateTime.of(2026, 3, 4, 0, 0)))) {
            Row row = out.getSheet("AUXILIAR").getRow(1);
            assertEquals("5", row.getCell(0).getStringCellValue());
            assertEquals("53", row.getCell(1).getStringCellValue());
            assertEquals("5395", row.getCell(2).getStringCellValue());
            assertEquals(179000.0, row.getCell(23).getNumericCellValue());

            for (Sheet sheet : out) {
                for (Row r : sheet) {
                    for (Cell cell : r) {
                        if (cell.getCellType() == CellType.FORMULA) {
                            assertTrue(((XSSFCell) cell).getCTCell().isSetV(),
                                    "no cached value in " + sheet.getSheetName() + "!"
                                            + cell.getAddress());
                        }
                    }
                }
            }
        }
    }

    // ---------------------------------------------------------------- TD pivot

    @Test
    void tdSheetIsLastAndHasPivot() throws IOException {
        try (XSSFWorkbook out = normalize(movement("110505", LocalDateTime.of(2026, 1, 5, 0, 0)))) {
            List<String> names = new ArrayList<>();
            out.forEach(sheet -> names.add(sheet.getSheetName()));

            assertEquals(List.of("MAYOR", "SALDO TERCERO", "AUXILIAR", "TD"), names);
            assertEquals(1, out.getSheet("TD").getPivotTables().size());
        }
    }

    @Test
    void pivotLayoutMatchesReference() throws IOException {
        try (XSSFWorkbook out = normalize(movement("110505", LocalDateTime.of(2026, 1, 5, 0, 0)))) {
            XSSFPivotTable pivot = out.getSheet("TD").getPivotTables().getFirst();
            CTPivotTableDefinition definition = pivot.getCTPivotTableDefinition();
            int n1 = 0, n2 = 1, n3 = 2, nomcta = 4, cnomnit = 12, neto = 23, meses = 24;

            List<Integer> rowFields = new ArrayList<>();
            definition.getRowFields().getFieldList().forEach(f -> rowFields.add(f.getX()));
            assertEquals(List.of(n1, n2, n3, nomcta, cnomnit), rowFields);
            assertEquals(meses, definition.getColFields().getFieldArray(0).getX());

            CTDataField data = definition.getDataFields().getDataFieldArray(0);
            assertEquals(neto, data.getFld());
            assertEquals(STDataConsolidateFunction.SUM, data.getSubtotal());
            assertEquals("Suma de NETO", data.getName());

            CTPivotCacheDefinition cache = pivot.getPivotCacheDefinition().getCTPivotCacheDefinition();
            assertTrue(cache.getRefreshOnLoad());

            for (int index : List.of(n1, n2, n3, nomcta, cnomnit, meses)) {
                assertEquals(STFieldSortType.ASCENDING, field(definition, index).getSortType(),
                        "sortType of field " + index);
            }
            for (int index : List.of(n1, n2, n3, nomcta)) {
                CTPivotField field = field(definition, index);
                assertFalse(field.getOutline(), "outline of field " + index);
                assertFalse(field.getDefaultSubtotal(), "defaultSubtotal of field " + index);
                assertFalse(field.isSetItems(), "items of field " + index);
            }
            for (int index : List.of(cnomnit, meses)) {
                assertEquals(1, field(definition, index).getItems().getCount(), "items of " + index);
            }
        }
    }

    @Test
    void pivotSourceCoversAllRows() throws IOException {
        List<Map<String, Object>> movements = new ArrayList<>();
        for (int day = 1; day <= 5; day++) {
            movements.add(movement("110505", LocalDateTime.of(2026, 1, day, 0, 0)));
        }
        try (XSSFWorkbook out = normalize(movements)) {
            XSSFPivotTable pivot = out.getSheet("TD").getPivotTables().getFirst();
            CTPivotCacheDefinition cache = pivot.getPivotCacheDefinition().getCTPivotCacheDefinition();
            CTCacheSource source = cache.getCacheSource();

            assertEquals("AUXILIAR", source.getWorksheetSource().getSheet());
            assertEquals("A1:Y6", source.getWorksheetSource().getRef());
            assertEquals(25, cache.getCacheFields().getCount());
        }
    }

    // ---------------------------------------------------------------- fixtures

    private static Map<String, Object> movement(String codcta, Object fecha) {
        Map<String, Object> row = new java.util.HashMap<>();
        row.put("codcta", codcta);
        row.put("nomcta", "CUENTA " + codcta);
        row.put("FECHA", fecha);
        row.put("DEBITO", 179000.0);
        row.put("CREDITO", 0.0);
        row.put("cNOMNIT", "TERCERO S.A.S.");
        return row;
    }

    @SafeVarargs
    private XSSFWorkbook normalize(Map<String, Object>... movements) throws IOException {
        return normalize(List.of(movements));
    }

    private XSSFWorkbook normalize(List<Map<String, Object>> movements) throws IOException {
        writeSources(movements);
        Path output = dir.resolve("out.xlsx");
        new NormalizationService().normalize(sourcePaths(), output);
        try (InputStream in = Files.newInputStream(output)) {
            return new XSSFWorkbook(in);
        }
    }

    private List<Path> sourcePaths() {
        return List.of(dir.resolve("mayor.xlsx"), dir.resolve("saldo.xlsx"),
                dir.resolve("movimientos.xlsx"));
    }

    /** Writes the three exports with the minimum headers each sheet writer requires. */
    private void writeSources(List<Map<String, Object>> movements) throws IOException {
        // Libro Mayor: one row per account class 1..7, needed by the MAYOR control formulas.
        List<Map<String, Object>> mayor = new ArrayList<>();
        for (int accountClass = 1; accountClass <= 7; accountClass++) {
            mayor.add(Map.of("CODCTA", String.valueOf(accountClass), "Tnomcta", "CLASE " + accountClass,
                    "ANTERIOR", 100.0, "DEBITOS", 50.0, "CREDITOS", 50.0, "SALDO_FINAL", 100.0));
        }
        writeSheet(dir.resolve("mayor.xlsx"), MayorSheet.COLUMNS_TO_KEEP, mayor);

        writeSheet(dir.resolve("saldo.xlsx"), SaldoTerceroSheet.COLUMNS_TO_KEEP, List.of(
                Map.of("Código Cuenta", "110505", "Nombre del Concepto Contable", "CAJA",
                        "NIT_CEDULA", "900123456", "Nombre de la Razon Social", "TERCERO S.A.S.",
                        "SALDO_ANTERIOR", 0.0, "DEBITOS", 100.0, "CREDITOS", 50.0,
                        "SALDO_FINAL", 50.0)));

        writeSheet(dir.resolve("movimientos.xlsx"), AuxiliarSheet.COLUMNS_TO_KEEP, movements);
    }

    private static void writeSheet(Path file, List<String> headers, List<Map<String, Object>> rows)
            throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(); OutputStream out = Files.newOutputStream(file)) {
            XSSFSheet sheet = workbook.createSheet("Sheet1");
            Row header = sheet.createRow(0);
            for (int c = 0; c < headers.size(); c++) {
                header.createCell(c).setCellValue(headers.get(c));
            }
            for (int r = 0; r < rows.size(); r++) {
                Row row = sheet.createRow(r + 1);
                for (int c = 0; c < headers.size(); c++) {
                    Object value = rows.get(r).get(headers.get(c));
                    Cell cell = row.createCell(c);
                    switch (value) {
                        case null -> { }
                        case String s -> cell.setCellValue(s);
                        case Double d -> cell.setCellValue(d);
                        case LocalDateTime t -> cell.setCellValue(t);
                        default -> throw new IllegalArgumentException("Unsupported " + value);
                    }
                }
            }
            workbook.write(out);
        }
    }

    private static LocalDate monthAt(Sheet auxiliar, int row) {
        return auxiliar.getRow(row).getCell(24).getLocalDateTimeCellValue().toLocalDate();
    }

    private static CTPivotField field(CTPivotTableDefinition definition, int index) {
        return definition.getPivotFields().getPivotFieldArray(index);
    }
}
