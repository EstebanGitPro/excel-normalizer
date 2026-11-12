package com.excelnormalizer;

import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTItems;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPivotTableDefinition;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STFieldSortType;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STItemType;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;

/**
 * Builds the TD sheet: a native Excel pivot table over AUXILIAR, laid out like the
 * reference report (rows N1 > N2 > N3 > nomcta > cNOMNIT in tabular form, one column
 * per month, Sum of NETO). POI leaves the pivot cache empty and marks it to refresh
 * on load, so desktop Excel fills it when the file is opened.
 */
final class TdPivotSheet {

    static final String SHEET_NAME = "TD";
    static final List<String> ROW_FIELDS = List.of("N1", "N2", "N3", "nomcta", "cNOMNIT");
    static final String VALUE_FIELD = "NETO";
    static final String VALUE_LABEL = "Suma de NETO";

    /** Row fields shown each in its own column without subtotals; the last one stays default. */
    private static final List<String> TABULAR_FIELDS = ROW_FIELDS.subList(0, ROW_FIELDS.size() - 1);
    private static final String ANCHOR_CELL = "A3";

    void write(XSSFWorkbook workbook) {
        XSSFSheet auxiliar = workbook.getSheet(AuxiliarSheet.SHEET_NAME);
        if (auxiliar == null) {
            throw new IllegalArgumentException("Sheet not found: " + AuxiliarSheet.SHEET_NAME);
        }
        Map<String, Integer> headers = Cells.headerIndex(auxiliar);
        List<String> required = new ArrayList<>(ROW_FIELDS);
        required.add(AuxiliarSheet.MONTH_COLUMN);
        required.add(VALUE_FIELD);
        List<String> missing = required.stream().filter(f -> !headers.containsKey(f)).toList();
        if (!missing.isEmpty()) {
            throw new IllegalArgumentException(
                    "Columns not found in " + AuxiliarSheet.SHEET_NAME + " sheet: " + missing);
        }

        int lastColumn = Collections.max(headers.values());
        AreaReference source = new AreaReference(
                new CellReference(0, 0),
                new CellReference(auxiliar.getLastRowNum(), lastColumn),
                SpreadsheetVersion.EXCEL2007);

        XSSFSheet td = workbook.createSheet(SHEET_NAME);
        XSSFPivotTable pivot = td.createPivotTable(source, new CellReference(ANCHOR_CELL), auxiliar);
        for (String field : ROW_FIELDS) {
            pivot.addRowLabel(headers.get(field));
        }
        int monthIndex = headers.get(AuxiliarSheet.MONTH_COLUMN);
        pivot.addColLabel(monthIndex, AuxiliarSheet.MONTH_FORMAT);
        pivot.addColumnLabel(DataConsolidateFunction.SUM, headers.get(VALUE_FIELD), VALUE_LABEL,
                ReportStyles.INTEGER_FORMAT);

        configureLayout(pivot, headers, monthIndex);
    }

    /**
     * Settings POI's high-level API does not expose:
     * <ul>
     *   <li>tabular layout without subtotals, like the reference;</li>
     *   <li>ascending sort on every row field and on the months;</li>
     *   <li>one placeholder item per field: POI writes one item per source row, which
     *       does not match the real distinct values. Excel rebuilds the items when it
     *       refreshes the pivot on load.</li>
     * </ul>
     */
    private static void configureLayout(XSSFPivotTable pivot, Map<String, Integer> headers,
                                        int monthIndex) {
        CTPivotTableDefinition definition = pivot.getCTPivotTableDefinition();
        definition.setOutline(true);
        definition.setOutlineData(true);
        CTPivotFields fields = definition.getPivotFields();

        for (String field : ROW_FIELDS) {
            fields.getPivotFieldArray(headers.get(field)).setSortType(STFieldSortType.ASCENDING);
        }
        fields.getPivotFieldArray(monthIndex).setSortType(STFieldSortType.ASCENDING);

        for (String field : TABULAR_FIELDS) {
            CTPivotField pivotField = fields.getPivotFieldArray(headers.get(field));
            pivotField.setOutline(false);
            pivotField.setDefaultSubtotal(false);
            pivotField.unsetItems();
        }
        String lastRowField = ROW_FIELDS.getLast();
        for (int index : List.of(headers.get(lastRowField), monthIndex)) {
            CTItems items = fields.getPivotFieldArray(index).getItems();
            while (items.sizeOfItemArray() > 0) {
                items.removeItem(0);
            }
            items.addNewItem().setT(STItemType.DEFAULT);
            items.setCount(1);
        }
    }
}
