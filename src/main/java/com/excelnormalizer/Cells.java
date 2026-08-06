package com.excelnormalizer;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.LinkedHashMap;
import java.util.Map;

/** Small helpers shared by the sheet writers. */
final class Cells {

    private Cells() {
    }

    /** Maps header names (first row of the sheet) to their column index. */
    static Map<String, Integer> headerIndex(Sheet sheet) {
        Map<String, Integer> index = new LinkedHashMap<>();
        Row header = sheet.getRow(sheet.getFirstRowNum());
        for (Cell cell : header) {
            if (cell.getCellType() == CellType.STRING) {
                index.put(cell.getStringCellValue().trim(), cell.getColumnIndex());
            }
        }
        return index;
    }

    static String stringValue(Cell cell) {
        if (cell == null) {
            return null;
        }
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue();
        }
        if (cell.getCellType() == CellType.NUMERIC) {
            return String.valueOf((long) cell.getNumericCellValue());
        }
        return null;
    }

    static void copyValue(Cell source, Cell target) {
        switch (source.getCellType()) {
            case NUMERIC -> target.setCellValue(source.getNumericCellValue());
            case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
            case FORMULA -> copyFormulaResult(source, target);
            case BLANK -> { }
            default -> target.setCellValue(source.getStringCellValue());
        }
    }

    private static void copyFormulaResult(Cell source, Cell target) {
        switch (source.getCachedFormulaResultType()) {
            case NUMERIC -> target.setCellValue(source.getNumericCellValue());
            case BOOLEAN -> target.setCellValue(source.getBooleanCellValue());
            default -> target.setCellValue(source.getStringCellValue());
        }
    }
}
