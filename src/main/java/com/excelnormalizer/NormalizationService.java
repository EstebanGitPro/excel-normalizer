package com.excelnormalizer;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collection;
import java.util.EnumMap;
import java.util.List;
import java.util.Map;

/**
 * Core use case: turn the three raw accounting exports into the client's
 * report workbook. Knows nothing about how it is invoked (CLI, desktop UI),
 * so it reports failures with exceptions, never by exiting the process.
 */
public class NormalizationService {

    /** Outcome of a successful normalization, for whoever needs to display it. */
    public record Result(Path output, int mayorRows, int saldoTerceroRows, int auxiliarRows) {}

    /**
     * Identifies which export a file is by reading its header row.
     *
     * @throws IllegalArgumentException if the file matches no known export layout
     */
    public static SourceKind detect(Path file) throws IOException {
        try (InputStream in = Files.newInputStream(file);
             Workbook workbook = WorkbookFactory.create(in)) {
            return detect(workbook.getSheetAt(0));
        }
    }

    private static SourceKind detect(Sheet sheet) {
        Map<String, Integer> headers = Cells.headerIndex(sheet);
        for (SourceKind kind : SourceKind.values()) {
            if (kind.matches(headers)) {
                return kind;
            }
        }
        throw new IllegalArgumentException(
                "Unrecognized export: sheet '" + sheet.getSheetName()
                        + "' does not match any known header layout");
    }

    /**
     * Normalizes the three exports (any order) into the report workbook.
     *
     * @throws IllegalArgumentException if inputs are missing, duplicated or malformed
     */
    public Result normalize(Collection<Path> inputs, Path output) throws IOException {
        List<Workbook> workbooks = new ArrayList<>();
        try (XSSFWorkbook target = new XSSFWorkbook()) {
            Map<SourceKind, Sheet> sources = new EnumMap<>(SourceKind.class);
            for (Path input : inputs) {
                InputStream in = Files.newInputStream(input);
                Workbook workbook = WorkbookFactory.create(in);
                workbooks.add(workbook);
                Sheet sheet = workbook.getSheetAt(0);
                SourceKind kind = detect(sheet);
                if (sources.putIfAbsent(kind, sheet) != null) {
                    throw new IllegalArgumentException(
                            "Two input files look like the same export ("
                                    + kind.label() + "): " + input.getFileName());
                }
            }
            for (SourceKind kind : SourceKind.values()) {
                if (!sources.containsKey(kind)) {
                    throw new IllegalArgumentException(
                            "Missing input file for " + kind.label());
                }
            }

            ReportStyles styles = new ReportStyles(target);
            new MayorSheet(styles).write(sources.get(SourceKind.LIBRO_MAYOR), target);
            new SaldoTerceroSheet(styles).write(sources.get(SourceKind.SALDO_TERCEROS), target);
            new AuxiliarSheet(styles).write(sources.get(SourceKind.MOVIMIENTOS), target);

            try (FileOutputStream out = new FileOutputStream(output.toFile())) {
                target.write(out);
            }
            return new Result(output.toAbsolutePath(),
                    target.getSheet("MAYOR").getLastRowNum() + 1,
                    target.getSheet("SALDO TERCERO").getLastRowNum() + 1,
                    target.getSheet("AUXILIAR").getLastRowNum() + 1);
        } finally {
            for (Workbook workbook : workbooks) {
                workbook.close();
            }
        }
    }
}
