package com.excelnormalizer;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

/**
 * Command-line adapter for the normalizer. All the real work happens in
 * {@link NormalizationService}; this class only parses arguments, reports
 * results on the console and maps failures to exit codes.
 */
public class ExcelNormalizer {

    private static final String DEFAULT_OUTPUT_NAME = "Informes_Contables_NORMALIZADO.xlsx";

    public static void main(String[] args) throws IOException {
        if (args.length < 3) {
            System.err.println("Usage: excel-normalizer <libro_mayor.xlsx> <saldo_terceros.xlsx>"
                    + " <movimientos.xlsx> [output.xlsx]  (input order does not matter)");
            System.exit(1);
        }
        List<Path> inputs = List.of(Path.of(args[0]), Path.of(args[1]), Path.of(args[2]));
        Path output = args.length > 3
                ? Path.of(args[3])
                : inputs.get(0).toAbsolutePath().getParent().resolve(DEFAULT_OUTPUT_NAME);

        try {
            NormalizationService.Result result =
                    new NormalizationService().normalize(inputs, output);
            System.out.println("MAYOR: " + result.mayorRows()
                    + " rows, SALDO TERCERO: " + result.saldoTerceroRows()
                    + " rows, AUXILIAR: " + result.auxiliarRows()
                    + " rows -> " + result.output());
        } catch (IllegalArgumentException e) {
            System.err.println(e.getMessage());
            System.exit(2);
        }
    }
}
