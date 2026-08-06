package com.excelnormalizer;

import java.util.Map;

/** The three accounting exports the normalizer accepts, recognized by their header row. */
public enum SourceKind {

    LIBRO_MAYOR("Libro Mayor") {
        @Override
        boolean matches(Map<String, Integer> headers) {
            return headers.containsKey("CODCTA") && headers.containsKey("Tnomcta");
        }
    },
    SALDO_TERCEROS("Saldo de Terceros") {
        @Override
        boolean matches(Map<String, Integer> headers) {
            return headers.containsKey("Código Cuenta") && headers.containsKey("NIT_CEDULA");
        }
    },
    MOVIMIENTOS("Movimientos de Contabilidad") {
        @Override
        boolean matches(Map<String, Integer> headers) {
            return headers.containsKey("codcta") && headers.containsKey("FECHA");
        }
    };

    private final String label;

    SourceKind(String label) {
        this.label = label;
    }

    /** Human-readable name for CLI messages and UI labels. */
    public String label() {
        return label;
    }

    abstract boolean matches(Map<String, Integer> headers);
}
