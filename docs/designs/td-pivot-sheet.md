# Hoja TD: tabla dinámica nativa sobre AUXILIAR

*   **Author:** EstebanGitPro
*   **Status:** Implemented (pending Windows check)
*   **Date:** 2026-11-12

--------------------------------------------------------------------------------

## Overview

excel-normalizer es una aplicación de escritorio (Java 25, JavaFX, Apache POI 5.4.1). Toma tres exportaciones del programa contable y arma un libro con tres hojas:

| Archivo de origen (nombre en el código) | Archivo real | Hoja que genera |
| --- | --- | --- |
| Libro Mayor | export con `CODCTA` y `Tnomcta` | MAYOR |
| Saldo de Terceros | export con `Código Cuenta` y `NIT_CEDULA` | SALDO TERCERO |
| Movimiento Contabilidad | `frm_con_Movimiento_Contabilidad.xlsx` | AUXILIAR |

Este diseño agrega una cuarta hoja, **TD**: una tabla dinámica de Excel construida sobre AUXILIAR.

## The Problem

`docs/COMO-FUNCIONA.md` (líneas 111 y 236-240) registra que la TD **no** se generaba a propósito, por YAGNI: el contador tomaba su propia tabla dinámica del mes anterior y la apuntaba al archivo nuevo.

**Qué cambió:** el 2026-11-12 el cliente pidió que la TD venga incluida en el informe generado. La necesidad ya no es una suposición, así que la decisión YAGNI deja de aplicar.

Los contadores usan la TD de forma interactiva:

- expanden las cuentas desde N1 hasta el tercero (cNOMNIT);
- filtran;
- hacen doble clic en un valor para que Excel cree una hoja de detalle (`Detalle1`) con las filas de origen que forman ese valor.

Por eso un resumen con números fijos no sirve: tiene que ser una tabla dinámica real, con su caché. El caché es la copia de los datos de origen que Excel guarda dentro de la tabla dinámica; el doble clic y los filtros trabajan sobre esa copia.

Hay dos complicaciones:

- **Los datos crecen.** AUXILIAR trae el año acumulado. El informe que se genera en noviembre trae datos hasta octubre. La tabla tiene que mostrar los meses nuevos sin cambios en el código.
- **Hoy las fórmulas se guardan sin su resultado.** N1, N2 y N3 (`LEFT` sobre codcta) y NETO (DEBITO − CREDITO) se guardan como fórmulas sin valor calculado. A confirmar en OQ1: si Excel arma la tabla dinámica antes de calcularlas, las filas pueden salir vacías.

### Columnas de AUXILIAR que usa la TD

| Columna | Encabezado | Qué es |
| --- | --- | --- |
| A, B, C | N1, N2, N3 | Primeros 1, 2 y 4 dígitos de la cuenta (clase, grupo, cuenta). Fórmulas `LEFT` |
| D | codcta | Código de la subcuenta |
| E | nomcta | Nombre de la cuenta |
| F | FECHA | Fecha del movimiento |
| I, J | DEBITO, CREDITO | Valores del movimiento |
| M | cNOMNIT | Nombre del tercero |
| X | NETO | DEBITO − CREDITO. Fórmula |
| Y | Meses | **Nueva.** Primer día del mes de FECHA |

### Libro de referencia

`Informes_Contables_Septiembre_2026.xlsx` es el informe que hoy arma el contador a mano, con la TD incluida. Está en la raíz del repo y `.gitignore` lo excluye porque tiene datos reales del cliente. Su TD tiene:

- origen `AUXILIAR!A1:X1223` (1.222 filas de datos, enero a septiembre);
- filas N1 → N2 → N3 → nomcta → cNOMNIT;
- columnas: FECHA agrupada por meses con el campo "Meses";
- valores: Suma de NETO.

El export real `frm_con_Movimiento_Contabilidad.xlsx` (también fuera del repo) trae 4.577 filas de enero a abril. Es el tamaño que se usa para probar OQ4.

## Goals

*   **G1. Hoja TD generada.** Es la última hoja del libro y contiene una tabla dinámica nativa con:
    *   **Filas:** N1 → N2 → N3 → nomcta → cNOMNIT.
    *   **Columnas:** Meses, más Total general.
    *   **Valores:** Suma de NETO.
*   **G2. Meses en orden de calendario,** con etiquetas en español que incluyen el año (ene-26, feb-26, mar-26...). Si un archivo trae dos años, ene-25 y ene-26 quedan en columnas separadas.
*   **G3. Meses nuevos sin cambios en el código.** Si AUXILIAR trae octubre, la TD muestra oct-26.
*   **G4. El doble clic funciona desde la primera apertura.** Un doble clic en cualquier valor crea la hoja de detalle con exactamente las filas de AUXILIAR que forman ese valor.
*   **G5. Los números coinciden con AUXILIAR.**
    *   El total de cada mes es igual a la suma de NETO de las filas de ese mes.
    *   El Total general es igual a la suma de la columna NETO.
*   **G6 (condicional, depende de OQ1). Fórmulas con su valor calculado.** Si el spike muestra que la TD sale vacía sin los valores calculados, el libro se guarda con el resultado de cada fórmula. Si no hace falta, G6 se elimina.
*   **G7. Las fechas inválidas no pasan en silencio.** Si alguna FECHA del archivo Movimiento Contabilidad no es una fecha válida, la conversión se cancela y el error indica las filas del archivo de origen (hasta 20). No se genera ningún archivo.

## Non-Goals

*   Compatibilidad con LibreOffice, Google Sheets, Excel web, Excel en el celular o vistas previas de correo. Supuesto informado por el autor: todos los usuarios usan Excel de escritorio en Windows.
*   Expandir un mes para verlo partido en días dentro de la TD. Los contadores ven los días en la hoja de detalle que crea Excel.
*   Tablas dinámicas o gráficos adicionales.
*   Cambiar las columnas A a X de AUXILIAR (orden, nombres o formato).
*   Copiar el diseño visual de la TD de referencia (colores, estilo de tabla). Se usa el estilo por defecto de Excel.
*   Dejar la TD con las filas N1 colapsadas al abrir. Excel la abre expandida o como la deje la actualización.
*   Archivos de origen con el sistema de fechas 1904. Los exports del programa contable usan el sistema 1900.

## Resolved Questions

*   **OQ3. Formato de los meses:** siempre `mmm-yy` (ene-26), aunque el informe suele ser de un solo año. Así las validaciones y los mensajes de error no son ambiguos entre años. Decisión del autor.
*   **OQ6. FECHA inválida:** cancelar la conversión (ver Alternativa F). Decisión del autor.
*   **¿Generar la TD o no?:** generarla, por pedido del cliente (2026-11-12). Ver Alternativa 0.
*   **¿Hay filas vacías o de totales en el export?:** no. Se revisó `frm_con_Movimiento_Contabilidad.xlsx`: 4.577 filas, todas con codcta y con FECHA como fecha real. Los totales vienen como columnas repetidas en cada fila (`TOTAL_DEBITOS`, `TOTAL_CREDITOS`), no como filas. Por eso se valida FECHA en todas las filas, sin regla para saltar filas.
*   **AUXILIAR sin filas de datos:** la conversión se cancela con un error (decisión del autor). Una tabla dinámica sin datos no le sirve a nadie y puede hacer que Excel pida reparar el archivo.

## Open Questions

Las cuatro se responden con el spike (paso 0 del plan), en Excel de escritorio en Windows, usando el export real de 4.577 filas.

**Resultado del spike en Excel para Mac (2026-11-12):** la TD abrió con valores correctos, meses ordenados y filas ordenadas, con la variante 3 de OQ1 y el plan B de OQ4 (un ítem por campo). La implementación usa esa configuración. Queda confirmar OQ4 (sin aviso de reparación) y la Vista protegida en Windows con el `.msi`.

*   **OQ1:** ¿la TD sale bien al abrir? Se prueban tres variantes y se queda la más barata que funcione:
    1.  sin hacer nada con las fórmulas;
    2.  con `setForceFormulaRecalculation(true)` (Excel calcula todo al abrir);
    3.  con `XSSFFormulaEvaluator.evaluateAllFormulaCells` (Java guarda los resultados).
*   **OQ2:** las etiquetas de la columna Meses, ¿se ven como "ene-26" (formato `[$-240A]mmm-yy`) o como fechas completas? Lo mismo para el título de la hoja de detalle ("Meses: feb-26").
*   **OQ4:** una tabla dinámica creada con POI, con 5 campos de fila y 4.577 filas, ¿abre sin que Excel pida reparar el archivo?
*   **OQ5:** ¿el orden ascendente de los meses se mantiene después de que Excel actualiza la tabla al abrir?

## Critical User Journeys (CUJs)

### CUJ 1: Generar el informe del mes con la TD incluida

*   **Given**: en octubre de 2026, un empleado tiene los tres archivos de origen con datos de enero a septiembre, y abre la aplicación en Windows.
*   **When**: carga los tres archivos y pulsa convertir.
*   **Then**: el libro guardado tiene las hojas MAYOR, SALDO TERCERO, AUXILIAR y TD, en ese orden. Al abrirlo en Excel, TD muestra las columnas ene-26 a sep-26, en ese orden, más Total general. No aparece ningún aviso de reparación.

### CUJ 2: Investigar un valor hasta el tercero y ver su detalle

*   **Given**: el libro del CUJ 1 está abierto en Excel.
*   **When**: el contador expande N1 = 1 hasta cNOMNIT y hace doble clic en el valor de feb-26 de una fila.
*   **Then**: Excel crea una hoja de detalle con exactamente las filas de AUXILIAR de esa combinación y de feb-26, con todas sus columnas, incluida Y "Meses". La suma de NETO de esa hoja es igual al valor en el que se hizo clic.

### CUJ 3: Un mes nuevo aparece solo

*   **Given**: en noviembre, los archivos de origen traen datos de enero a octubre.
*   **When**: el empleado genera el informe y lo abre.
*   **Then**: TD muestra las columnas ene-26 a oct-26 en orden de calendario, sin cambios en el código ni en la configuración.

### CUJ 4: Una FECHA inválida en el origen

*   **Given**: las filas 57 y 302 del archivo Movimiento Contabilidad tienen FECHA como texto.
*   **When**: el empleado convierte.
*   **Then**: la conversión se cancela, no se escribe ningún archivo y la aplicación muestra: `FECHA is not a valid date in Movimiento Contabilidad file, rows: 57, 302`.

--------------------------------------------------------------------------------

## Technical Plan (Architecture)

### Flujo de datos

```text
 3 archivos de origen
        │
        ▼
 NormalizationService.normalize()
        │
        ├─► MayorSheet.write()          → hoja MAYOR
        ├─► SaldoTerceroSheet.write()   → hoja SALDO TERCERO
        ├─► AuxiliarSheet.write()       → hoja AUXILIAR A..X + NUEVA columna Y "Meses"
        │        └─ valida FECHA de todas las filas; si alguna falla → IllegalArgumentException
        │
        ├─► (solo si OQ1 lo exige) calcular o marcar las fórmulas   (G6)
        │
        ├─► NUEVO: TdPivotSheet.write(target)  → hoja TD
        │        └─ tabla dinámica con origen AUXILIAR!A1:{último encabezado}{última fila}
        │
        └─► target.write(out)            → archivo .xlsx
                                             │
                                             ▼
                           Excel de escritorio al abrir:
                           refreshOnLoad → actualiza el caché y dibuja la TD
```

En este documento, **calcular** se refiere a las fórmulas de las celdas y **actualizar** al caché de la tabla dinámica.

### Componente 1: columna Y "Meses" en AUXILIAR

`AuxiliarSheet` agrega una columna después de NETO:

*   **Encabezado:** "Meses", con el mismo estilo que NETO. Coincide con el nombre del campo en la TD de referencia, así el título de la hoja de detalle sigue diciendo "Meses: ...".
*   **Valor:** el primer día del mes de la FECHA de la fila, como fecha real de Excel, sin hora. Por ejemplo, `2026-05-31 10:37` pasa a ser `2026-05-01`. Lo escribe Java como valor. Si OQ1 muestra que G6 hace falta de todos modos, se puede revisar si conviene escribirlo como fórmula (ver Alternativa E).
*   **Formato:** `[$-240A]mmm-yy`. El código de idioma `240A` (español de Colombia) fija el idioma del nombre del mes. Como el valor es un número de fecha, los meses se ordenan por calendario y cada año queda en su propia columna.
*   **Validación (G7):** la FECHA de origen es válida si la celda es numérica (o una fórmula con resultado numérico) y su valor está entre 1 (1900-01-01) y 2958465 (9999-12-31). Así se rechazan el 0 y los números tipo `20260531`. Las celdas vacías, de texto, booleanas o con error son inválidas. Se recorren todas las filas, se juntan los números de fila del archivo de origen (la fila que ve el usuario en Excel) y, si hay alguna, se lanza una única `IllegalArgumentException`.
*   **Formato del error:** `FECHA is not a valid date in Movimiento Contabilidad file, rows: 57, 302`. Con más de 20 filas se listan las primeras 20 y se agrega ` and N more`.
*   **Sin datos:** si el archivo no tiene filas de datos, se lanza `IllegalArgumentException("Movimiento Contabilidad file has no data rows")`.

### Componente 2: fórmulas con resultado (G6, condicional)

Se implementa la variante de OQ1 que haya funcionado:

*   **Variante 1:** no se hace nada y G6 se elimina.
*   **Variante 2:** `target.setForceFormulaRecalculation(true)` antes de guardar. Excel calcula todo al abrir.
*   **Variante 3:** `XSSFFormulaEvaluator.evaluateAllFormulaCells(target)` después de escribir las tres hojas y antes de crear la TD. POI calcula `LEFT`, `LEN` y las restas, y guarda el resultado junto a cada fórmula. Las fórmulas se mantienen.

### Componente 3: `TdPivotSheet` (clase nueva)

Una clase con la misma forma que las demás hojas (`final class`, método `write(XSSFWorkbook)`). No recibe `ReportStyles`. Hace lo siguiente:

1.  Toma la hoja AUXILIAR y busca por encabezado (`Cells.headerIndex`) los campos N1, N2, N3, nomcta, cNOMNIT, Meses y NETO. Si falta alguno, lanza `IllegalArgumentException("Columns not found in AUXILIAR sheet: [...]")`, igual que las demás hojas. Una columna movida no rompe nada; una renombrada o borrada da este error.
2.  Calcula el rango de origen: desde A1 hasta la columna del último encabezado y la última fila.
3.  Crea la hoja "TD" al final y la tabla dinámica en A3 (como en la referencia), con `createPivotTable`.
4.  Filas: `addRowLabel` para N1, N2, N3, nomcta y cNOMNIT, en ese orden. Diseño tabular como en la referencia: N1, N2, N3 y nomcta llevan `outline = false` y `defaultSubtotal = false`, así cada campo ocupa su propia columna y no hay subtotales; la tabla lleva `outline = true` y `outlineData = true`. Sin esto, POI deja el diseño compacto y los cinco campos se apilan en una sola columna.
5.  Columnas: `addColLabel(índice de Meses, "[$-240A]mmm-yy")`. Este overload ya fija el formato del campo.
6.  Valores: `addColumnLabel(SUM, índice de NETO, "Suma de NETO", ReportStyles.INTEGER_FORMAT)`. Este overload ya fija el formato del valor.
7.  Orden: `sortType = ascending` en el campo Meses y en los cinco campos de fila (sin esto, Excel deja las filas en el orden en que aparecen en los datos), con `CTPivotField` y `STFieldSortType` de `poi-ooxml-lite` (están en el jar 5.4.1).

POI ya marca el caché con `refreshOnLoad = true` por defecto; el test lo verifica. No hace falta agregar dependencias.

--------------------------------------------------------------------------------

## Risks & Mitigations

*   **El caché sale vacío desde Java.** POI no escribe los registros del caché; Excel los llena al abrir (`refreshOnLoad`). Solo se soporta Excel de escritorio (ver Non-Goals).
*   **Vista protegida.** Si el archivo llega por correo, Excel lo abre en Vista protegida y no actualiza la tabla hasta que el usuario pulsa "Habilitar edición", así que TD se ve vacía hasta ese momento. Mitigación: abrir el archivo una vez y guardarlo antes de enviarlo (el caché queda lleno). Se documenta en COMO-FUNCIONA y se verifica en el spike.
*   **Excel pide guardar al cerrar.** Como la tabla se actualiza al abrir, Excel considera el archivo modificado. Es esperable.
*   **POI declara los ítems de los campos de forma aproximada (OQ4).** `addRowLabel` y `addColLabel` escriben un ítem por fila de origen en cada uno de los 6 campos, en lugar de uno por valor distinto. Excel puede pedir reparar el archivo. Plan B si OQ4 falla: en cada uno de los 6 campos, reemplazar la lista de ítems por un solo `<item t="default"/>` con `count = 1`, y que Excel cree los ítems al actualizar. El test 4 pasa a verificar ese único ítem.
*   **Formas nuevas de fallar la conversión.** Con este cambio la conversión puede cancelarse por: FECHA inválida (G7), AUXILIAR sin datos, o encabezados faltantes en AUXILIAR. Las tres usan `IllegalArgumentException`, que la CLI y la ventana ya muestran como mensaje.
*   **Rango fijo.** El rango de origen queda fijo en las filas de esa ejecución. No es un problema porque cada mes se genera un archivo nuevo. Si alguien agrega filas a mano en AUXILIAR, tiene que cambiar el origen de la tabla en Excel.
*   **Seguridad.** Sin cambios: no hay red, ni datos nuevos, ni dependencias nuevas.
*   **Rollback.** Volver atrás es revertir el commit y generar otro instalador .msi. Los archivos ya generados no se ven afectados.

--------------------------------------------------------------------------------

## Alternatives Considered

### Alternativa 0: statu quo (el contador apunta su propia TD al archivo nuevo)

*   **Pros**: cero código. Es la decisión YAGNI registrada en COMO-FUNCIONA.
*   **Cons**: cada mes alguien repite el trabajo manual y puede elegir mal el rango.
*   **Reason for Rejection**: Rejected by author: el cliente pidió que la TD venga incluida en el informe (2026-11-12).

### Alternativa A: resumen estático (valores fijos o fórmulas SUMIFS)

*   **Pros**: sin riesgo de que Excel pida reparar el archivo. Se ve igual en cualquier programa. Las fórmulas generadas también pueden sumar columnas por cada mes nuevo.
*   **Cons**: no hay doble clic al detalle, ni expandir o colapsar, ni filtros.
*   **Reason for Rejection**: Rejected by author: los contadores usan la TD de forma interactiva, y nada de eso existe sin un caché de tabla dinámica real.

### Alternativa B: plantilla .xlsx con la agrupación nativa de fechas de Excel

*   **Pros**: copia exacta de la TD de referencia, incluido el "+" que abre un mes en días. Si se agrupa por Años y Meses, los años quedan separados. No hace falta la columna Y.
*   **Cons**:
    *   Un archivo binario en el repo que no se puede revisar en un diff.
    *   Hay que reescribir el rango de origen de la plantilla en cada ejecución.
    *   Las columnas de la plantilla tienen que coincidir con las del código sin que nada lo controle.
    *   Quedan registros viejos en el caché.
    *   `NormalizationService` tendría que abrir la plantilla en lugar de crear el libro desde cero.
*   **Reason for Rejection**: Rejected by author: su única ventaja real es abrir un mes en días, y los contadores siempre trabajan con el mes completo (conversación con el autor, 2026-11-12). Los días los ven en la hoja de detalle que crea el doble clic, que funciona igual con la alternativa elegida.

### Alternativa C: tabla de POI más la agrupación por meses escrita en XML desde el código

*   **Pros**: sin plantilla, con la misma agrupación nativa que la referencia, incluido el "+" que abre días.
*   **Cons**:
    *   `poi-ooxml-lite` no trae `CTFieldGroup`, `CTRangePr` ni `CTGroupItems`. Habría que agregar `poi-ooxml-full`, un jar mucho más pesado que también agranda el instalador .msi, o armar el XML a mano.
    *   Más riesgo de que Excel pida reparar el archivo, porque la agrupación exige ítems coherentes entre el caché y la tabla.
*   **Reason for Rejection**: Rejected by author: más dependencia y más riesgo, y el "+" para días no se usa.

### Alternativa D: columna de ayuda con el mes como texto ("ene") o como número (1..12)

*   **Pros**: simple de escribir.
*   **Cons**: el texto se ordena alfabéticamente (abr, ago, dic, ene...). Un número 1..12 con formato de fecha muestra "ene" en todos los casos, porque 1 a 12 son días de enero de 1900. Ninguno de los dos distingue años.
*   **Reason for Rejection**: Rejected by author: no cumplen G2. La fecha del primer día del mes cumple orden y etiqueta a la vez.

### Alternativa E: columna Meses calculada con fórmula (`=DATE(YEAR(F2),MONTH(F2),1)`)

*   **Pros**: si alguien corrige una FECHA a mano en AUXILIAR, el mes se actualiza solo. Sigue la regla del proyecto de dejar fórmulas vivas.
*   **Cons**: depende de que la fórmula esté calculada cuando Excel actualiza la tabla, que es la misma pregunta que OQ1. La validación de FECHA (G7) se sigue haciendo en Java de todos modos.
*   **Reason for Rejection**: Pendiente de OQ1. Si G6 hace falta igual, esta alternativa no cuesta nada extra y se reconsidera. Si no hace falta, el valor escrito por Java evita depender del orden de cálculo.

### Alternativa F: ante FECHA inválida, seguir y avisar, o interpretar el texto como fecha

*   **Pros**: el informe sale igual, o se aceptan más formatos de exportación.
*   **Cons**:
    *   Si se sigue, la TD muestra una columna "(en blanco)" con esas filas, y el informe sale con datos sin mes.
    *   Si se interpreta el texto, `03/04/2026` puede quedar en marzo o en abril según el formato supuesto, sin que nadie se entere.
*   **Reason for Rejection**: Rejected by author: es preferible no generar el informe antes que generar uno con meses mal asignados (OQ6).

--------------------------------------------------------------------------------

## Non-Functional Requirements

*   **Rendimiento:** estimado, sin medir. El export real más grande visto tiene 4.577 filas en 4 meses; un año completo rondaría las 14.000. Si se usa la variante 3 de OQ1, calcular `LEFT`, `LEN` y restas sobre ese volumen debería tardar pocos segundos. El spike registra el tiempo real.
*   **Tamaño del archivo:** crece poco: un caché vacío más una columna de fechas.
*   **Compatibilidad:** solo Excel de escritorio en Windows (ver Non-Goals).

--------------------------------------------------------------------------------

## Detailed Implementation Plan

### Orden de trabajo

0.  **Spike** (antes de cualquier código de producción): en una rama descartable, generar un libro con la TD usando el export real y abrirlo en Excel de Windows. Responde OQ1, OQ2, OQ4 y OQ5, y la verificación de Vista protegida (ver Verification Plan, paso 1). Registrar los resultados en la sección Open Questions de este documento.
1.  **Tests:** JUnit y surefire en `pom.xml` (Archivo 4) y el esqueleto de `NormalizationServiceTest` con los helpers de datos de prueba (Archivo 5).
2.  **Columna Meses y validación de FECHA** (Archivo 1), con los tests 1, 6, 7 y 8.
3.  **G6 según OQ1** (Archivo 2), con el test 2 si aplica.
4.  **`TdPivotSheet`** (Archivo 3) y su llamada desde `NormalizationService` (Archivo 2), con los tests 3, 4, 5 y 9.
5.  **Prueba manual** con el informe real (Verification Plan, paso 2).
6.  **Documentación** (Archivo 6).

### Target Files to Create/Modify

#### 1. `src/main/java/com/excelnormalizer/AuxiliarSheet.java` (Modify)

*   **Type / Scope**: escritor de la hoja AUXILIAR.
*   **Proposed Changes**:
    *   Constantes package-private nuevas: `SHEET_NAME = "AUXILIAR"` (reemplaza el literal de `createSheet`), `MONTH_COLUMN = "Meses"` y `MONTH_FORMAT = "[$-240A]mmm-yy"`.
    *   `COLUMNS_TO_KEEP` pasa de `private` a package-private, para que los tests armen los datos de prueba con la misma lista.
    *   La columna Meses va en `netoIndex + 1` (columna Y).
    *   Encabezado "Meses" con el mismo estilo que NETO.
    *   En cada fila de datos: si FECHA es válida (reglas del Componente 1), escribe el primer día del mes con un estilo de fecha `MONTH_FORMAT` creado una vez con `styles.custom`. Si no es válida, guarda el número de fila del archivo de origen (índice de fila + 1).
    *   Al terminar las filas: si no hubo filas de datos, o si hubo FECHAs inválidas, lanza `IllegalArgumentException` con los mensajes del Componente 1.
    *   Agrega el ancho de la columna Meses en `setColumnWidths`.
    *   Actualiza el Javadoc de la clase: menciona la columna Meses y que la TD se alimenta de esta hoja.

#### 2. `src/main/java/com/excelnormalizer/NormalizationService.java` (Modify)

*   **Type / Scope**: caso de uso principal.
*   **Proposed Changes**:
    *   Después de escribir AUXILIAR y antes de abrir el `FileOutputStream`, en este orden:
        1.  La variante de OQ1 elegida (nada, `setForceFormulaRecalculation(true)` o `evaluateAllFormulaCells`).
        2.  `new TdPivotSheet().write(target)`.
    *   La línea 92 usa `AuxiliarSheet.SHEET_NAME` en lugar del literal "AUXILIAR".
    *   El record `Result` no cambia.
    *   Las excepciones de `AuxiliarSheet` y `TdPivotSheet` ocurren antes de abrir el `FileOutputStream`, así que no se crea ningún archivo.

#### 3. `src/main/java/com/excelnormalizer/TdPivotSheet.java` (Create)

*   **Type / Scope**: escritor de la hoja TD.
*   **Proposed Changes**: implementa el Componente 3. Constantes package-private: `SHEET_NAME = "TD"`, la lista de campos de fila, `VALUE_FIELD = "NETO"` y `VALUE_LABEL = "Suma de NETO"`. Si el spike activa el plan B de OQ4, también implementa el reemplazo de ítems descrito en Risks & Mitigations.

#### 4. `pom.xml` (Modify)

*   **Type / Scope**: build.
*   **Proposed Changes**:
    *   Dependencia `org.junit.jupiter:junit-jupiter:5.11.4` con `<scope>test</scope>`, así no entra en el jar sombreado ni en el .msi.
    *   Plugin `maven-surefire-plugin` versión `3.5.2`, para que la versión que corre JUnit 5 no dependa de la versión de Maven instalada.

#### 5. `src/test/java/com/excelnormalizer/NormalizationServiceTest.java` (Create)

*   **Type / Scope**: tests de integración del caso de uso con JUnit 5. Escriben los archivos de origen y el de salida en un `@TempDir` y vuelven a abrir la salida con POI.
*   **Datos de prueba**: un helper privado arma los tres archivos de origen con POI. No se guardan archivos .xlsx en el repo.
    *   **Libro Mayor:** encabezados `CODCTA, Tnomcta, ANTERIOR, DEBITOS, CREDITOS, SALDO_FINAL`. Siete filas, una por clase, con `CODCTA` como texto de un dígito ("1" a "7") y los importes como números. Sin las filas de clase, la fórmula PARTIDA DOBLE de MAYOR queda vacía.
    *   **Saldo de Terceros:** los 8 encabezados de `SaldoTerceroSheet` (`Código Cuenta, Nombre del Concepto Contable, NIT_CEDULA, Nombre de la Razon Social, SALDO_ANTERIOR, DEBITOS, CREDITOS, SALDO_FINAL`), con dos filas; cuenta y NIT como texto, importes como números. Para usar la lista directamente, `SaldoTerceroSheet.COLUMNS_TO_KEEP` pasa a package-private, igual que en `MayorSheet`.
    *   **Movimiento Contabilidad:** los encabezados de `AuxiliarSheet.COLUMNS_TO_KEEP`, en ese orden. `codcta`, `nomcta` y `cNOMNIT` como texto; `FECHA` como fecha; `DEBITO` y `CREDITO` como números; el resto vacío. Cada test indica sus filas.
*   **Tests** (Goal que cubre entre paréntesis):
    1.  `auxiliarHasMonthColumnWithFirstOfMonth` (G2): con FECHA `2026-05-31 10:37`, AUXILIAR Y1 = "Meses", Y2 = 2026-05-01 con formato `[$-240A]mmm-yy`.
    2.  `formulasHaveCachedValues` (G6, solo si OQ1 eligió la variante 3): con codcta "53959596", DEBITO 179000 y CREDITO 0, sin evaluar al leer, AUXILIAR A2/B2/C2 tienen en caché "5", "53" y "5395", y X2 tiene 179000. En MAYOR, A6 (LEN de la clase "1") tiene en caché 1.
    3.  `tdSheetIsLastAndHasPivot` (G1): las hojas son, en orden, MAYOR, SALDO TERCERO, AUXILIAR y TD, y TD tiene exactamente una tabla dinámica.
    4.  `pivotLayoutMatchesReference` (G1, G2): los campos de fila, en este orden, son los índices de N1, N2, N3, nomcta y cNOMNIT; el campo de columna es Meses; hay un campo de datos SUM sobre NETO llamado "Suma de NETO"; `refreshOnLoad` es true; Meses y los cinco campos de fila tienen `sortType` ascendente; N1, N2, N3 y nomcta tienen `outline` y `defaultSubtotal` en false.
    5.  `pivotSourceCoversAllRows` (G3): con 5 filas de datos, el origen es la hoja `AUXILIAR` con ref `A1:Y6`; la definición del caché tiene 25 campos.
    6.  `invalidDatesCancelWithoutWritingFile` (G7): con FECHA texto en la fila 3 del origen, vacía en la fila 5 y `0` en la fila 6, se lanza `IllegalArgumentException` con el mensaje exacto `FECHA is not a valid date in Movimiento Contabilidad file, rows: 3, 5, 6`, y el archivo de salida no existe.
    7.  `monthsOfDifferentYearsStayDistinct` (G2): con filas de 2025-12-15 y 2026-01-10, Y tiene 2025-12-01 y 2026-01-01.
    8.  `moreThanTwentyInvalidDatesAreTruncated` (G7): con 25 FECHAs inválidas, el mensaje lista 20 filas y termina en ` and 5 more`.
    9.  `emptyMovimientosCancels`: con solo la fila de encabezados, se lanza `IllegalArgumentException("Movimiento Contabilidad file has no data rows")`.
*   G4 y G5 dependen de que Excel actualice la tabla, así que solo se verifican a mano (Verification Plan, paso 2).

#### 6. `docs/COMO-FUNCIONA.md` (Modify)

*   **Type / Scope**: guía para el usuario y el desarrollador.
*   **Proposed Changes**:
    *   Línea 21 (diagrama): "3 hojas" pasa a "4 hojas".
    *   Línea 111 (nota de AUXILIAR): reemplazar "se decidió NO generarla" por la columna Y "Meses" y la referencia a la hoja TD.
    *   Líneas 121-152 (capas y responsabilidades): agregar `TdPivotSheet` al diagrama y a la tabla.
    *   Líneas 236-240 (YAGNI): quitar la TD de la lista y aclarar que se generó cuando el cliente la pidió (ejemplo de YAGNI bien aplicado: se construye cuando la necesidad es real).
    *   Sección nueva "Hoja TD": campos, que se actualiza al abrir, que solo funciona en Excel de escritorio y que hay que abrirla y guardarla una vez antes de enviarla por correo (Vista protegida).
    *   Mensajes de error nuevos (FECHA inválida, archivo sin filas) y qué hacer: corregir el archivo de origen.
    *   Si OQ1 eligió la variante 3: una nota en la tabla de fórmulas de que el archivo se guarda con los valores calculados.

#### Fuera de este cambio

Borrar los artefactos de diseño y revisión de la raíz del repo (`concepts_brief.md`, `design_draft.md`, `zoom_out_report.md`, `comprehension_report.md`, `critic_report.md`, `intermediate_zoom_out_report.md`, `implementation_readiness_report.md`, `consolidated_review_report.md`). Va en un commit `chore` aparte.

### Verification Plan

*   **Compilation/Build**: `mvn -q package` sin errores. El jar sombreado no debe incluir JUnit: `unzip -l target/excel-normalizer-0.1.0.jar | rg -i junit` no devuelve nada.
*   **Unit Tests**: `mvn test`, con los tests del Archivo 5 en verde (el test 2 solo si aplica).
*   **Static Analysis & Linting**: el proyecto no tiene. No se agrega ninguno.
*   **Mapa de Goals**: G1 → tests 3 y 4; G2 → tests 1, 4 y 7, más la parte visual a mano; G3 → test 5 y paso 2.5; G4 y G5 → solo a mano; G6 → test 2; G7 → tests 6 y 8.
*   **E2E/Integration (manual, Excel de escritorio en Windows)**:
    1.  **Spike (paso 0)**, con `frm_con_Movimiento_Contabilidad.xlsx`:
        *   OQ1: generar las tres variantes y anotar cuál es la más barata que muestra la TD con datos.
        *   OQ4: ningún aviso de reparación al abrir.
        *   OQ2: las columnas dicen ene-26 a abr-26; el título de la hoja de detalle dice "Meses: feb-26" o una fecha completa.
        *   OQ5: las columnas siguen en orden de calendario después de actualizar.
        *   Vista protegida: abrir una copia descargada (por ejemplo, enviada por correo) y confirmar que la TD se llena al pulsar "Habilitar edición".
        *   Anotar el tiempo de la conversión.
    2.  **Con la implementación terminada:**
        1.  Generar el informe con los archivos reales de septiembre 2026 y abrirlo: hojas en orden, sin aviso de reparación, columnas ene-26 a sep-26.
        2.  Comparar con `Informes_Contables_Septiembre_2026.xlsx`: los totales de cada N1 por mes y el Total general coinciden (G5).
        3.  Doble clic en un valor de feb-26: se crea la hoja de detalle y su suma de NETO es igual al valor (G4).
        4.  Actualizar la tabla, guardar, cerrar y reabrir: los totales y el orden de los meses son los mismos.
        5.  CUJ 3: agregar filas de octubre a una copia del archivo de origen, generar el informe y confirmar que aparece oct-26 al final.

--------------------------------------------------------------------------------

*Drafted with EGM*
