package com.excelnormalizer.ui;

import com.excelnormalizer.NormalizationService;
import com.excelnormalizer.SourceKind;
import javafx.application.Application;
import javafx.concurrent.Task;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Hyperlink;
import javafx.scene.control.Label;
import javafx.scene.control.ProgressIndicator;
import javafx.scene.input.DragEvent;
import javafx.scene.input.TransferMode;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.List;
import java.util.Map;

/**
 * Main window: three ordered file slots (validated by header detection),
 * a convert button, and a result line with a link to the output folder.
 * All Excel work runs on background tasks so the window never freezes.
 */
public class NormalizerApp extends Application {

    private final NormalizationService service = new NormalizationService();
    private final Map<SourceKind, FileSlot> slots = new EnumMap<>(SourceKind.class);
    private final Button convertButton = new Button("Convertir");
    private final ProgressIndicator progress = new ProgressIndicator();
    private final Label resultLabel = new Label();
    private final Hyperlink openFolderLink = new Hyperlink("Abrir carpeta");
    private Stage stage;

    @Override
    public void start(Stage stage) {
        this.stage = stage;

        Label heading = new Label("Normalizador de Informes Contables");
        heading.setStyle("-fx-font-size: 18px; -fx-font-weight: bold;");
        Label subtitle = new Label(
                "Cargue los tres archivos exportados del sistema contable, en cualquier orden.");

        VBox slotsBox = new VBox(10);
        int position = 1;
        for (SourceKind kind : SourceKind.values()) {
            FileSlot slot = new FileSlot(position++, kind, this::chooseFile);
            slots.put(kind, slot);
            slotsBox.getChildren().add(slot);
        }

        convertButton.setDefaultButton(true);
        convertButton.setDisable(true);
        convertButton.setStyle("-fx-font-size: 14px; -fx-padding: 8 32;");
        convertButton.setOnAction(e -> convert());

        progress.setVisible(false);
        progress.setPrefSize(28, 28);

        openFolderLink.setVisible(false);
        HBox resultRow = new HBox(8, resultLabel, openFolderLink);
        resultRow.setAlignment(Pos.CENTER_LEFT);

        HBox actionRow = new HBox(12, convertButton, progress);
        actionRow.setAlignment(Pos.CENTER_LEFT);

        VBox root = new VBox(14, heading, subtitle, slotsBox, actionRow, resultRow);
        root.setPadding(new Insets(20));

        Scene scene = new Scene(root, 680, 360);
        enableWindowDrop(scene);
        stage.setTitle("Excel Normalizer");
        stage.setScene(scene);
        stage.show();
    }

    /** Accepts .xlsx files dropped anywhere on the window and routes them by detection. */
    private void enableWindowDrop(Scene scene) {
        scene.setOnDragOver((DragEvent e) -> {
            if (e.getDragboard().hasFiles()) {
                e.acceptTransferModes(TransferMode.COPY);
            }
            e.consume();
        });
        scene.setOnDragDropped((DragEvent e) -> {
            List<File> files = e.getDragboard().getFiles();
            if (files != null) {
                for (File file : files) {
                    assignByDetection(file.toPath());
                }
            }
            e.setDropCompleted(true);
            e.consume();
        });
    }

    private void chooseFile(FileSlot slot) {
        FileChooser chooser = new FileChooser();
        chooser.setTitle("Seleccionar " + slot.kind().label());
        chooser.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("Archivos de Excel", "*.xlsx", "*.xls"));
        File file = chooser.showOpenDialog(stage);
        if (file != null) {
            validateAndAssign(slot, file.toPath());
        }
    }

    /** Validates that the chosen file really is the export this slot expects. */
    private void validateAndAssign(FileSlot slot, Path path) {
        slot.checking(path);
        runDetection(path, kind -> {
            if (kind == slot.kind()) {
                slot.accept(path);
            } else {
                slot.reject("Este archivo parece ser \"" + kind.label()
                        + "\", no \"" + slot.kind().label() + "\"");
            }
            refreshConvertButton();
        }, message -> {
            slot.reject(message);
            refreshConvertButton();
        });
    }

    /** For window-level drops: sends the file to whichever slot matches its headers. */
    private void assignByDetection(Path path) {
        runDetection(path, kind -> {
            slots.get(kind).accept(path);
            refreshConvertButton();
        }, message -> {
            resultLabel.setText(path.getFileName() + ": " + message);
            resultLabel.setStyle("-fx-text-fill: #c0392b;");
        });
    }

    private void runDetection(Path path, java.util.function.Consumer<SourceKind> onDetected,
                              java.util.function.Consumer<String> onError) {
        Task<SourceKind> task = new Task<>() {
            @Override
            protected SourceKind call() throws Exception {
                return NormalizationService.detect(path);
            }
        };
        task.setOnSucceeded(e -> onDetected.accept(task.getValue()));
        task.setOnFailed(e -> onError.accept(friendlyMessage(task.getException())));
        Thread.ofVirtual().start(task);
    }

    private void refreshConvertButton() {
        boolean ready = slots.values().stream().allMatch(s -> s.file() != null);
        convertButton.setDisable(!ready);
    }

    private void convert() {
        FileChooser chooser = new FileChooser();
        chooser.setTitle("Guardar informe normalizado");
        chooser.setInitialFileName("Informes_Contables_NORMALIZADO.xlsx");
        chooser.getExtensionFilters().add(
                new FileChooser.ExtensionFilter("Archivo de Excel", "*.xlsx"));
        File outputFile = chooser.showSaveDialog(stage);
        if (outputFile == null) {
            return;
        }
        Path output = outputFile.toPath();

        List<Path> inputs = new ArrayList<>();
        for (FileSlot slot : slots.values()) {
            inputs.add(slot.file());
        }

        convertButton.setDisable(true);
        progress.setVisible(true);
        resultLabel.setText("Generando informe...");
        resultLabel.setStyle("");
        openFolderLink.setVisible(false);

        Task<NormalizationService.Result> task = new Task<>() {
            @Override
            protected NormalizationService.Result call() throws Exception {
                return service.normalize(inputs, output);
            }
        };
        task.setOnSucceeded(e -> {
            NormalizationService.Result result = task.getValue();
            progress.setVisible(false);
            convertButton.setDisable(false);
            resultLabel.setText("Listo: MAYOR " + result.mayorRows()
                    + " filas, SALDO TERCERO " + result.saldoTerceroRows()
                    + " filas, AUXILIAR " + result.auxiliarRows() + " filas.");
            resultLabel.setStyle("-fx-text-fill: green;");
            openFolderLink.setVisible(true);
            openFolderLink.setOnAction(a -> getHostServices()
                    .showDocument(result.output().getParent().toUri().toString()));
        });
        task.setOnFailed(e -> {
            progress.setVisible(false);
            convertButton.setDisable(false);
            resultLabel.setText("Error: " + friendlyMessage(task.getException()));
            resultLabel.setStyle("-fx-text-fill: #c0392b;");
        });
        Thread.ofVirtual().start(task);
    }

    private static String friendlyMessage(Throwable error) {
        String message = error.getMessage();
        return message == null || message.isBlank() ? error.getClass().getSimpleName() : message;
    }

    public static void main(String[] args) {
        launch(args);
    }
}
