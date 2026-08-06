package com.excelnormalizer.ui;

import com.excelnormalizer.SourceKind;
import javafx.geometry.Insets;
import javafx.geometry.Pos;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.layout.HBox;
import javafx.scene.layout.Priority;

import java.nio.file.Path;
import java.util.function.Consumer;

/**
 * One input row of the main window: shows which export it expects,
 * the chosen file, and its validation state.
 */
final class FileSlot extends HBox {

    private final SourceKind kind;
    private final Label fileLabel = new Label("Arrastre el archivo aquí o use el botón");
    private final Label statusLabel = new Label();
    private Path file;

    FileSlot(int position, SourceKind kind, Consumer<FileSlot> onSelectRequest) {
        super(12);
        this.kind = kind;
        setAlignment(Pos.CENTER_LEFT);
        setPadding(new Insets(10));
        setStyle("-fx-background-color: -fx-control-inner-background;"
                + " -fx-background-radius: 8; -fx-border-color: #c0c0c0;"
                + " -fx-border-radius: 8; -fx-border-style: dashed;");

        Label title = new Label(position + ". " + kind.label());
        title.setStyle("-fx-font-weight: bold;");
        title.setMinWidth(220);

        fileLabel.setStyle("-fx-text-fill: #808080;");
        HBox.setHgrow(fileLabel, Priority.ALWAYS);
        fileLabel.setMaxWidth(Double.MAX_VALUE);

        Button selectButton = new Button("Seleccionar archivo...");
        selectButton.setOnAction(e -> onSelectRequest.accept(this));

        statusLabel.setMinWidth(24);
        getChildren().addAll(title, fileLabel, statusLabel, selectButton);
    }

    SourceKind kind() {
        return kind;
    }

    Path file() {
        return file;
    }

    void accept(Path path) {
        this.file = path;
        fileLabel.setText(path.getFileName().toString());
        fileLabel.setStyle("-fx-text-fill: -fx-text-inner-color;");
        statusLabel.setText("✓");
        statusLabel.setStyle("-fx-text-fill: green; -fx-font-weight: bold;");
    }

    void reject(String message) {
        this.file = null;
        fileLabel.setText(message);
        fileLabel.setStyle("-fx-text-fill: #c0392b;");
        statusLabel.setText("✗");
        statusLabel.setStyle("-fx-text-fill: #c0392b; -fx-font-weight: bold;");
    }

    void checking(Path path) {
        fileLabel.setText("Verificando " + path.getFileName() + "...");
        fileLabel.setStyle("-fx-text-fill: #808080;");
        statusLabel.setText("");
    }
}
