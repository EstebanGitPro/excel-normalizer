package com.excelnormalizer.ui;

/**
 * Entry point for the runnable jar. JavaFX refuses to start when the manifest
 * main class extends Application and the runtime modules are shaded, so this
 * plain class delegates to the real application.
 */
public final class Launcher {

    private Launcher() {
    }

    public static void main(String[] args) {
        NormalizerApp.main(args);
    }
}
