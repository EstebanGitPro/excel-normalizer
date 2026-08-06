@echo off
REM Builds the Windows installer (.msi) into dist\.
REM Run from the project root ON A WINDOWS MACHINE. Requirements:
REM   - JDK 25 installed (includes jpackage)
REM   - Maven installed
REM   - WiX Toolset v3 on PATH (jpackage needs it for .msi/.exe)
REM The Maven build resolves the Windows JavaFX natives automatically.

call mvn -q clean package -DskipTests || exit /b 1

jpackage ^
  --type msi ^
  --name "Excel Normalizer" ^
  --app-version 0.1.0 ^
  --vendor "Excel Normalizer" ^
  --input target ^
  --main-jar excel-normalizer-0.1.0.jar ^
  --main-class com.excelnormalizer.ui.Launcher ^
  --java-options "--enable-native-access=ALL-UNNAMED" ^
  --icon assets\icon\icon.ico ^
  --win-menu ^
  --win-shortcut ^
  --win-per-user-install ^
  --dest dist

echo Done: dist\Excel Normalizer-0.1.0.msi
