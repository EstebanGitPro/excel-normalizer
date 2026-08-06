#!/bin/sh
# Builds the macOS application bundle (double-clickable .app) into dist/.
# Run from the project root. Requires JDK 25 (jpackage included).
set -e

mvn -q clean package -DskipTests

jpackage \
  --type app-image \
  --name "Excel Normalizer" \
  --app-version 0.1.0 \
  --vendor "Excel Normalizer" \
  --input target \
  --main-jar excel-normalizer-0.1.0.jar \
  --main-class com.excelnormalizer.ui.Launcher \
  --java-options "--enable-native-access=ALL-UNNAMED" \
  --dest dist

echo "Done: dist/Excel Normalizer.app"
