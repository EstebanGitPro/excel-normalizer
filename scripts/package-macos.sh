#!/bin/sh
# Builds the macOS application bundle (double-clickable .app) into dist/.
# Run from the project root. Requires JDK 25 (jpackage included).
set -e

mvn -q clean package -DskipTests

# jpackage's app-image build (and its ad-hoc codesign step) runs in a scratch
# directory outside ~/Documents. On this machine ~/Documents is synced by
# iCloud Drive, and the sync daemon tags freshly created files with
# com.apple.FinderInfo / com.apple.fileprovider.* xattrs in real time — racing
# codesign, which refuses to sign a bundle carrying those attributes. Building
# outside the synced tree avoids the race entirely.
BUILD_DIR=$(mktemp -d /private/tmp/excel-normalizer-macos-build.XXXXXX)
trap 'rm -rf "$BUILD_DIR"' EXIT

jpackage \
  --type app-image \
  --name "Excel Normalizer" \
  --app-version 1.0.0 \
  --vendor "Excel Normalizer" \
  --input target \
  --main-jar excel-normalizer-0.2.0.jar \
  --main-class com.excelnormalizer.ui.Launcher \
  --java-options "--enable-native-access=ALL-UNNAMED" \
  --icon assets/icon/icon.icns \
  --dest "$BUILD_DIR"

APP="$BUILD_DIR/Excel Normalizer.app"
codesign --deep --force -s - "$APP"
codesign --verify --deep --strict "$APP"

rm -rf dist
mkdir -p dist
cp -R "$APP" dist/

echo "Done: dist/Excel Normalizer.app"
