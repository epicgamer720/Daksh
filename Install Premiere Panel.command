#!/bin/bash
# Installs the Clip Chopper Labels panel into Premiere Pro (macOS).
cd "$(dirname "$0")"
EXT="$HOME/Library/Application Support/Adobe/CEP/extensions/com.daksh.cliplabels"
mkdir -p "$EXT/CSXS" "$EXT/jsx"
cp premiere_panel/CSXS/manifest.xml "$EXT/CSXS/"
cp premiere_panel/index.html "$EXT/"
cp add_labels.jsx "$EXT/jsx/"
for v in 10 11 12 13 14; do defaults write com.adobe.CSXS.$v PlayerDebugMode 1; done
echo "Installed. Restart Premiere, then: Window > Extensions (Legacy) > Clip Chopper Labels"
