#!/bin/sh
# zypak (from the Electron BaseApp) lets Chromium's sandbox run inside Flatpak.
exec zypak-wrapper /app/main/outlook-for-linux "$@"
