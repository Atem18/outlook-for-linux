.PHONY: install start wayland deploy uninstall build build-rpm help

help:
	@echo "outlook-for-linux — Outlook Web client packaged as a native desktop app for Linux"
	@echo ""
	@echo "  make install     Install npm dependencies"
	@echo "  make start       Run the app (dev, auto-detects Wayland/X11)"
	@echo "  make wayland     Run the app forcing Wayland"
	@echo "  make build       Build AppImage + deb + rpm packages"
	@echo "  make deploy      Run the full installer (→ ~/.local/)"
	@echo "  make uninstall   Remove the installed app"

install:
	npm install

start: install
	npm start

wayland: install
	ELECTRON_OZONE_PLATFORM_HINT=wayland npm start

build: install
	npm run build

deploy:
	chmod +x deploy/install.sh
	./deploy/install.sh

uninstall:
	chmod +x deploy/install.sh
	./deploy/install.sh --uninstall
