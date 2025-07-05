# Makefile for localwriter development

UNOPKG := $(shell which unopkg 2>/dev/null || echo "/usr/lib/libreoffice/program/unopkg")
LIBREOFFICE := $(shell which libreoffice 2>/dev/null || echo "/usr/lib/libreoffice/program/soffice")

.PHONY: help install-deps dev-install dev-refresh dev-run package clean

help:
	@echo "LocalWriter Development Makefile"
	@echo ""
	@echo "Available targets:"
	@echo "  help           - Show this help message"
	@echo "  install-deps   - Install Python dependencies in lib/"
	@echo "  dev-install    - First-time dev setup (installs deps and registers extension)"
	@echo "  dev-refresh    - Refresh extension after making changes"
	@echo "  dev-run        - Start LibreOffice Writer"
	@echo "  package        - Create localwriter.oxt package"
	@echo "  clean          - Remove built files"
	@echo ""
	@echo "Typical workflow:"
	@echo "  make dev-install   # First time setup"
	@echo "  make dev-refresh   # After making changes"
	@echo "  make dev-run       # To start LibreOffice"

# Set help as default target
.DEFAULT_GOAL := help

install-deps:
	@echo "Installing Python dependencies..."
	mkdir -p lib
	pip install litellm -t lib

dev-install: install-deps
	@echo "Installing development version..."
	$(UNOPKG) remove org.extension.sample || true
	yes yes | $(UNOPKG) add . > /dev/null

dev-refresh:
	@echo "Refreshing development installation..."
	$(UNOPKG) remove org.extension.sample
	yes yes | $(UNOPKG) add . > /dev/null

dev-run:
	@echo "Starting LibreOffice Writer..."
	$(LIBREOFFICE) --writer &

package:
	@echo "Creating localwriter.oxt package..."
	zip -r localwriter.oxt \
		Accelerators.xcu \
		Addons.xcu \
		assets \
		description.xml \
		main.py \
		META-INF \
		registration \
		README.md \
		lib

clean:
	@echo "Cleaning up..."
	rm -f localwriter.oxt
