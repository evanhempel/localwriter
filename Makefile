# Makefile for localwriter development

UNOPKG := $(shell which unopkg 2>/dev/null || echo "/usr/lib/libreoffice/program/unopkg")
LIBREOFFICE := $(shell which libreoffice 2>/dev/null || echo "/usr/lib/libreoffice/program/soffice")

.PHONY: install-deps dev-install dev-refresh dev-run package clean

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
