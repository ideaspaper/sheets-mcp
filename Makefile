.PHONY: build clean install test run help

BINARY_NAME=sheets-mcp
INSTALL_PATH=$(HOME)/.local/bin

help:
	@echo "Available targets:"
	@echo "  build    - Build the binary"
	@echo "  clean    - Remove build artifacts"
	@echo "  install  - Install to $(INSTALL_PATH)"
	@echo "  test     - Run tests"
	@echo "  run      - Run the server"

build:
	go build -o $(BINARY_NAME) .

clean:
	rm -f $(BINARY_NAME)
	go clean

install: build
	mkdir -p $(INSTALL_PATH)
	cp $(BINARY_NAME) $(INSTALL_PATH)/

test:
	go test -v ./...

run: build
	./$(BINARY_NAME)

.DEFAULT_GOAL := help
