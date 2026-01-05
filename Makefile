.PHONY: all build test clean install lint fmt vet cover help examples docker docker-build docker-run docker-push

# Binary name
BINARY=goxls

# Build directory
BUILD_DIR=bin

# Docker settings
DOCKER_IMAGE=goxls
DOCKER_TAG=latest

# Go parameters
GOCMD=go
GOBUILD=$(GOCMD) build
GOTEST=$(GOCMD) test
GOGET=$(GOCMD) get
GOMOD=$(GOCMD) mod
GOFMT=$(GOCMD) fmt
GOVET=$(GOCMD) vet

# Default target
all: test build

## build: Build the CLI binary
build:
	@echo "Building $(BINARY)..."
	@mkdir -p $(BUILD_DIR)
	$(GOBUILD) -o $(BUILD_DIR)/$(BINARY) ./cmd
	@echo "Binary created: $(BUILD_DIR)/$(BINARY)"

## test: Run all tests
test:
	@echo "Running tests..."
	$(GOTEST) -v ./...

## test-short: Run tests without verbose output
test-short:
	@echo "Running tests..."
	$(GOTEST) ./...

## cover: Run tests with coverage report
cover:
	@echo "Running tests with coverage..."
	$(GOTEST) -coverprofile=coverage.out ./...
	$(GOCMD) tool cover -func=coverage.out
	@echo "\nCoverage report: coverage.out"
	@echo "Run 'make cover-html' to view in browser"

## cover-html: Generate HTML coverage report and open in browser
cover-html: cover
	$(GOCMD) tool cover -html=coverage.out -o coverage.html
	@echo "Opening coverage.html..."
	@xdg-open coverage.html 2>/dev/null || open coverage.html 2>/dev/null || echo "Open coverage.html in your browser"

## fmt: Format all Go files
fmt:
	@echo "Formatting code..."
	$(GOFMT) ./...

## vet: Run go vet
vet:
	@echo "Running go vet..."
	$(GOVET) ./...

## lint: Run golangci-lint (must be installed)
lint:
	@echo "Running linter..."
	@which golangci-lint > /dev/null || (echo "Install golangci-lint: https://golangci-lint.run/usage/install/" && exit 1)
	golangci-lint run

## clean: Remove build artifacts
clean:
	@echo "Cleaning..."
	@rm -rf $(BUILD_DIR)
	@rm -f coverage.out coverage.html
	@echo "Clean complete"

## install: Install binary to GOPATH/bin
install: build
	@echo "Installing $(BINARY)..."
	@cp $(BUILD_DIR)/$(BINARY) $(GOPATH)/bin/$(BINARY)
	@echo "Installed to $(GOPATH)/bin/$(BINARY)"

## deps: Download dependencies
deps:
	@echo "Downloading dependencies..."
	$(GOMOD) download

## tidy: Tidy and verify dependencies
tidy:
	@echo "Tidying dependencies..."
	$(GOMOD) tidy
	$(GOMOD) verify

## examples: Build all examples
examples:
	@echo "Building examples..."
	@for dir in examples/*/; do \
		echo "  Building $$dir..."; \
		(cd $$dir && $(GOBUILD) -o main .); \
	done
	@echo "Examples built"

## run: Build and run the CLI with sample data
run: build
	@echo "Running $(BINARY) with sample data..."
	./$(BUILD_DIR)/$(BINARY) testdata/sample.xlsx

## run-json: Run CLI and output JSON
run-json: build
	./$(BUILD_DIR)/$(BINARY) -f json --pretty testdata/sample.xlsx

## run-csv: Run CLI and output CSV
run-csv: build
	./$(BUILD_DIR)/$(BINARY) -f csv testdata/sample.xlsx

## run-summary: Run CLI in summary mode
run-summary: build
	./$(BUILD_DIR)/$(BINARY) --summary testdata/sample.xlsx

## check: Run all checks (fmt, vet, test)
check: fmt vet test
	@echo "All checks passed!"

## docker: Build Docker image (alias for docker-build)
docker: docker-build

## docker-build: Build Docker image
docker-build:
	@echo "Building Docker image $(DOCKER_IMAGE):$(DOCKER_TAG)..."
	docker build -t $(DOCKER_IMAGE):$(DOCKER_TAG) .
	@echo "Docker image built: $(DOCKER_IMAGE):$(DOCKER_TAG)"

## docker-run: Run goxls in Docker (usage: make docker-run ARGS="file.xlsx")
docker-run:
	@docker run --rm -v "$(PWD):/data" $(DOCKER_IMAGE):$(DOCKER_TAG) $(ARGS)

## docker-test: Test Docker image with sample data
docker-test: docker-build
	@echo "Testing Docker image..."
	@docker run --rm -v "$(PWD)/testdata:/data" $(DOCKER_IMAGE):$(DOCKER_TAG) sample.xlsx --summary

## help: Show this help message
help:
	@echo "Goxls - Excel file reader library"
	@echo ""
	@echo "Usage: make [target]"
	@echo ""
	@echo "Targets:"
	@grep -E '^## ' Makefile | sed 's/## /  /' | column -t -s ':'
