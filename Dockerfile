# Build stage
FROM golang:1.23-alpine AS builder

WORKDIR /app

# Install git for go mod download (some dependencies may need it)
RUN apk add --no-cache git

# Copy go mod files first for better caching
COPY go.mod go.sum ./
RUN go mod download

# Copy source code
COPY . .

# Build the binary
RUN CGO_ENABLED=0 GOOS=linux go build -ldflags="-w -s" -o goxls ./cmd

# Runtime stage
FROM alpine:3.19

# Add ca-certificates for HTTPS (if needed in future)
RUN apk add --no-cache ca-certificates

# Create non-root user for security
RUN adduser -D -g '' goxls
USER goxls

WORKDIR /data

# Copy binary from builder
COPY --from=builder /app/goxls /usr/local/bin/goxls

ENTRYPOINT ["goxls"]
CMD ["--help"]
