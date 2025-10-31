package main

import (
	"context"
	"fmt"
	"os"
)

func main() {
	ctx := context.Background()

	srv, err := NewSheetsMCPServer(ctx)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Failed to create sheets MCP server: %v\n", err)
		os.Exit(1)
	}

	if err := srv.Run(ctx); err != nil {
		fmt.Fprintf(os.Stderr, "Server error: %v\n", err)
		os.Exit(1)
	}
}
