package main

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/modelcontextprotocol/go-sdk/mcp"
	"google.golang.org/api/sheets/v4"
)

type SheetsMCPServer struct {
	mcpServer     *mcp.Server
	sheetsService *sheets.Service
}

func NewSheetsMCPServer(ctx context.Context) (*SheetsMCPServer, error) {
	authConfig := LoadAuthConfig()

	sheetsService, _, err := authConfig.CreateServices(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to create services: %w", err)
	}

	s := &SheetsMCPServer{
		sheetsService: sheetsService,
	}

	mcpServer := mcp.NewServer(
		&mcp.Implementation{
			Name:    "Google Spreadsheet",
			Version: "1.0.0",
		},
		nil,
	)

	s.mcpServer = mcpServer
	s.registerTools()
	s.registerResources()

	return s, nil
}

func (s *SheetsMCPServer) Run(ctx context.Context) error {
	return s.mcpServer.Run(ctx, &mcp.StdioTransport{})
}

func (s *SheetsMCPServer) registerTools() {
	// Sheet data operations
	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "get_sheet_data",
		Description: "Get data from a specific sheet in a Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id":    map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":             map[string]any{"type": "string", "description": "The name of the sheet"},
				"range":             map[string]any{"type": "string", "description": "Optional cell range in A1 notation"},
				"include_grid_data": map[string]any{"type": "boolean", "description": "If True, includes cell formatting and metadata"},
			},
			"required": []string{"spreadsheet_id", "sheet"},
		}),
	}, s.handleGetSheetData)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "get_sheet_formulas",
		Description: "Get formulas from a specific sheet in a Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"range":          map[string]any{"type": "string", "description": "Optional cell range in A1 notation"},
			},
			"required": []string{"spreadsheet_id", "sheet"},
		}),
	}, s.handleGetSheetFormulas)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "update_cells",
		Description: "Update cells in a Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"range":          map[string]any{"type": "string", "description": "Cell range in A1 notation"},
				"data":           map[string]any{"type": "array", "description": "2D array of values to update"},
			},
			"required": []string{"spreadsheet_id", "sheet", "range", "data"},
		}),
	}, s.handleUpdateCells)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "batch_update_cells",
		Description: "Batch update multiple ranges in a Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"ranges":         map[string]any{"type": "object", "description": "Dictionary mapping range strings to 2D arrays of values"},
			},
			"required": []string{"spreadsheet_id", "sheet", "ranges"},
		}),
	}, s.handleBatchUpdateCells)

	// Row and column operations
	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "add_rows",
		Description: "Add rows to a sheet in a Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"count":          map[string]any{"type": "number", "description": "Number of rows to add"},
				"start_row":      map[string]any{"type": "number", "description": "0-based row index to start adding"},
			},
			"required": []string{"spreadsheet_id", "sheet", "count"},
		}),
	}, s.handleAddRows)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "add_columns",
		Description: "Add columns to a sheet in a Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"count":          map[string]any{"type": "number", "description": "Number of columns to add"},
				"start_column":   map[string]any{"type": "number", "description": "0-based column index to start adding"},
			},
			"required": []string{"spreadsheet_id", "sheet", "count"},
		}),
	}, s.handleAddColumns)

	// Sheet management
	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "list_sheets",
		Description: "List all sheets in a Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
			},
			"required": []string{"spreadsheet_id"},
		}),
	}, s.handleListSheets)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "create_sheet",
		Description: "Create a new sheet tab in an existing Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"title":          map[string]any{"type": "string", "description": "The title for the new sheet"},
			},
			"required": []string{"spreadsheet_id", "title"},
		}),
	}, s.handleCreateSheet)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "copy_sheet",
		Description: "Copy a sheet from one spreadsheet to another",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"src_spreadsheet": map[string]any{"type": "string", "description": "Source spreadsheet ID"},
				"src_sheet":       map[string]any{"type": "string", "description": "Source sheet name"},
				"dst_spreadsheet": map[string]any{"type": "string", "description": "Destination spreadsheet ID"},
				"dst_sheet":       map[string]any{"type": "string", "description": "Destination sheet name"},
			},
			"required": []string{"src_spreadsheet", "src_sheet", "dst_spreadsheet", "dst_sheet"},
		}),
	}, s.handleCopySheet)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "rename_sheet",
		Description: "Rename a sheet in a Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet": map[string]any{"type": "string", "description": "Spreadsheet ID"},
				"sheet":       map[string]any{"type": "string", "description": "Current sheet name"},
				"new_name":    map[string]any{"type": "string", "description": "New sheet name"},
			},
			"required": []string{"spreadsheet", "sheet", "new_name"},
		}),
	}, s.handleRenameSheet)

	// Spreadsheet operations
	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "create_spreadsheet",
		Description: "Create a new Google Spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"title": map[string]any{"type": "string", "description": "The title of the new spreadsheet"},
			},
			"required": []string{"title"},
		}),
	}, s.handleCreateSpreadsheet)

	// Multiple queries
	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "get_multiple_sheet_data",
		Description: "Get data from multiple specific ranges in Google Spreadsheets",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"queries": map[string]any{"type": "array", "description": "List of query objects with spreadsheet_id, sheet, and range"},
			},
			"required": []string{"queries"},
		}),
	}, s.handleGetMultipleSheetData)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "get_multiple_spreadsheet_summary",
		Description: "Get a summary of multiple Google Spreadsheets",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_ids": map[string]any{"type": "array", "description": "List of spreadsheet IDs"},
				"rows_to_fetch":   map[string]any{"type": "number", "description": "Number of rows to fetch for summary (default: 5)"},
			},
			"required": []string{"spreadsheet_ids"},
		}),
	}, s.handleGetMultipleSpreadsheetSummary)

	// Advanced data operations
	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "append_data",
		Description: "Append data to the end of a sheet without specifying exact range",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"data":           map[string]any{"type": "array", "description": "2D array of values to append"},
			},
			"required": []string{"spreadsheet_id", "sheet", "data"},
		}),
	}, s.handleAppendData)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "clear_range",
		Description: "Clear content from a specific range in a sheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"range":          map[string]any{"type": "string", "description": "Cell range in A1 notation to clear"},
			},
			"required": []string{"spreadsheet_id", "sheet", "range"},
		}),
	}, s.handleClearRange)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "delete_sheet",
		Description: "Delete a sheet tab from a spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet to delete"},
			},
			"required": []string{"spreadsheet_id", "sheet"},
		}),
	}, s.handleDeleteSheet)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "duplicate_sheet",
		Description: "Duplicate a sheet within the same spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet to duplicate"},
				"new_title":      map[string]any{"type": "string", "description": "Title for the duplicated sheet"},
			},
			"required": []string{"spreadsheet_id", "sheet"},
		}),
	}, s.handleDuplicateSheet)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "find_replace",
		Description: "Find and replace text in a sheet or entire spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id":    map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"find":              map[string]any{"type": "string", "description": "The text to find"},
				"replacement":       map[string]any{"type": "string", "description": "The replacement text"},
				"sheet":             map[string]any{"type": "string", "description": "Sheet name (required if all_sheets is false)"},
				"all_sheets":        map[string]any{"type": "boolean", "description": "Search all sheets (default: false)"},
				"match_case":        map[string]any{"type": "boolean", "description": "Match case (default: false)"},
				"match_entire_cell": map[string]any{"type": "boolean", "description": "Match entire cell (default: false)"},
			},
			"required": []string{"spreadsheet_id", "find"},
		}),
	}, s.handleFindReplace)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "sort_range",
		Description: "Sort a range of data in a sheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"range":          map[string]any{"type": "string", "description": "Cell range in A1:B2 notation to sort"},
				"sort_column":    map[string]any{"type": "number", "description": "0-based column index to sort by (default: 0)"},
				"ascending":      map[string]any{"type": "boolean", "description": "Sort in ascending order (default: true)"},
			},
			"required": []string{"spreadsheet_id", "sheet", "range"},
		}),
	}, s.handleSortRange)

	// Formatting operations
	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "format_cells",
		Description: "Apply formatting to cells (colors, fonts, text styles)",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id":   map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":            map[string]any{"type": "string", "description": "The name of the sheet"},
				"range":            map[string]any{"type": "string", "description": "Cell range in A1:B2 notation"},
				"background_color": map[string]any{"type": "object", "description": "Background color {red, green, blue, alpha} (0.0-1.0)"},
				"text_color":       map[string]any{"type": "object", "description": "Text color {red, green, blue, alpha} (0.0-1.0)"},
				"bold":             map[string]any{"type": "boolean", "description": "Make text bold"},
				"italic":           map[string]any{"type": "boolean", "description": "Make text italic"},
				"font_size":        map[string]any{"type": "number", "description": "Font size in points"},
			},
			"required": []string{"spreadsheet_id", "sheet", "range"},
		}),
	}, s.handleFormatCells)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "merge_cells",
		Description: "Merge cells in a range",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"range":          map[string]any{"type": "string", "description": "Cell range in A1:B2 notation to merge"},
				"merge_type":     map[string]any{"type": "string", "description": "Merge type: MERGE_ALL, MERGE_COLUMNS, MERGE_ROWS (default: MERGE_ALL)"},
			},
			"required": []string{"spreadsheet_id", "sheet", "range"},
		}),
	}, s.handleMergeCells)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "unmerge_cells",
		Description: "Unmerge cells in a range",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet"},
				"range":          map[string]any{"type": "string", "description": "Cell range in A1:B2 notation to unmerge"},
			},
			"required": []string{"spreadsheet_id", "sheet", "range"},
		}),
	}, s.handleUnmergeCells)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "hide_sheet",
		Description: "Hide a sheet in a spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet to hide"},
			},
			"required": []string{"spreadsheet_id", "sheet"},
		}),
	}, s.handleHideSheet)

	s.mcpServer.AddTool(&mcp.Tool{
		Name:        "unhide_sheet",
		Description: "Unhide a sheet in a spreadsheet",
		InputSchema: mustSchema(map[string]any{
			"type": "object",
			"properties": map[string]any{
				"spreadsheet_id": map[string]any{"type": "string", "description": "The ID of the spreadsheet"},
				"sheet":          map[string]any{"type": "string", "description": "The name of the sheet to unhide"},
			},
			"required": []string{"spreadsheet_id", "sheet"},
		}),
	}, s.handleUnhideSheet)
}

func (s *SheetsMCPServer) registerResources() {
	s.mcpServer.AddResourceTemplate(&mcp.ResourceTemplate{
		URITemplate: "spreadsheet://{spreadsheet_id}/info",
		Name:        "Spreadsheet Info",
		Description: "Get basic information about a Google Spreadsheet",
		MIMEType:    "application/json",
	}, s.handleGetSpreadsheetInfo)
}

func mustSchema(schema map[string]any) map[string]any {
	return schema
}

func parseArgument[T any](args map[string]any, key string, defaultValue T) T {
	if val, ok := args[key]; ok {
		if typed, ok := val.(T); ok {
			return typed
		}
	}
	return defaultValue
}

func respondWithJSON(result any) (*mcp.CallToolResult, error) {
	data, err := json.Marshal(result)
	if err != nil {
		return &mcp.CallToolResult{
			Content: []mcp.Content{&mcp.TextContent{Text: fmt.Sprintf("failed to marshal result: %v", err)}},
			IsError: true,
		}, nil
	}
	return &mcp.CallToolResult{
		Content: []mcp.Content{&mcp.TextContent{Text: string(data)}},
	}, nil
}

func respondWithError(errMsg string) (*mcp.CallToolResult, error) {
	errorResult := map[string]string{"error": errMsg}
	return respondWithJSON(errorResult)
}
