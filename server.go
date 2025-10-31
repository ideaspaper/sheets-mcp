package main

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/mark3labs/mcp-go/mcp"
	"github.com/mark3labs/mcp-go/server"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/sheets/v4"
)

type SheetsMCPServer struct {
	mcpServer     *server.MCPServer
	sheetsService *sheets.Service
	driveService  *drive.Service
	folderID      string
}

func NewSheetsMCPServer(ctx context.Context) (*SheetsMCPServer, error) {
	authConfig := LoadAuthConfig()

	sheetsService, driveService, err := authConfig.CreateServices(ctx)
	if err != nil {
		return nil, fmt.Errorf("failed to create services: %w", err)
	}

	s := &SheetsMCPServer{
		sheetsService: sheetsService,
		driveService:  driveService,
		folderID:      authConfig.DriveFolderID,
	}

	mcpServer := server.NewMCPServer(
		"Google Spreadsheet",
		"1.0.0",
		server.WithToolCapabilities(true),
		server.WithResourceCapabilities(true, false),
	)

	s.mcpServer = mcpServer
	s.registerTools()
	s.registerResources()

	return s, nil
}

func (s *SheetsMCPServer) Run(ctx context.Context) error {
	return server.ServeStdio(s.mcpServer)
}

func (s *SheetsMCPServer) registerTools() {
	// Sheet data operations
	s.mcpServer.AddTool(mcp.NewTool("get_sheet_data",
		mcp.WithDescription("Get data from a specific sheet in a Google Spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithString("range", mcp.Description("Optional cell range in A1 notation")),
		mcp.WithBoolean("include_grid_data", mcp.Description("If True, includes cell formatting and metadata")),
	), s.handleGetSheetData)

	s.mcpServer.AddTool(mcp.NewTool("get_sheet_formulas",
		mcp.WithDescription("Get formulas from a specific sheet in a Google Spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithString("range", mcp.Description("Optional cell range in A1 notation")),
	), s.handleGetSheetFormulas)

	s.mcpServer.AddTool(mcp.NewTool("update_cells",
		mcp.WithDescription("Update cells in a Google Spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithString("range", mcp.Required(), mcp.Description("Cell range in A1 notation")),
		mcp.WithObject("data", mcp.Required(), mcp.Description("2D array of values to update")),
	), s.handleUpdateCells)

	s.mcpServer.AddTool(mcp.NewTool("batch_update_cells",
		mcp.WithDescription("Batch update multiple ranges in a Google Spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithObject("ranges", mcp.Required(), mcp.Description("Dictionary mapping range strings to 2D arrays of values")),
	), s.handleBatchUpdateCells)

	// Row and column operations
	s.mcpServer.AddTool(mcp.NewTool("add_rows",
		mcp.WithDescription("Add rows to a sheet in a Google Spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithNumber("count", mcp.Required(), mcp.Description("Number of rows to add")),
		mcp.WithNumber("start_row", mcp.Description("0-based row index to start adding")),
	), s.handleAddRows)

	s.mcpServer.AddTool(mcp.NewTool("add_columns",
		mcp.WithDescription("Add columns to a sheet in a Google Spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithNumber("count", mcp.Required(), mcp.Description("Number of columns to add")),
		mcp.WithNumber("start_column", mcp.Description("0-based column index to start adding")),
	), s.handleAddColumns)

	// Sheet management
	s.mcpServer.AddTool(mcp.NewTool("list_sheets",
		mcp.WithDescription("List all sheets in a Google Spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
	), s.handleListSheets)

	s.mcpServer.AddTool(mcp.NewTool("create_sheet",
		mcp.WithDescription("Create a new sheet tab in an existing Google Spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("title", mcp.Required(), mcp.Description("The title for the new sheet")),
	), s.handleCreateSheet)

	s.mcpServer.AddTool(mcp.NewTool("copy_sheet",
		mcp.WithDescription("Copy a sheet from one spreadsheet to another"),
		mcp.WithString("src_spreadsheet", mcp.Required(), mcp.Description("Source spreadsheet ID")),
		mcp.WithString("src_sheet", mcp.Required(), mcp.Description("Source sheet name")),
		mcp.WithString("dst_spreadsheet", mcp.Required(), mcp.Description("Destination spreadsheet ID")),
		mcp.WithString("dst_sheet", mcp.Required(), mcp.Description("Destination sheet name")),
	), s.handleCopySheet)

	s.mcpServer.AddTool(mcp.NewTool("rename_sheet",
		mcp.WithDescription("Rename a sheet in a Google Spreadsheet"),
		mcp.WithString("spreadsheet", mcp.Required(), mcp.Description("Spreadsheet ID")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("Current sheet name")),
		mcp.WithString("new_name", mcp.Required(), mcp.Description("New sheet name")),
	), s.handleRenameSheet)

	// Spreadsheet operations
	s.mcpServer.AddTool(mcp.NewTool("list_spreadsheets",
		mcp.WithDescription("List all spreadsheets in the configured Google Drive folder"),
	), s.handleListSpreadsheets)

	s.mcpServer.AddTool(mcp.NewTool("create_spreadsheet",
		mcp.WithDescription("Create a new Google Spreadsheet"),
		mcp.WithString("title", mcp.Required(), mcp.Description("The title of the new spreadsheet")),
	), s.handleCreateSpreadsheet)

	// Multiple queries
	s.mcpServer.AddTool(mcp.NewTool("get_multiple_sheet_data",
		mcp.WithDescription("Get data from multiple specific ranges in Google Spreadsheets"),
		mcp.WithObject("queries", mcp.Required(), mcp.Description("List of query objects with spreadsheet_id, sheet, and range")),
	), s.handleGetMultipleSheetData)

	s.mcpServer.AddTool(mcp.NewTool("get_multiple_spreadsheet_summary",
		mcp.WithDescription("Get a summary of multiple Google Spreadsheets"),
		mcp.WithObject("spreadsheet_ids", mcp.Required(), mcp.Description("List of spreadsheet IDs")),
		mcp.WithNumber("rows_to_fetch", mcp.Description("Number of rows to fetch for summary (default: 5)")),
	), s.handleGetMultipleSpreadsheetSummary)

	// Sharing
	s.mcpServer.AddTool(mcp.NewTool("share_spreadsheet",
		mcp.WithDescription("Share a Google Spreadsheet with multiple users"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithObject("recipients", mcp.Required(), mcp.Description("List of recipient objects with email_address and role")),
		mcp.WithBoolean("send_notification", mcp.Description("Whether to send notification emails (default: true)")),
	), s.handleShareSpreadsheet)

	// Advanced data operations
	s.mcpServer.AddTool(mcp.NewTool("append_data",
		mcp.WithDescription("Append data to the end of a sheet without specifying exact range"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithObject("data", mcp.Required(), mcp.Description("2D array of values to append")),
	), s.handleAppendData)

	s.mcpServer.AddTool(mcp.NewTool("clear_range",
		mcp.WithDescription("Clear content from a specific range in a sheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithString("range", mcp.Required(), mcp.Description("Cell range in A1 notation to clear")),
	), s.handleClearRange)

	s.mcpServer.AddTool(mcp.NewTool("delete_sheet",
		mcp.WithDescription("Delete a sheet tab from a spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet to delete")),
	), s.handleDeleteSheet)

	s.mcpServer.AddTool(mcp.NewTool("duplicate_sheet",
		mcp.WithDescription("Duplicate a sheet within the same spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet to duplicate")),
		mcp.WithString("new_title", mcp.Description("Title for the duplicated sheet")),
	), s.handleDuplicateSheet)

	s.mcpServer.AddTool(mcp.NewTool("find_replace",
		mcp.WithDescription("Find and replace text in a sheet or entire spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("find", mcp.Required(), mcp.Description("The text to find")),
		mcp.WithString("replacement", mcp.Description("The replacement text")),
		mcp.WithString("sheet", mcp.Description("Sheet name (required if all_sheets is false)")),
		mcp.WithBoolean("all_sheets", mcp.Description("Search all sheets (default: false)")),
		mcp.WithBoolean("match_case", mcp.Description("Match case (default: false)")),
		mcp.WithBoolean("match_entire_cell", mcp.Description("Match entire cell (default: false)")),
	), s.handleFindReplace)

	s.mcpServer.AddTool(mcp.NewTool("sort_range",
		mcp.WithDescription("Sort a range of data in a sheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithString("range", mcp.Required(), mcp.Description("Cell range in A1:B2 notation to sort")),
		mcp.WithNumber("sort_column", mcp.Description("0-based column index to sort by (default: 0)")),
		mcp.WithBoolean("ascending", mcp.Description("Sort in ascending order (default: true)")),
	), s.handleSortRange)

	// Formatting operations
	s.mcpServer.AddTool(mcp.NewTool("format_cells",
		mcp.WithDescription("Apply formatting to cells (colors, fonts, text styles)"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithString("range", mcp.Required(), mcp.Description("Cell range in A1:B2 notation")),
		mcp.WithObject("background_color", mcp.Description("Background color {red, green, blue, alpha} (0.0-1.0)")),
		mcp.WithObject("text_color", mcp.Description("Text color {red, green, blue, alpha} (0.0-1.0)")),
		mcp.WithBoolean("bold", mcp.Description("Make text bold")),
		mcp.WithBoolean("italic", mcp.Description("Make text italic")),
		mcp.WithNumber("font_size", mcp.Description("Font size in points")),
	), s.handleFormatCells)

	s.mcpServer.AddTool(mcp.NewTool("merge_cells",
		mcp.WithDescription("Merge cells in a range"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithString("range", mcp.Required(), mcp.Description("Cell range in A1:B2 notation to merge")),
		mcp.WithString("merge_type", mcp.Description("Merge type: MERGE_ALL, MERGE_COLUMNS, MERGE_ROWS (default: MERGE_ALL)")),
	), s.handleMergeCells)

	s.mcpServer.AddTool(mcp.NewTool("unmerge_cells",
		mcp.WithDescription("Unmerge cells in a range"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet")),
		mcp.WithString("range", mcp.Required(), mcp.Description("Cell range in A1:B2 notation to unmerge")),
	), s.handleUnmergeCells)

	s.mcpServer.AddTool(mcp.NewTool("hide_sheet",
		mcp.WithDescription("Hide a sheet in a spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet to hide")),
	), s.handleHideSheet)

	s.mcpServer.AddTool(mcp.NewTool("unhide_sheet",
		mcp.WithDescription("Unhide a sheet in a spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("sheet", mcp.Required(), mcp.Description("The name of the sheet to unhide")),
	), s.handleUnhideSheet)

	// Permission management
	s.mcpServer.AddTool(mcp.NewTool("list_permissions",
		mcp.WithDescription("List all permissions for a spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
	), s.handleListPermissions)

	s.mcpServer.AddTool(mcp.NewTool("remove_permission",
		mcp.WithDescription("Remove a permission from a spreadsheet"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("permission_id", mcp.Required(), mcp.Description("The permission ID to remove")),
	), s.handleRemovePermission)

	// Export operations
	s.mcpServer.AddTool(mcp.NewTool("export_spreadsheet",
		mcp.WithDescription("Export a spreadsheet to different formats (csv, pdf, xlsx, ods, tsv)"),
		mcp.WithString("spreadsheet_id", mcp.Required(), mcp.Description("The ID of the spreadsheet")),
		mcp.WithString("format", mcp.Description("Export format: csv, pdf, xlsx, ods, tsv (default: csv)")),
	), s.handleExportSpreadsheet)
}

func (s *SheetsMCPServer) registerResources() {
	resource := mcp.Resource{
		URI:         "spreadsheet://{spreadsheet_id}/info",
		Name:        "Spreadsheet Info",
		Description: "Get basic information about a Google Spreadsheet",
		MIMEType:    "application/json",
	}
	s.mcpServer.AddResource(resource, s.handleGetSpreadsheetInfo)
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
		return mcp.NewToolResultError(fmt.Sprintf("failed to marshal result: %v", err)), nil
	}
	return mcp.NewToolResultText(string(data)), nil
}

func respondWithError(errMsg string) (*mcp.CallToolResult, error) {
	errorResult := map[string]string{"error": errMsg}
	return respondWithJSON(errorResult)
}
