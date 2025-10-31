package main

import (
	"context"
	"encoding/json"
	"fmt"
	"strings"

	"github.com/mark3labs/mcp-go/mcp"
	"google.golang.org/api/drive/v3"
	"google.golang.org/api/sheets/v4"
)

func getArgsFromRequest(request mcp.CallToolRequest) (map[string]any, error) {
	args, ok := request.Params.Arguments.(map[string]any)
	if !ok {
		return nil, fmt.Errorf("invalid arguments")
	}
	return args, nil
}

func (s *SheetsMCPServer) handleGetSheetData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")
	sheet := parseArgument(args, "sheet", "")
	rangeStr := parseArgument(args, "range", "")
	includeGridData := parseArgument(args, "include_grid_data", false)

	if spreadsheetID == "" || sheet == "" {
		return respondWithError("spreadsheet_id and sheet are required")
	}

	fullRange := sheet
	if rangeStr != "" {
		fullRange = fmt.Sprintf("%s!%s", sheet, rangeStr)
	}

	if includeGridData {
		result, err := s.sheetsService.Spreadsheets.Get(spreadsheetID).
			Ranges(fullRange).
			IncludeGridData(true).
			Do()
		if err != nil {
			return respondWithError(fmt.Sprintf("failed to get sheet data: %v", err))
		}
		return respondWithJSON(result)
	}

	valuesResult, err := s.sheetsService.Spreadsheets.Values.Get(spreadsheetID, fullRange).Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to get sheet values: %v", err))
	}

	response := map[string]any{
		"spreadsheetId": spreadsheetID,
		"valueRanges": []map[string]any{
			{
				"range":  fullRange,
				"values": valuesResult.Values,
			},
		},
	}

	return respondWithJSON(response)
}

func (s *SheetsMCPServer) handleGetSheetFormulas(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")
	sheet := parseArgument(args, "sheet", "")
	rangeStr := parseArgument(args, "range", "")

	if spreadsheetID == "" || sheet == "" {
		return respondWithError("spreadsheet_id and sheet are required")
	}

	fullRange := sheet
	if rangeStr != "" {
		fullRange = fmt.Sprintf("%s!%s", sheet, rangeStr)
	}

	result, err := s.sheetsService.Spreadsheets.Values.Get(spreadsheetID, fullRange).
		ValueRenderOption("FORMULA").
		Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to get formulas: %v", err))
	}

	return respondWithJSON(result.Values)
}

func (s *SheetsMCPServer) handleUpdateCells(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")
	sheet := parseArgument(args, "sheet", "")
	rangeStr := parseArgument(args, "range", "")

	if spreadsheetID == "" || sheet == "" || rangeStr == "" {
		return respondWithError("spreadsheet_id, sheet, and range are required")
	}

	dataRaw, ok := args["data"]
	if !ok {
		return respondWithError("data is required")
	}

	data, err := convertToValues(dataRaw)
	if err != nil {
		return respondWithError(fmt.Sprintf("invalid data format: %v", err))
	}

	fullRange := fmt.Sprintf("%s!%s", sheet, rangeStr)

	valueRange := &sheets.ValueRange{
		Values: data,
	}

	result, err := s.sheetsService.Spreadsheets.Values.Update(spreadsheetID, fullRange, valueRange).
		ValueInputOption("USER_ENTERED").
		Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to update cells: %v", err))
	}

	return respondWithJSON(result)
}

func (s *SheetsMCPServer) handleBatchUpdateCells(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")
	sheet := parseArgument(args, "sheet", "")

	if spreadsheetID == "" || sheet == "" {
		return respondWithError("spreadsheet_id and sheet are required")
	}

	rangesRaw, ok := args["ranges"]
	if !ok {
		return respondWithError("ranges is required")
	}

	rangesMap, ok := rangesRaw.(map[string]any)
	if !ok {
		return respondWithError("ranges must be an object/map")
	}

	var valueRanges []*sheets.ValueRange
	for rangeStr, valuesRaw := range rangesMap {
		values, err := convertToValues(valuesRaw)
		if err != nil {
			return respondWithError(fmt.Sprintf("invalid data format for range %s: %v", rangeStr, err))
		}

		fullRange := fmt.Sprintf("%s!%s", sheet, rangeStr)
		valueRanges = append(valueRanges, &sheets.ValueRange{
			Range:  fullRange,
			Values: values,
		})
	}

	batchUpdate := &sheets.BatchUpdateValuesRequest{
		ValueInputOption: "USER_ENTERED",
		Data:             valueRanges,
	}

	result, err := s.sheetsService.Spreadsheets.Values.BatchUpdate(spreadsheetID, batchUpdate).Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to batch update cells: %v", err))
	}

	return respondWithJSON(result)
}

func (s *SheetsMCPServer) handleAddRows(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")
	sheet := parseArgument(args, "sheet", "")
	count := int64(parseArgument(args, "count", float64(0)))

	if spreadsheetID == "" || sheet == "" || count <= 0 {
		return respondWithError("spreadsheet_id, sheet, and count are required")
	}

	sheetID, err := s.getSheetID(spreadsheetID, sheet)
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to get sheet ID: %v", err))
	}

	startRow := int64(0)
	if _, ok := args["start_row"]; ok {
		startRow = int64(parseArgument(args, "start_row", float64(0)))
	}

	requests := []*sheets.Request{
		{
			InsertDimension: &sheets.InsertDimensionRequest{
				Range: &sheets.DimensionRange{
					SheetId:    sheetID,
					Dimension:  "ROWS",
					StartIndex: startRow,
					EndIndex:   startRow + count,
				},
				InheritFromBefore: startRow > 0,
			},
		},
	}

	batchUpdate := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}

	result, err := s.sheetsService.Spreadsheets.BatchUpdate(spreadsheetID, batchUpdate).Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to add rows: %v", err))
	}

	return respondWithJSON(result)
}

func (s *SheetsMCPServer) handleAddColumns(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")
	sheet := parseArgument(args, "sheet", "")
	count := int64(parseArgument(args, "count", float64(0)))

	if spreadsheetID == "" || sheet == "" || count <= 0 {
		return respondWithError("spreadsheet_id, sheet, and count are required")
	}

	sheetID, err := s.getSheetID(spreadsheetID, sheet)
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to get sheet ID: %v", err))
	}

	startColumn := int64(0)
	if _, ok := args["start_column"]; ok {
		startColumn = int64(parseArgument(args, "start_column", float64(0)))
	}

	requests := []*sheets.Request{
		{
			InsertDimension: &sheets.InsertDimensionRequest{
				Range: &sheets.DimensionRange{
					SheetId:    sheetID,
					Dimension:  "COLUMNS",
					StartIndex: startColumn,
					EndIndex:   startColumn + count,
				},
				InheritFromBefore: startColumn > 0,
			},
		},
	}

	batchUpdate := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}

	result, err := s.sheetsService.Spreadsheets.BatchUpdate(spreadsheetID, batchUpdate).Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to add columns: %v", err))
	}

	return respondWithJSON(result)
}

func (s *SheetsMCPServer) handleListSheets(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")

	if spreadsheetID == "" {
		return respondWithError("spreadsheet_id is required")
	}

	spreadsheet, err := s.sheetsService.Spreadsheets.Get(spreadsheetID).Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to get spreadsheet: %v", err))
	}

	var sheetNames []string
	for _, sheet := range spreadsheet.Sheets {
		sheetNames = append(sheetNames, sheet.Properties.Title)
	}

	return respondWithJSON(sheetNames)
}

func (s *SheetsMCPServer) handleCreateSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")
	title := parseArgument(args, "title", "")

	if spreadsheetID == "" || title == "" {
		return respondWithError("spreadsheet_id and title are required")
	}

	requests := []*sheets.Request{
		{
			AddSheet: &sheets.AddSheetRequest{
				Properties: &sheets.SheetProperties{
					Title: title,
				},
			},
		},
	}

	batchUpdate := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}

	result, err := s.sheetsService.Spreadsheets.BatchUpdate(spreadsheetID, batchUpdate).Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to create sheet: %v", err))
	}

	if len(result.Replies) > 0 && result.Replies[0].AddSheet != nil {
		props := result.Replies[0].AddSheet.Properties
		response := map[string]any{
			"sheetId":       props.SheetId,
			"title":         props.Title,
			"index":         props.Index,
			"spreadsheetId": spreadsheetID,
		}
		return respondWithJSON(response)
	}

	return respondWithJSON(result)
}

func (s *SheetsMCPServer) handleCopySheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	srcSpreadsheet := parseArgument(args, "src_spreadsheet", "")
	srcSheet := parseArgument(args, "src_sheet", "")
	dstSpreadsheet := parseArgument(args, "dst_spreadsheet", "")
	dstSheet := parseArgument(args, "dst_sheet", "")

	if srcSpreadsheet == "" || srcSheet == "" || dstSpreadsheet == "" || dstSheet == "" {
		return respondWithError("src_spreadsheet, src_sheet, dst_spreadsheet, and dst_sheet are required")
	}

	srcSheetID, err := s.getSheetID(srcSpreadsheet, srcSheet)
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to get source sheet ID: %v", err))
	}

	copyRequest := &sheets.CopySheetToAnotherSpreadsheetRequest{
		DestinationSpreadsheetId: dstSpreadsheet,
	}

	copyResult, err := s.sheetsService.Spreadsheets.Sheets.CopyTo(srcSpreadsheet, srcSheetID, copyRequest).Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to copy sheet: %v", err))
	}

	result := map[string]any{
		"copy": copyResult,
	}

	if copyResult.Title != dstSheet {
		requests := []*sheets.Request{
			{
				UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
					Properties: &sheets.SheetProperties{
						SheetId: copyResult.SheetId,
						Title:   dstSheet,
					},
					Fields: "title",
				},
			},
		}

		batchUpdate := &sheets.BatchUpdateSpreadsheetRequest{
			Requests: requests,
		}

		renameResult, err := s.sheetsService.Spreadsheets.BatchUpdate(dstSpreadsheet, batchUpdate).Do()
		if err != nil {
			return respondWithError(fmt.Sprintf("failed to rename copied sheet: %v", err))
		}
		result["rename"] = renameResult
	}

	return respondWithJSON(result)
}

func (s *SheetsMCPServer) handleRenameSheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheet := parseArgument(args, "spreadsheet", "")
	sheet := parseArgument(args, "sheet", "")
	newName := parseArgument(args, "new_name", "")

	if spreadsheet == "" || sheet == "" || newName == "" {
		return respondWithError("spreadsheet, sheet, and new_name are required")
	}

	sheetID, err := s.getSheetID(spreadsheet, sheet)
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to get sheet ID: %v", err))
	}

	requests := []*sheets.Request{
		{
			UpdateSheetProperties: &sheets.UpdateSheetPropertiesRequest{
				Properties: &sheets.SheetProperties{
					SheetId: sheetID,
					Title:   newName,
				},
				Fields: "title",
			},
		},
	}

	batchUpdate := &sheets.BatchUpdateSpreadsheetRequest{
		Requests: requests,
	}

	result, err := s.sheetsService.Spreadsheets.BatchUpdate(spreadsheet, batchUpdate).Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to rename sheet: %v", err))
	}

	return respondWithJSON(result)
}

func (s *SheetsMCPServer) getSheetID(spreadsheetID, sheetName string) (int64, error) {
	spreadsheet, err := s.sheetsService.Spreadsheets.Get(spreadsheetID).Do()
	if err != nil {
		return 0, err
	}

	for _, sheet := range spreadsheet.Sheets {
		if sheet.Properties.Title == sheetName {
			return sheet.Properties.SheetId, nil
		}
	}

	return 0, fmt.Errorf("sheet '%s' not found", sheetName)
}

func convertToValues(data any) ([][]any, error) {
	dataBytes, err := json.Marshal(data)
	if err != nil {
		return nil, err
	}

	var values [][]any
	if err := json.Unmarshal(dataBytes, &values); err != nil {
		return nil, err
	}

	return values, nil
}

func (s *SheetsMCPServer) handleListSpreadsheets(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	_, _ = getArgsFromRequest(request)
	query := "mimeType='application/vnd.google-apps.spreadsheet'"

	if s.folderID != "" {
		query = fmt.Sprintf("%s and '%s' in parents", query, s.folderID)
		fmt.Printf("Searching for spreadsheets in folder: %s\n", s.folderID)
	} else {
		fmt.Println("Searching for spreadsheets in 'My Drive'")
	}

	results, err := s.driveService.Files.List().
		Q(query).
		Spaces("drive").
		IncludeItemsFromAllDrives(true).
		SupportsAllDrives(true).
		Fields("files(id, name)").
		OrderBy("modifiedTime desc").
		Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to list spreadsheets: %v", err))
	}

	var spreadsheets []map[string]string
	for _, file := range results.Files {
		spreadsheets = append(spreadsheets, map[string]string{
			"id":    file.Id,
			"title": file.Name,
		})
	}

	return respondWithJSON(spreadsheets)
}

func (s *SheetsMCPServer) handleCreateSpreadsheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	title := parseArgument(args, "title", "")

	if title == "" {
		return respondWithError("title is required")
	}

	driveFile := &drive.File{
		Name:     title,
		MimeType: "application/vnd.google-apps.spreadsheet",
	}

	if s.folderID != "" {
		driveFile.Parents = []string{s.folderID}
	}

	file, err := s.driveService.Files.Create(driveFile).
		SupportsAllDrives(true).
		Fields("id, name, parents").
		Do()
	if err != nil {
		return respondWithError(fmt.Sprintf("failed to create spreadsheet: %v", err))
	}

	spreadsheetID := file.Id
	fmt.Printf("Spreadsheet created with ID: %s\n", spreadsheetID)

	folder := "root"
	if len(file.Parents) > 0 {
		folder = file.Parents[0]
	}

	response := map[string]any{
		"spreadsheetId": spreadsheetID,
		"title":         file.Name,
		"folder":        folder,
	}

	return respondWithJSON(response)
}

func (s *SheetsMCPServer) handleGetMultipleSheetData(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	queriesRaw, ok := args["queries"]
	if !ok {
		return respondWithError("queries is required")
	}

	queriesBytes, err := json.Marshal(queriesRaw)
	if err != nil {
		return respondWithError("invalid queries format")
	}

	var queries []map[string]string
	if err := json.Unmarshal(queriesBytes, &queries); err != nil {
		return respondWithError("invalid queries format")
	}

	var results []map[string]any

	for _, query := range queries {
		spreadsheetID := query["spreadsheet_id"]
		sheet := query["sheet"]
		rangeStr := query["range"]

		if spreadsheetID == "" || sheet == "" || rangeStr == "" {
			results = append(results, map[string]any{
				"spreadsheet_id": spreadsheetID,
				"sheet":          sheet,
				"range":          rangeStr,
				"error":          "Missing required keys (spreadsheet_id, sheet, range)",
			})
			continue
		}

		fullRange := fmt.Sprintf("%s!%s", sheet, rangeStr)

		valuesResult, err := s.sheetsService.Spreadsheets.Values.Get(spreadsheetID, fullRange).Do()
		if err != nil {
			results = append(results, map[string]any{
				"spreadsheet_id": spreadsheetID,
				"sheet":          sheet,
				"range":          rangeStr,
				"error":          err.Error(),
			})
			continue
		}

		results = append(results, map[string]any{
			"spreadsheet_id": spreadsheetID,
			"sheet":          sheet,
			"range":          rangeStr,
			"data":           valuesResult.Values,
		})
	}

	return respondWithJSON(results)
}

func (s *SheetsMCPServer) handleGetMultipleSpreadsheetSummary(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	idsRaw, ok := args["spreadsheet_ids"]
	if !ok {
		return respondWithError("spreadsheet_ids is required")
	}

	idsBytes, err := json.Marshal(idsRaw)
	if err != nil {
		return respondWithError("invalid spreadsheet_ids format")
	}

	var spreadsheetIDs []string
	if err := json.Unmarshal(idsBytes, &spreadsheetIDs); err != nil {
		return respondWithError("invalid spreadsheet_ids format")
	}

	rowsToFetch := max(1, int(parseArgument(args, "rows_to_fetch", float64(5))))

	var summaries []map[string]any

	for _, spreadsheetID := range spreadsheetIDs {
		summary := map[string]any{
			"spreadsheet_id": spreadsheetID,
			"title":          nil,
			"sheets":         []map[string]any{},
			"error":          nil,
		}

		spreadsheet, err := s.sheetsService.Spreadsheets.Get(spreadsheetID).
			Fields("properties.title,sheets(properties(title,sheetId))").
			Do()
		if err != nil {
			summary["error"] = fmt.Sprintf("Error fetching spreadsheet %s: %v", spreadsheetID, err)
			summaries = append(summaries, summary)
			continue
		}

		summary["title"] = spreadsheet.Properties.Title

		var sheetSummaries []map[string]any

		for _, sheet := range spreadsheet.Sheets {
			sheetTitle := sheet.Properties.Title
			sheetID := sheet.Properties.SheetId

			sheetSummary := map[string]any{
				"title":      sheetTitle,
				"sheet_id":   sheetID,
				"headers":    []any{},
				"first_rows": []any{},
				"error":      nil,
			}

			if sheetTitle == "" {
				sheetSummary["error"] = "Sheet title not found"
				sheetSummaries = append(sheetSummaries, sheetSummary)
				continue
			}

			rangeToGet := fmt.Sprintf("%s!A1:%d", sheetTitle, rowsToFetch)

			valuesResult, err := s.sheetsService.Spreadsheets.Values.Get(spreadsheetID, rangeToGet).Do()
			if err != nil {
				sheetSummary["error"] = fmt.Sprintf("Error fetching data for sheet %s: %v", sheetTitle, err)
				sheetSummaries = append(sheetSummaries, sheetSummary)
				continue
			}

			if len(valuesResult.Values) > 0 {
				sheetSummary["headers"] = valuesResult.Values[0]
				if len(valuesResult.Values) > 1 {
					sheetSummary["first_rows"] = valuesResult.Values[1:]
				}
			}

			sheetSummaries = append(sheetSummaries, sheetSummary)
		}

		summary["sheets"] = sheetSummaries
		summaries = append(summaries, summary)
	}

	return respondWithJSON(summaries)
}

func (s *SheetsMCPServer) handleShareSpreadsheet(ctx context.Context, request mcp.CallToolRequest) (*mcp.CallToolResult, error) {
	args, err := getArgsFromRequest(request)
	if err != nil {
		return respondWithError(err.Error())
	}
	spreadsheetID := parseArgument(args, "spreadsheet_id", "")
	sendNotification := parseArgument(args, "send_notification", true)

	if spreadsheetID == "" {
		return respondWithError("spreadsheet_id is required")
	}

	recipientsRaw, ok := args["recipients"]
	if !ok {
		return respondWithError("recipients is required")
	}

	recipientsBytes, err := json.Marshal(recipientsRaw)
	if err != nil {
		return respondWithError("invalid recipients format")
	}

	var recipients []map[string]string
	if err := json.Unmarshal(recipientsBytes, &recipients); err != nil {
		return respondWithError("invalid recipients format")
	}

	var successes []map[string]any
	var failures []map[string]any

	for _, recipient := range recipients {
		emailAddress := recipient["email_address"]
		role := recipient["role"]
		if role == "" {
			role = "writer"
		}

		if emailAddress == "" {
			failures = append(failures, map[string]any{
				"email_address": nil,
				"error":         "Missing email_address in recipient entry.",
			})
			continue
		}

		if role != "reader" && role != "commenter" && role != "writer" {
			failures = append(failures, map[string]any{
				"email_address": emailAddress,
				"error":         fmt.Sprintf("Invalid role '%s'. Must be 'reader', 'commenter', or 'writer'.", role),
			})
			continue
		}

		permission := &drive.Permission{
			Type:         "user",
			Role:         role,
			EmailAddress: emailAddress,
		}

		result, err := s.driveService.Permissions.Create(spreadsheetID, permission).
			SendNotificationEmail(sendNotification).
			Fields("id").
			Do()
		if err != nil {
			failures = append(failures, map[string]any{
				"email_address": emailAddress,
				"error":         fmt.Sprintf("Failed to share: %v", err),
			})
			continue
		}

		successes = append(successes, map[string]any{
			"email_address": emailAddress,
			"role":          role,
			"permissionId":  result.Id,
		})
	}

	response := map[string]any{
		"successes": successes,
		"failures":  failures,
	}

	return respondWithJSON(response)
}

func (s *SheetsMCPServer) handleGetSpreadsheetInfo(ctx context.Context, request mcp.ReadResourceRequest) ([]mcp.ResourceContents, error) {
	uri := request.Params.URI

	parts := strings.Split(uri, "://")
	if len(parts) != 2 {
		return nil, fmt.Errorf("invalid URI format")
	}

	pathParts := strings.Split(parts[1], "/")
	if len(pathParts) < 1 {
		return nil, fmt.Errorf("invalid URI format: missing spreadsheet_id")
	}

	spreadsheetID := pathParts[0]

	spreadsheet, err := s.sheetsService.Spreadsheets.Get(spreadsheetID).Do()
	if err != nil {
		return nil, fmt.Errorf("failed to get spreadsheet: %w", err)
	}

	type SheetInfo struct {
		Title          string                 `json:"title"`
		SheetID        int64                  `json:"sheetId"`
		GridProperties *sheets.GridProperties `json:"gridProperties"`
	}

	type SpreadsheetInfo struct {
		Title  string      `json:"title"`
		Sheets []SheetInfo `json:"sheets"`
	}

	info := SpreadsheetInfo{
		Title: spreadsheet.Properties.Title,
	}

	for _, sheet := range spreadsheet.Sheets {
		info.Sheets = append(info.Sheets, SheetInfo{
			Title:          sheet.Properties.Title,
			SheetID:        sheet.Properties.SheetId,
			GridProperties: sheet.Properties.GridProperties,
		})
	}

	infoBytes, err := json.MarshalIndent(info, "", "  ")
	if err != nil {
		return nil, fmt.Errorf("failed to marshal info: %w", err)
	}

	content := &mcp.TextResourceContents{
		URI:      uri,
		MIMEType: "application/json",
		Text:     string(infoBytes),
	}

	return []mcp.ResourceContents{content}, nil
}
