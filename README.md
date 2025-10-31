# Google Sheets MCP Server

A standalone Model Context Protocol (MCP) server that provides comprehensive tools for managing Google Sheets through AI assistants.

## Overview

This MCP server enables AI assistants to interact with Google Sheets, providing tools for:

- Reading and writing sheet data
- Managing sheets and spreadsheets (create, delete, duplicate, hide/unhide)
- Batch operations across multiple sheets
- Advanced data operations (append, clear, find & replace, sort)
- Cell formatting (colors, fonts, merge/unmerge)
- Sharing and permissions management
- Export to multiple formats (CSV, PDF, Excel, etc.)

## Installation

### Prerequisites

- Go 1.24.4 or later
- Google Cloud project with Sheets API and Drive API enabled
- Google service account credentials

```bash
go install github.com/ideaspaper/sheets-mcp@latest
```

## Authentication Setup

### Service Account Setup

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the **Google Sheets API** and **Google Drive API**
   - Navigate to **APIs & Services** > **Library**
   - Search for and enable both APIs
4. Create a service account:
   - Navigate to **IAM & Admin** > **Service Accounts**
   - Click **Create Service Account**
   - Fill in the required fields and click **Create**
5. Create and download a JSON key:
   - Click on the service account email
   - Navigate to the **Keys** tab
   - Click **Add Key** > **Create new key**
   - Choose **JSON** format
   - Save the downloaded file securely
6. **Important**: Share your spreadsheets with the service account email (found in the JSON file as `client_email`)

## Configuration

Set the following environment variables to configure the server:

```bash
export SERVICE_ACCOUNT_PATH="/path/to/service-account-key.json"
# Optional: specify a Drive folder to work with
export DRIVE_FOLDER_ID="your_folder_id_here"
```

## Usage

### OpenCode MCP Client Configuration

Add this to your OpenCode configuration file (Mac: `~/.config/opencode/config.json`):

```json
{
  "mcp": {
    "sheets-mcp": {
      "type": "local",
      "enabled": true,
      "command": ["/path/to/sheets-mcp"],
      "environment": {
        "SERVICE_ACCOUNT_PATH": "/path/to/service-account.json"
      }
    }
  }
}
```

## Available Tools

### Sheet Data Operations

- **get_sheet_data**: Get data from a specific sheet
  - Parameters: `spreadsheet_id`, `sheet`, `range` (optional), `include_grid_data` (optional)

- **get_sheet_formulas**: Get formulas from a specific sheet
  - Parameters: `spreadsheet_id`, `sheet`, `range` (optional)

- **update_cells**: Update cells in a sheet
  - Parameters: `spreadsheet_id`, `sheet`, `range`, `data`

- **batch_update_cells**: Batch update multiple ranges
  - Parameters: `spreadsheet_id`, `sheet`, `ranges`

- **append_data**: Append data to the end of a sheet
  - Parameters: `spreadsheet_id`, `sheet`, `data`

- **clear_range**: Clear content from a specific range
  - Parameters: `spreadsheet_id`, `sheet`, `range`

- **find_replace**: Find and replace text in a sheet or entire spreadsheet
  - Parameters: `spreadsheet_id`, `find`, `replacement` (optional), `sheet` (optional), `all_sheets` (optional), `match_case` (optional), `match_entire_cell` (optional)

- **sort_range**: Sort a range of data
  - Parameters: `spreadsheet_id`, `sheet`, `range`, `sort_column` (optional), `ascending` (optional)

### Row and Column Operations

- **add_rows**: Add rows to a sheet
  - Parameters: `spreadsheet_id`, `sheet`, `count`, `start_row` (optional)

- **add_columns**: Add columns to a sheet
  - Parameters: `spreadsheet_id`, `sheet`, `count`, `start_column` (optional)

### Sheet Management

- **list_sheets**: List all sheets in a spreadsheet
  - Parameters: `spreadsheet_id`

- **create_sheet**: Create a new sheet tab
  - Parameters: `spreadsheet_id`, `title`

- **copy_sheet**: Copy a sheet to another spreadsheet
  - Parameters: `src_spreadsheet`, `src_sheet`, `dst_spreadsheet`, `dst_sheet`

- **rename_sheet**: Rename a sheet
  - Parameters: `spreadsheet`, `sheet`, `new_name`

- **delete_sheet**: Delete a sheet tab
  - Parameters: `spreadsheet_id`, `sheet`

- **duplicate_sheet**: Duplicate a sheet within the same spreadsheet
  - Parameters: `spreadsheet_id`, `sheet`, `new_title` (optional)

- **hide_sheet**: Hide a sheet
  - Parameters: `spreadsheet_id`, `sheet`

- **unhide_sheet**: Unhide a sheet
  - Parameters: `spreadsheet_id`, `sheet`

### Spreadsheet Operations

- **list_spreadsheets**: List all spreadsheets in configured Drive folder
  - No parameters required

- **create_spreadsheet**: Create a new spreadsheet
  - Parameters: `title`

- **share_spreadsheet**: Share a spreadsheet with users
  - Parameters: `spreadsheet_id`, `recipients`, `send_notification` (optional)

- **list_permissions**: List all permissions for a spreadsheet
  - Parameters: `spreadsheet_id`

- **remove_permission**: Remove a permission from a spreadsheet
  - Parameters: `spreadsheet_id`, `permission_id`

- **export_spreadsheet**: Export a spreadsheet to different formats
  - Parameters: `spreadsheet_id`, `format` (csv, pdf, xlsx, ods, tsv - default: csv)

### Formatting Operations

- **format_cells**: Apply formatting to cells (colors, fonts, text styles)
  - Parameters: `spreadsheet_id`, `sheet`, `range`, `background_color` (optional), `text_color` (optional), `bold` (optional), `italic` (optional), `font_size` (optional)

- **merge_cells**: Merge cells in a range
  - Parameters: `spreadsheet_id`, `sheet`, `range`, `merge_type` (optional: MERGE_ALL, MERGE_COLUMNS, MERGE_ROWS)

- **unmerge_cells**: Unmerge cells in a range
  - Parameters: `spreadsheet_id`, `sheet`, `range`

### Batch Operations

- **get_multiple_sheet_data**: Get data from multiple ranges
  - Parameters: `queries` (array of query objects)

- **get_multiple_spreadsheet_summary**: Get summary of multiple spreadsheets
  - Parameters: `spreadsheet_ids`, `rows_to_fetch` (optional, default: 5)

## Troubleshooting

### Authentication Errors

- Ensure your service account email is shared with the spreadsheet you're trying to access
- Verify that both Sheets API and Drive API are enabled in your Google Cloud project
- Check that the credentials file path is correct and the file is readable

### Permission Errors

- Make sure the service account has been granted access to the spreadsheet
- For Drive folder operations, ensure the service account has access to the specified folder

### API Quota Errors

- Google Sheets API has rate limits. If you hit them, wait a few minutes before retrying
- Consider implementing exponential backoff in your client application
