# Jira Google Sheets Automations

Google Apps Scripts for syncing between Jira and Google Sheets.

## Files

### `jira-export.gs`
Exports selected rows from Google Sheets to Jira CSV format.
- Select rows → Run `processSelectedRows()` → CSV downloads automatically
- Creates separate web (🌐) and mobile (📱) versions
- Maps columns, adds descriptions, epic links, and checklist

**Usage:** Select rows → Menu: `Jira Exporter` → `Process Selected Rows`

### `webhook-reciever.gs`
Webhook endpoint that updates Google Sheets when Jira issue status changes.
- Receives GET requests: `?key=ISSUE-KEY&status=STATUS`
- Maps Jira statuses to "In Progress" or "Done"
- Updates status in "Core" and "Pattern" sheets
- Finds columns by header: `🌐 Task`, `🌐 Status`, `📱 Task`, `📱 Status`

**Status Mapping:**
- `BLOCKED/REJECTED`, `TESTING`, `IN-PROGRESS`, `STORYBOOK`, `PLANNING WEB`, `PLANNING APP` → "In Progress"
- `UAT` → "Done"
- Others → Ignored

## Setup

1. Copy scripts to Google Apps Script editor
2. Deploy `webhook-reciever.gs` as a web app (GET access)
3. Configure Jira webhook to call the deployed URL




