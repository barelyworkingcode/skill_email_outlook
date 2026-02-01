# Outlook Email Exporter

Export Outlook.com emails to Markdown files with attachments.

## Setup

```bash
# Create conda environment
conda env create -f environment.yml
conda activate outlook-export
```

## First-Time Authorization

```bash
python outlook_export.py --reauth
```

1. A URL and code will display
2. Open https://microsoft.com/devicelogin in your browser
3. Enter the code and sign in with your Microsoft account
4. Token saves automatically to `.token_cache.json` (valid ~90 days)

## Usage

```bash
# List folders
python outlook_export.py --list-folders

# Export folder (outputs to ./mail/Inbox/)
python outlook_export.py -f Inbox

# Limit emails
python outlook_export.py -f Inbox --limit 100

# Batch size (default 50)
python outlook_export.py -f Inbox --batch-size 25

# Group by conversation thread
python outlook_export.py -f Inbox --threads

# Filter emails
python outlook_export.py -f Inbox --filter "flag/flagStatus eq 'flagged'"
python outlook_export.py -f Inbox --filter "isRead eq false"
python outlook_export.py -f Inbox --filter "receivedDateTime ge 2024-01-01"

# Show all filter options
python outlook_export.py -?
```

## Output Structure

**Without threads:**
```
./mail/Inbox/
  20240130_Meeting_notes.md
  20240129_Invoice/
    message.md
    attachments/
      invoice.pdf
```

**With threads (`--threads`):**
```
./mail/Inbox/
  20240115_Project_Discussion/
    20240115_093022.md
    20240116_141532.md
    attachments/
      20240116_report.pdf
```

## Resumable Downloads

- `.manifest.json` tracks downloaded message IDs
- Re-running skips already downloaded emails
- Safe to interrupt and resume

## Agent/Automation Usage

For non-interactive use (CI, AI agents, scripts):

```bash
# Step 1: Get tokens interactively (once)
python outlook_export.py --auth-only

# Step 2: Use refresh token (~90 days, recommended)
python outlook_export.py --refresh-token "0.AYIA..." -f Inbox

# Or use token cache file
python outlook_export.py --token-cache /path/to/.token_cache.json -f Inbox

# Or use access token (~1 hour only)
python outlook_export.py --token "eyJ0eX..." -f Inbox
```

## Options Reference

| Option | Description |
|--------|-------------|
| `-f, --folder` | Folder name to export |
| `-o, --output` | Output directory (default: `./mail/{folder}`) |
| `-n, --limit` | Max emails to export |
| `-b, --batch-size` | Emails per API request (default: 50) |
| `-s, --skip` | Skip first N emails |
| `-t, --threads` | Group by conversation thread |
| `--filter` | OData filter query |
| `-?`, `--filter-help` | Show filter examples |
| `--list-folders` | List available folders |
| `--reauth` | Clear cache and re-authenticate |
| `--auth-only` | Print tokens and exit |
| `--token` | Use access token (~1 hour) |
| `--refresh-token` | Use refresh token (~90 days) |
| `--token-cache` | Use token cache file |

## Claude Desktop Integration (MCP Server)

Expose email tools to Claude Desktop for prompts like "check my recent emails and list action items."

### Why MCP instead of a cloud connector?

| | MCP Server | Cloud Connector (Zapier, etc.) |
|---|---|---|
| **Data location** | Stays on your machine | Flows through third-party servers |
| **Privacy** | Emails never leave your computer | Provider can access your data |
| **Auth** | You control tokens locally | OAuth grants access to connector service |
| **Cost** | Free | Often subscription-based |
| **Latency** | Direct API calls | Extra hop through cloud |
| **Offline** | Works with cached tokens | Requires internet to connector |

MCP (Model Context Protocol) is Anthropic's open standard for AI-tool integration. The server runs locally alongside Claude Desktop - when Claude needs email data, it calls your local MCP server, which fetches directly from Microsoft's API. Your emails are read by Claude but never stored or transmitted elsewhere.

### Setup

1. **Authenticate first:**
   ```bash
   python outlook_export.py --auth-only
   ```

2. **Configure Claude Desktop** (`~/Library/Application Support/Claude/claude_desktop_config.json`):
   ```json
   {
     "mcpServers": {
       "outlook-email": {
         "command": "/Users/{username}/miniconda3/envs/outlook-export/bin/python",
         "args": ["/Users/{username}/source/email/mcp_server.py"]
       }
     }
   }
   ```

3. **Restart Claude Desktop**

### Available Tools

| Tool | Description |
|------|-------------|
| `list_email_folders` | List all folders in mailbox |
| `get_recent_emails` | Get emails from last N days (filter by unread/flagged) |
| `search_emails` | Search with OData filter query |
| `get_email_by_id` | Get full email content by ID |

### Example Prompts

- "Check my inbox for emails from the last 3 days and summarize any action items"
- "Show me my unread emails"
- "Find all flagged emails and list what needs my attention"
- "Search for emails from boss@company.com"
