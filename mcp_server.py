#!/usr/bin/env python3
"""MCP Server for Outlook Email access."""

import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent

# Import from our existing module
from outlook_export import (
    get_token_from_refresh_token,
    get_token_from_cache_file,
    list_folders,
    get_emails_batch,
    get_attachments,
    api_get,
    GRAPH_API,
    TOKEN_CACHE,
)

server = Server("outlook-email")

# Configuration - set via environment or config file
CONFIG_FILE = Path(__file__).parent / ".mcp_config.json"


def load_config() -> dict:
    if CONFIG_FILE.exists():
        return json.loads(CONFIG_FILE.read_text())
    return {}


def get_token() -> str:
    """Get access token from config."""
    config = load_config()

    if "refresh_token" in config:
        return get_token_from_refresh_token(config["refresh_token"])
    elif "token_cache_path" in config:
        return get_token_from_cache_file(Path(config["token_cache_path"]))
    elif TOKEN_CACHE.exists():
        return get_token_from_cache_file(TOKEN_CACHE)
    else:
        raise Exception("No authentication configured. Run: python outlook_export.py --auth-only")


@server.list_tools()
async def list_tools():
    return [
        Tool(
            name="list_email_folders",
            description="List all email folders in the Outlook mailbox",
            inputSchema={
                "type": "object",
                "properties": {},
                "required": []
            }
        ),
        Tool(
            name="get_recent_emails",
            description="Get recent emails from a folder. Returns email metadata and body content.",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder": {
                        "type": "string",
                        "description": "Folder name (e.g., 'Inbox', 'Sent Items')",
                        "default": "Inbox"
                    },
                    "days": {
                        "type": "integer",
                        "description": "Get emails from the last N days",
                        "default": 3
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of emails to return",
                        "default": 20
                    },
                    "unread_only": {
                        "type": "boolean",
                        "description": "Only return unread emails",
                        "default": False
                    },
                    "flagged_only": {
                        "type": "boolean",
                        "description": "Only return flagged emails",
                        "default": False
                    }
                },
                "required": []
            }
        ),
        Tool(
            name="search_emails",
            description="Search emails with a custom filter query",
            inputSchema={
                "type": "object",
                "properties": {
                    "folder": {
                        "type": "string",
                        "description": "Folder name",
                        "default": "Inbox"
                    },
                    "filter": {
                        "type": "string",
                        "description": "OData filter query (e.g., \"from/emailAddress/address eq 'someone@example.com'\")"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of emails",
                        "default": 20
                    }
                },
                "required": ["filter"]
            }
        ),
        Tool(
            name="get_email_by_id",
            description="Get full content of a specific email by ID",
            inputSchema={
                "type": "object",
                "properties": {
                    "email_id": {
                        "type": "string",
                        "description": "The email ID"
                    }
                },
                "required": ["email_id"]
            }
        )
    ]


def format_email_summary(email: dict) -> str:
    """Format email for display."""
    from_addr = email.get("from", {}).get("emailAddress", {})
    from_str = from_addr.get("name") or from_addr.get("address", "Unknown")

    date_str = email.get("receivedDateTime", "")[:16].replace("T", " ")
    subject = email.get("subject", "(No Subject)")
    is_read = "read" if email.get("isRead") else "UNREAD"
    flagged = " [FLAGGED]" if email.get("flag", {}).get("flagStatus") == "flagged" else ""
    has_attach = " [ATTACHMENTS]" if email.get("hasAttachments") else ""

    body = email.get("body", {}).get("content", "")
    # Convert HTML to plain text (simple version)
    if email.get("body", {}).get("contentType") == "html":
        import re
        body = re.sub(r'<[^>]+>', '', body)
        body = re.sub(r'\s+', ' ', body).strip()

    # Truncate body
    if len(body) > 500:
        body = body[:500] + "..."

    return f"""
---
ID: {email.get('id', '')}
From: {from_str}
Date: {date_str}
Subject: {subject}
Status: {is_read}{flagged}{has_attach}

{body}
"""


def get_folder_id(token: str, folder_name: str) -> str:
    """Get folder ID by name."""
    folders = list_folders(token)
    for folder in folders:
        if folder["displayName"].lower() == folder_name.lower():
            return folder["id"]
    raise Exception(f"Folder '{folder_name}' not found")


@server.call_tool()
async def call_tool(name: str, arguments: dict):
    try:
        token = get_token()

        if name == "list_email_folders":
            folders = list_folders(token)
            result = "Email Folders:\n"
            for f in sorted(folders, key=lambda x: x["displayName"].lower()):
                result += f"  - {f['displayName']} ({f['totalItemCount']} emails)\n"
            return [TextContent(type="text", text=result)]

        elif name == "get_recent_emails":
            folder = arguments.get("folder", "Inbox")
            days = arguments.get("days", 3)
            limit = arguments.get("limit", 20)
            unread_only = arguments.get("unread_only", False)
            flagged_only = arguments.get("flagged_only", False)

            folder_id = get_folder_id(token, folder)

            # Build filter
            date_filter = (datetime.utcnow() - timedelta(days=days)).strftime("%Y-%m-%dT00:00:00Z")
            filters = [f"receivedDateTime ge {date_filter}"]

            if unread_only:
                filters.append("isRead eq false")
            if flagged_only:
                filters.append("flag/flagStatus eq 'flagged'")

            filter_query = " and ".join(filters)

            # Fetch emails
            url = f"{GRAPH_API}/me/mailFolders/{folder_id}/messages"
            params = {
                "$orderby": "receivedDateTime desc",
                "$top": limit,
                "$filter": filter_query,
                "$select": "id,subject,from,receivedDateTime,isRead,flag,hasAttachments,body",
            }

            data = api_get(token, url, params)
            emails = data.get("value", [])

            if not emails:
                return [TextContent(type="text", text=f"No emails found in {folder} from the last {days} days.")]

            result = f"Found {len(emails)} emails in {folder} (last {days} days):\n"
            for email in emails:
                result += format_email_summary(email)

            return [TextContent(type="text", text=result)]

        elif name == "search_emails":
            folder = arguments.get("folder", "Inbox")
            filter_query = arguments["filter"]
            limit = arguments.get("limit", 20)

            folder_id = get_folder_id(token, folder)

            url = f"{GRAPH_API}/me/mailFolders/{folder_id}/messages"
            params = {
                "$orderby": "receivedDateTime desc",
                "$top": limit,
                "$filter": filter_query,
                "$select": "id,subject,from,receivedDateTime,isRead,flag,hasAttachments,body",
            }

            data = api_get(token, url, params)
            emails = data.get("value", [])

            if not emails:
                return [TextContent(type="text", text=f"No emails found matching filter: {filter_query}")]

            result = f"Found {len(emails)} emails:\n"
            for email in emails:
                result += format_email_summary(email)

            return [TextContent(type="text", text=result)]

        elif name == "get_email_by_id":
            email_id = arguments["email_id"]

            url = f"{GRAPH_API}/me/messages/{email_id}"
            params = {"$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments"}

            email = api_get(token, url, params)

            # Get attachments if any
            attachments_info = ""
            if email.get("hasAttachments"):
                attachments = get_attachments(token, email_id)
                if attachments:
                    attachments_info = "\n\nAttachments:\n"
                    for att in attachments:
                        attachments_info += f"  - {att.get('name')} ({att.get('size', 0)} bytes)\n"

            result = format_email_summary(email) + attachments_info
            return [TextContent(type="text", text=result)]

        else:
            return [TextContent(type="text", text=f"Unknown tool: {name}")]

    except Exception as e:
        return [TextContent(type="text", text=f"Error: {str(e)}")]


async def main():
    async with stdio_server() as (read_stream, write_stream):
        await server.run(read_stream, write_stream, server.create_initialization_options())


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
