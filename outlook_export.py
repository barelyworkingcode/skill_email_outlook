#!/usr/bin/env python3
"""Export Outlook.com emails to Markdown files via Microsoft Graph API."""

from __future__ import annotations

import argparse
import base64
import json
import re
import sys
from pathlib import Path

import html2text
import msal
import requests
from dateutil import parser as date_parser

GRAPH_API = "https://graph.microsoft.com/v1.0"
SCOPES = ["Mail.Read"]
TOKEN_CACHE = Path(__file__).parent / ".token_cache.json"
CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"

FILTER_HELP = """
Filter Examples (use with --filter):

  Flagged messages:
    --filter "flag/flagStatus eq 'flagged'"

  Unread messages:
    --filter "isRead eq false"

  Read messages:
    --filter "isRead eq true"

  High importance:
    --filter "importance eq 'high'"

  Has attachments:
    --filter "hasAttachments eq true"

  From specific sender:
    --filter "from/emailAddress/address eq 'someone@example.com'"

  After date:
    --filter "receivedDateTime ge 2024-01-01"

  Before date:
    --filter "receivedDateTime lt 2024-06-01"

  Date range:
    --filter "receivedDateTime ge 2024-01-01 and receivedDateTime lt 2024-07-01"

  Subject contains:
    --filter "contains(subject, 'invoice')"

  Combine with 'and'/'or':
    --filter "isRead eq false and importance eq 'high'"
    --filter "flag/flagStatus eq 'flagged' and receivedDateTime ge 2024-01-01"
"""


def clear_token_cache():
    """Delete token cache file."""
    if TOKEN_CACHE.exists():
        TOKEN_CACHE.unlink()
        print(f"Deleted {TOKEN_CACHE}")
    else:
        print("No token cache to delete.")


def get_token_from_refresh_token(refresh_token: str) -> str:
    """Get access token using a refresh token."""
    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority="https://login.microsoftonline.com/consumers",
    )

    result = app.acquire_token_by_refresh_token(refresh_token, scopes=SCOPES)

    if "access_token" not in result:
        raise Exception(f"Token refresh failed: {result.get('error_description', result)}")

    return result["access_token"]


def get_token_from_cache_file(cache_path: Path) -> str:
    """Get access token using a token cache file (auto-refreshes)."""
    if not cache_path.exists():
        raise Exception(f"Token cache not found: {cache_path}")

    cache = msal.SerializableTokenCache()
    cache.deserialize(cache_path.read_text())

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority="https://login.microsoftonline.com/consumers",
        token_cache=cache,
    )

    accounts = app.get_accounts()
    if not accounts:
        raise Exception("No accounts in token cache. Run --reauth first.")

    result = app.acquire_token_silent(SCOPES, account=accounts[0])

    if not result or "access_token" not in result:
        raise Exception("Token refresh failed. Run --reauth to re-authenticate.")

    # Save updated cache if changed
    if cache.has_state_changed:
        cache_path.write_text(cache.serialize())

    return result["access_token"]


def get_access_token(force_reauth: bool = False) -> str:
    """Authenticate using device code flow."""
    cache = msal.SerializableTokenCache()

    if not force_reauth and TOKEN_CACHE.exists():
        cache.deserialize(TOKEN_CACHE.read_text())

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority="https://login.microsoftonline.com/consumers",
        token_cache=cache,
    )

    if not force_reauth:
        accounts = app.get_accounts()
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            if result and "access_token" in result:
                return result["access_token"]

    # Device code flow
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception(f"Device flow failed: {flow.get('error_description', 'Unknown')}")

    print(f"\n{flow['message']}\n")
    result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise Exception(f"Auth failed: {result.get('error_description', result)}")

    if cache.has_state_changed:
        TOKEN_CACHE.write_text(cache.serialize())
        print(f"Token saved to {TOKEN_CACHE}")

    return result["access_token"]


def api_get(token: str, url: str, params: dict = None) -> dict:
    """Make authenticated GET request."""
    headers = {"Authorization": f"Bearer {token}"}
    resp = requests.get(url, headers=headers, params=params)
    resp.raise_for_status()
    return resp.json()


def list_folders(token: str) -> list[dict]:
    """List all mail folders."""
    url = f"{GRAPH_API}/me/mailFolders"
    folders = []
    while url:
        data = api_get(token, url)
        folders.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return folders


def get_folder_by_name(token: str, folder_name: str) -> dict | None:
    """Find folder by name (case-insensitive)."""
    for folder in list_folders(token):
        if folder["displayName"].lower() == folder_name.lower():
            return folder
    return None


def get_emails_batch(token: str, folder_id: str, skip: int = 0, limit: int = 50,
                     filter_query: str = None) -> tuple[list[dict], bool]:
    """Fetch a batch of emails. Returns (emails, has_more)."""
    url = f"{GRAPH_API}/me/mailFolders/{folder_id}/messages"
    params = {
        "$orderby": "receivedDateTime desc",
        "$top": limit,
        "$skip": skip,
        "$select": "id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,hasAttachments,conversationId",
    }
    if filter_query:
        params["$filter"] = filter_query

    data = api_get(token, url, params)
    emails = data.get("value", [])
    has_more = "@odata.nextLink" in data or len(emails) == limit
    return emails, has_more


def get_attachments(token: str, message_id: str) -> list[dict]:
    """Fetch attachments for a message."""
    url = f"{GRAPH_API}/me/messages/{message_id}/attachments"
    data = api_get(token, url)
    return data.get("value", [])


def sanitize_filename(name: str, max_length: int = 80) -> str:
    """Convert to safe filename."""
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '', name)
    name = re.sub(r'\s+', ' ', name).strip()
    if len(name) > max_length:
        name = name[:max_length].rsplit(' ', 1)[0]
    return name or "untitled"


def load_manifest(output_dir: Path) -> dict:
    """Load manifest with downloaded IDs and thread info."""
    manifest_path = output_dir / ".manifest.json"
    if manifest_path.exists():
        return json.loads(manifest_path.read_text())
    return {"downloaded_ids": [], "threads": {}}


def save_manifest(output_dir: Path, manifest: dict):
    """Save manifest."""
    manifest_path = output_dir / ".manifest.json"
    manifest_path.write_text(json.dumps(manifest, indent=2))


def email_to_markdown(email: dict, attachment_files: list[str] = None) -> str:
    """Convert email to Markdown with YAML frontmatter."""
    received = email.get("receivedDateTime", "")
    if received:
        dt = date_parser.parse(received)
        date_str = dt.strftime("%Y-%m-%d %H:%M:%S %z")
    else:
        date_str = "Unknown"

    from_addr = email.get("from", {}).get("emailAddress", {})
    from_str = f"{from_addr.get('name', '')} <{from_addr.get('address', '')}>".strip()

    to_list = email.get("toRecipients", [])
    to_str = ", ".join(
        f"{r.get('emailAddress', {}).get('name', '')} <{r.get('emailAddress', {}).get('address', '')}>".strip()
        for r in to_list
    )

    cc_list = email.get("ccRecipients", [])
    cc_str = ", ".join(
        f"{r.get('emailAddress', {}).get('name', '')} <{r.get('emailAddress', {}).get('address', '')}>".strip()
        for r in cc_list
    ) if cc_list else None

    subject = email.get("subject", "(No Subject)")

    def yaml_escape(s: str) -> str:
        return s.replace('\\', '\\\\').replace('"', '\\"')

    lines = [
        "---",
        f'subject: "{yaml_escape(subject)}"',
        f'from: "{yaml_escape(from_str)}"',
        f'to: "{yaml_escape(to_str)}"',
    ]
    if cc_str:
        lines.append(f'cc: "{yaml_escape(cc_str)}"')
    lines.append(f'date: "{date_str}"')
    lines.append(f'id: "{email.get("id", "")}"')
    lines.append(f'conversation_id: "{email.get("conversationId", "")}"')

    if attachment_files:
        lines.append("attachments:")
        for af in attachment_files:
            lines.append(f'  - "{af}"')

    lines.extend(["---", "", f"# {subject}", ""])

    body = email.get("body", {})
    content = body.get("content", "")

    if body.get("contentType") == "html":
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.ignore_images = False
        h.body_width = 0
        content = h.handle(content)

    lines.append(content.strip())

    if attachment_files:
        lines.extend(["", "---", "## Attachments", ""])
        for af in attachment_files:
            lines.append(f"- [{af}](attachments/{af})")

    return "\n".join(lines)


def get_thread_folder_name(email: dict, manifest: dict) -> str:
    """Get or create thread folder name for a conversation."""
    conv_id = email.get("conversationId", "")

    if conv_id in manifest["threads"]:
        return manifest["threads"][conv_id]

    # Create new thread folder name from this email
    received = email.get("receivedDateTime", "")
    if received:
        dt = date_parser.parse(received)
        date_prefix = dt.strftime("%Y%m%d")
    else:
        date_prefix = "unknown"

    subject = email.get("subject", "untitled")
    # Remove Re:/Fwd: prefixes for cleaner folder names
    clean_subject = re.sub(r'^(Re|Fwd|Fw):\s*', '', subject, flags=re.IGNORECASE).strip()
    clean_subject = sanitize_filename(clean_subject, max_length=60)

    folder_name = f"{date_prefix}_{clean_subject}"
    manifest["threads"][conv_id] = folder_name
    return folder_name


def export_single_email(token: str, email: dict, output_dir: Path, manifest: dict,
                        use_threads: bool) -> str:
    """Export a single email with attachments. Returns the created filename."""
    received = email.get("receivedDateTime", "")
    if received:
        dt = date_parser.parse(received)
        date_prefix = dt.strftime("%Y%m%d_%H%M%S")
    else:
        date_prefix = "unknown"

    subject = sanitize_filename(email.get("subject", "untitled"))

    # Determine base directory
    if use_threads:
        thread_folder = get_thread_folder_name(email, manifest)
        base_dir = output_dir / thread_folder
        base_name = date_prefix  # Just timestamp within thread
    else:
        base_dir = output_dir
        base_name = f"{date_prefix}_{subject}"

    # Handle attachments
    attachment_files = []
    if email.get("hasAttachments"):
        attachments = get_attachments(token, email["id"])
        if attachments:
            attach_dir = base_dir / "attachments"
            attach_dir.mkdir(parents=True, exist_ok=True)

            for att in attachments:
                if att.get("@odata.type") == "#microsoft.graph.fileAttachment":
                    att_name = sanitize_filename(att.get("name", "attachment"))
                    # Prefix with timestamp to avoid collisions in thread mode
                    if use_threads:
                        att_name = f"{date_prefix}_{att_name}"
                    att_path = attach_dir / att_name

                    counter = 1
                    while att_path.exists():
                        stem = Path(att_name).stem
                        suffix = Path(att_name).suffix
                        att_path = attach_dir / f"{stem}_{counter}{suffix}"
                        counter += 1

                    content_bytes = base64.b64decode(att.get("contentBytes", ""))
                    att_path.write_bytes(content_bytes)
                    attachment_files.append(att_path.name)

    # Write markdown
    markdown = email_to_markdown(email, attachment_files)

    base_dir.mkdir(parents=True, exist_ok=True)
    filepath = base_dir / f"{base_name}.md"

    counter = 1
    while filepath.exists():
        filepath = base_dir / f"{base_name}_{counter}.md"
        counter += 1

    filepath.write_text(markdown, encoding="utf-8")
    return filepath.relative_to(output_dir).as_posix()


def export_emails(token: str, folder_name: str, output_dir: Path,
                  limit: int | None = None, batch_size: int = 50, skip: int = 0,
                  filter_query: str = None, use_threads: bool = False):
    """Export emails from a folder."""
    folder = get_folder_by_name(token, folder_name)
    if not folder:
        print(f"Error: Folder '{folder_name}' not found.", file=sys.stderr)
        print("\nAvailable folders:", file=sys.stderr)
        for f in list_folders(token):
            print(f"  - {f['displayName']}", file=sys.stderr)
        sys.exit(1)

    total = folder["totalItemCount"]
    print(f"Folder: {folder['displayName']} ({total} total emails)")
    print(f"Batch size: {batch_size}, Skip: {skip}, Limit: {limit or 'all'}")
    if filter_query:
        print(f"Filter: {filter_query}")
    if use_threads:
        print("Thread mode: enabled (grouping by conversation)")

    output_dir.mkdir(parents=True, exist_ok=True)
    manifest = load_manifest(output_dir)
    downloaded_ids = set(manifest.get("downloaded_ids", []))
    print(f"Already downloaded: {len(downloaded_ids)} emails")

    exported = 0
    skipped_dup = 0
    current_skip = skip
    target = limit if limit else total

    while exported < target:
        batch_limit = min(batch_size, target - exported)
        print(f"\nFetching batch at offset {current_skip}...")

        try:
            emails, has_more = get_emails_batch(token, folder["id"], current_skip,
                                                 batch_limit, filter_query)
        except requests.HTTPError as e:
            print(f"API Error: {e}", file=sys.stderr)
            if "filter" in str(e).lower():
                print("Check your filter syntax. Use --filter-help for examples.", file=sys.stderr)
            break

        if not emails:
            print("No more emails.")
            break

        for email in emails:
            msg_id = email["id"]

            if msg_id in downloaded_ids:
                skipped_dup += 1
                continue

            try:
                filename = export_single_email(token, email, output_dir, manifest, use_threads)
                downloaded_ids.add(msg_id)
                exported += 1
                print(f"  [{exported}/{target}] {filename}")

                if exported % 10 == 0:
                    manifest["downloaded_ids"] = sorted(downloaded_ids)
                    save_manifest(output_dir, manifest)

            except Exception as e:
                print(f"  Error exporting {email.get('subject', 'unknown')}: {e}", file=sys.stderr)

            if limit and exported >= limit:
                break

        current_skip += len(emails)

        if not has_more:
            break

    manifest["downloaded_ids"] = sorted(downloaded_ids)
    save_manifest(output_dir, manifest)
    print(f"\nExported: {exported}, Skipped (duplicates): {skipped_dup}")
    print(f"Output: {output_dir}")


def main():
    parser = argparse.ArgumentParser(
        description="Export Outlook.com emails to Markdown files",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --list-folders
  %(prog)s -f Inbox
  %(prog)s -f Inbox --batch-size 25 --limit 100
  %(prog)s -f Inbox --filter "flag/flagStatus eq 'flagged'"
  %(prog)s -f Inbox --threads
  %(prog)s --filter-help
  %(prog)s --reauth

Agent usage (run --auth-only first to get tokens):
  %(prog)s --refresh-token <TOKEN> -f Inbox  # Best: ~90 day token, auto-refreshes
  %(prog)s --token-cache /path/to/cache.json -f Inbox  # Use cache file
  %(prog)s --token <TOKEN> -f Inbox          # Short-lived ~1 hour token
        """,
    )
    parser.add_argument("--list-folders", action="store_true", help="List mail folders")
    parser.add_argument("--folder", "-f", help="Folder name to export")
    parser.add_argument("--output", "-o", type=Path, help="Output directory (default: ./mail/{folder})")
    parser.add_argument("--limit", "-n", type=int, help="Max emails to export")
    parser.add_argument("--batch-size", "-b", type=int, default=50, help="Emails per API request (default: 50)")
    parser.add_argument("--skip", "-s", type=int, default=0, help="Skip first N emails")
    parser.add_argument("--filter", dest="filter_query", help="OData filter query")
    parser.add_argument("--filter-help", "-?", action="store_true", help="Show filter examples")
    parser.add_argument("--threads", "-t", action="store_true", help="Group emails by conversation thread")
    parser.add_argument("--reauth", action="store_true", help="Clear token cache and re-authenticate")
    parser.add_argument("--auth-only", action="store_true", help="Authenticate and print tokens, then exit")
    parser.add_argument("--token", dest="access_token", help="Use provided access token (~1 hour)")
    parser.add_argument("--refresh-token", dest="refresh_token", help="Use refresh token to get access token (~90 days)")
    parser.add_argument("--token-cache", dest="token_cache_path", type=Path, help="Path to token cache file")

    args = parser.parse_args()

    if args.filter_help:
        print(FILTER_HELP)
        sys.exit(0)

    # Auth-only mode: authenticate and print tokens
    if args.auth_only:
        if args.reauth:
            clear_token_cache()
        print("Authenticating...\n")
        token = get_access_token(force_reauth=args.reauth)

        # Extract refresh token from cache
        refresh_token = None
        if TOKEN_CACHE.exists():
            cache_data = json.loads(TOKEN_CACHE.read_text())
            rt_entries = cache_data.get("RefreshToken", {})
            if rt_entries:
                refresh_token = list(rt_entries.values())[0].get("secret")

        print(f"\n=== ACCESS TOKEN (expires ~1 hour) ===\n{token}\n")
        if refresh_token:
            print(f"=== REFRESH TOKEN (expires ~90 days) ===\n{refresh_token}\n")
            print("Agent usage:")
            print("  --token <access_token>           # Short-lived, ~1 hour")
            print("  --refresh-token <refresh_token>  # Long-lived, ~90 days, auto-refreshes")
            print(f"  --token-cache {TOKEN_CACHE}  # Use cache file directly")
        sys.exit(0)

    # Reauth mode
    if args.reauth:
        clear_token_cache()
        print("\nStarting fresh authentication...\n")
        print("Instructions:")
        print("  1. A URL and code will be displayed below")
        print("  2. Open the URL in your browser")
        print("  3. Enter the code when prompted")
        print("  4. Sign in with your Microsoft account")
        print("  5. The token will be saved automatically\n")
        token = get_access_token(force_reauth=True)
        print("\nAuthentication successful.")
        if not args.list_folders and not args.folder:
            sys.exit(0)

    # Use provided token or get from cache
    if args.access_token:
        token = args.access_token
        print("Using provided access token.\n")
    elif args.refresh_token:
        print("Refreshing access token...")
        token = get_token_from_refresh_token(args.refresh_token)
        print("Token refreshed.\n")
    elif args.token_cache_path:
        print(f"Using token cache: {args.token_cache_path}")
        token = get_token_from_cache_file(args.token_cache_path)
        print("Authenticated.\n")
    elif not args.reauth:
        if not args.list_folders and not args.folder:
            parser.print_help()
            sys.exit(1)
        print("Authenticating...")
        token = get_access_token()
        print("Authenticated.\n")

    if args.list_folders:
        folders = list_folders(token)
        print("Mail Folders:")
        for folder in sorted(folders, key=lambda f: f["displayName"].lower()):
            print(f"  {folder['displayName']} ({folder['totalItemCount']} emails)")
    else:
        output_dir = args.output or Path("./mail") / sanitize_filename(args.folder)
        export_emails(token, args.folder, output_dir, args.limit, args.batch_size,
                      args.skip, args.filter_query, args.threads)


if __name__ == "__main__":
    main()
