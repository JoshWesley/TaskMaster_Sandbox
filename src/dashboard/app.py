"""Dash application for the Outlook Intelligence Dashboard.

Connects to Outlook via COM on startup, syncs emails into SQLite,
and displays live data about what has been pulled in.

On first run: full sync with progress bar.
On subsequent runs: loads cached data instantly, then incrementally syncs new emails only.
"""

from __future__ import annotations

import sqlite3
import sys
import time
import traceback
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

try:
    from dash import Dash, Input, Output, dash_table, dcc, html
except ImportError:  # pragma: no cover
    Dash = None  # type: ignore[assignment]
    Input = Output = dash_table = dcc = html = None  # type: ignore[assignment]

# Ensure src is importable when running this file directly.
_SRC_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(_SRC_ROOT))

from src.conversations.engine import ConversationEngine
from src.outlook.extractor import OutlookExtractor
from src.storage.database import DatabaseManager

DB_PATH = "database/dashboard.db"


def _print_progress(current: int, total: int, folder: str, start_time: float) -> None:
    """Print a terminal progress bar."""
    if total == 0:
        return
    pct = current / total
    bar_len = 40
    filled = int(bar_len * pct)
    bar = "█" * filled + "░" * (bar_len - filled)
    elapsed = time.time() - start_time
    rate = current / elapsed if elapsed > 0 else 0
    eta = (total - current) / rate if rate > 0 else 0
    sys.stdout.write(
        f"\r  [{bar}] {current}/{total} ({pct:.0%}) | {folder} | {rate:.0f} emails/sec | ETA: {eta:.0f}s  "
    )
    sys.stdout.flush()


def _has_cached_data(db: DatabaseManager) -> bool:
    """Check if we have previously synced data in SQLite."""
    try:
        with db.connect() as conn:
            row = conn.execute("SELECT COUNT(*) FROM emails").fetchone()
            return row[0] > 0
    except (sqlite3.OperationalError, Exception):
        return False


def _load_from_cache(db: DatabaseManager) -> dict[str, Any]:
    """Load existing synced data from SQLite for instant startup."""
    report: dict[str, Any] = {
        "connected": True,
        "error": None,
        "folders_synced": [],
        "total_emails": 0,
        "total_conversations": 0,
        "emails_by_folder": {},
        "conversations": [],
        "ownership_counts": Counter(),
        "type_counts": Counter(),
        "top_senders": Counter(),
        "top_domains": Counter(),
        "unread_count": 0,
        "flagged_count": 0,
        "sync_time": None,
        "from_cache": True,
    }

    with db.connect() as conn:
        # Total emails
        report["total_emails"] = conn.execute("SELECT COUNT(*) FROM emails").fetchone()[0]

        # Emails by folder
        rows = conn.execute("SELECT folder, COUNT(*) as cnt FROM emails GROUP BY folder").fetchall()
        for row in rows:
            report["emails_by_folder"][row[0]] = row[1]
            report["folders_synced"].append(row[0])

        # Unread / flagged
        report["unread_count"] = conn.execute(
            "SELECT COUNT(*) FROM emails WHERE read_status = 'Unread'"
        ).fetchone()[0]
        report["flagged_count"] = conn.execute(
            "SELECT COUNT(*) FROM emails WHERE flag_status NOT IN ('0', '', 'None') AND flag_status IS NOT NULL"
        ).fetchone()[0]

        # Top senders
        rows = conn.execute(
            "SELECT sender_name, COUNT(*) as cnt FROM emails GROUP BY sender_name ORDER BY cnt DESC LIMIT 10"
        ).fetchall()
        for row in rows:
            report["top_senders"][row[0]] = row[1]

        # Top domains
        rows = conn.execute(
            "SELECT sender_domain, COUNT(*) as cnt FROM emails GROUP BY sender_domain ORDER BY cnt DESC LIMIT 10"
        ).fetchall()
        for row in rows:
            report["top_domains"][row[0]] = row[1]

        # Last sync time
        row = conn.execute("SELECT MAX(last_synced_at) FROM sync_state").fetchone()
        report["sync_time"] = row[0] if row and row[0] else "Unknown"

        # Conversations from DB
        conv_rows = conn.execute("SELECT * FROM conversations ORDER BY last_activity DESC").fetchall()
        report["total_conversations"] = len(conv_rows)
        for crow in conv_rows:
            crow_dict = dict(crow)
            report["ownership_counts"][crow_dict.get("ownership", "Unknown")] += 1
            report["type_counts"][crow_dict.get("conversation_type", "Unknown")] += 1

        # Build conversation objects for the table
        engine = ConversationEngine()
        email_rows = conn.execute("SELECT * FROM emails ORDER BY received_time DESC").fetchall()
        email_dicts = [dict(r) for r in email_rows]
        if email_dicts:
            report["conversations"] = engine.build_conversations(email_dicts)

    return report


def _get_last_sync_time(db: DatabaseManager, folder_name: str) -> str | None:
    """Get the last sync timestamp for a folder."""
    try:
        with db.connect() as conn:
            row = conn.execute(
                "SELECT last_synced_at FROM sync_state WHERE folder_name = ?", (folder_name,)
            ).fetchone()
            return row[0] if row else None
    except (sqlite3.OperationalError, Exception):
        return None


def _sync_outlook(limit_per_folder: int | None = None, force_full: bool = False) -> dict[str, Any]:
    """Connect to Outlook, extract emails with progress bar, and return a status report.

    Uses incremental sync on subsequent runs — only fetches emails newer than
    the last sync time. Data is persisted to SQLite for fast future startups.
    Set limit_per_folder to cap emails per folder (None = sync all).
    """

    db = DatabaseManager(DB_PATH)
    db.initialize()

    # If we have cached data and not forcing full sync, load cache first then sync incrementally
    has_cache = _has_cached_data(db)
    if has_cache and not force_full:
        print("  ⚡ Loading cached data from previous sync...")

    report: dict[str, Any] = {
        "connected": False,
        "error": None,
        "folders_synced": [],
        "total_emails": 0,
        "total_conversations": 0,
        "emails_by_folder": {},
        "conversations": [],
        "ownership_counts": Counter(),
        "type_counts": Counter(),
        "top_senders": Counter(),
        "top_domains": Counter(),
        "unread_count": 0,
        "flagged_count": 0,
        "sync_time": None,
        "from_cache": False,
        "new_emails_synced": 0,
    }

    try:
        extractor = OutlookExtractor()
        print("  🔌 Connecting to Outlook...")
        extractor.connect()
        report["connected"] = True
        print("  ✅ Connected to Outlook")

        # Sync each folder with progress
        all_records = []
        folders = extractor.folder_names
        for folder_name in folders:
            last_sync = _get_last_sync_time(db, folder_name) if has_cache and not force_full else None

            if last_sync:
                print(f"\n  📂 {folder_name} (incremental since {last_sync})")
            else:
                label = f"up to {limit_per_folder}" if limit_per_folder else "all emails"
                print(f"\n  📂 {folder_name} (full sync, {label})")

            start_time = time.time()
            count = 0
            folder_records = []

            # Get total count for progress bar when doing full sync
            folder_total = limit_per_folder  # Use as estimate if set
            if not limit_per_folder:
                try:
                    namespace = extractor.connect()
                    folder_obj = namespace.Folders.Item(1).Folders[folder_name]
                    folder_total = folder_obj.Items.Count
                except Exception:
                    folder_total = 0  # Unknown — progress will show count only

            for item in extractor.iter_folder_items(folder_name, limit=limit_per_folder):
                record = extractor.extract_message(item, folder_name)

                # Incremental: skip emails we've already synced
                if last_sync and record.received_time:
                    if record.received_time.isoformat() <= last_sync:
                        # We've reached emails we already have — stop
                        break

                folder_records.append(record)
                count += 1
                if folder_total and folder_total > 0:
                    _print_progress(count, folder_total, folder_name, start_time)
                else:
                    # Unknown total — just show count and rate
                    elapsed = time.time() - start_time
                    rate = count / elapsed if elapsed > 0 else 0
                    sys.stdout.write(f"\r  ⏳ {count} emails synced | {folder_name} | {rate:.0f}/sec  ")
                    sys.stdout.flush()

            # End of folder
            elapsed = time.time() - start_time
            print(f"\n  ✓ {folder_name}: {count} emails in {elapsed:.1f}s")

            all_records.extend(folder_records)

            # Persist records to SQLite
            for record in folder_records:
                db.upsert_email(record.to_dict())

            # Update sync state
            db.update_sync_state(
                folder_name,
                last_synced_at=datetime.now().isoformat(),
                last_entry_id=folder_records[0].entry_id if folder_records else "",
            )

        report["new_emails_synced"] = len(all_records)
        report["sync_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Now build report from ALL data in SQLite (cached + new)
        print("\n  📊 Building conversation intelligence...")
        cached_report = _load_from_cache(db)
        report["total_emails"] = cached_report["total_emails"]
        report["total_conversations"] = cached_report["total_conversations"]
        report["emails_by_folder"] = cached_report["emails_by_folder"]
        report["folders_synced"] = cached_report["folders_synced"]
        report["ownership_counts"] = cached_report["ownership_counts"]
        report["type_counts"] = cached_report["type_counts"]
        report["top_senders"] = cached_report["top_senders"]
        report["top_domains"] = cached_report["top_domains"]
        report["unread_count"] = cached_report["unread_count"]
        report["flagged_count"] = cached_report["flagged_count"]
        report["conversations"] = cached_report["conversations"]

        # Persist conversations
        engine = ConversationEngine()
        with db.connect() as conn:
            email_rows = conn.execute("SELECT * FROM emails ORDER BY received_time DESC").fetchall()
            email_dicts = [dict(r) for r in email_rows]
        conversations = engine.build_conversations(email_dicts)
        for conv in conversations:
            db.upsert_conversation(conv.to_dict())

        print(f"  ✅ Sync complete: {report['new_emails_synced']} new emails, {report['total_emails']} total in database")

    except Exception as e:
        report["error"] = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"
        print(f"\n  ❌ Error: {type(e).__name__}: {e}")

        # If sync failed but we have cached data, use it
        if has_cache:
            print("  ⚡ Falling back to cached data...")
            cached = _load_from_cache(db)
            cached["error"] = report["error"]
            cached["connected"] = False
            return cached

    return report


def create_app(sync_limit: int | None = None) -> "Dash":
    """Create the Dash app with live Outlook data."""

    if Dash is None:
        raise RuntimeError("Dash is not installed. Install requirements.txt to run the dashboard UI.")

    # Sync data from Outlook (incremental if cache exists)
    report = _sync_outlook(limit_per_folder=sync_limit)

    app = Dash(__name__)
    app.layout = _build_layout(report)

    # Callback: filter conversations table by ownership
    @app.callback(
        Output("conversations-table", "data"),
        Input("ownership-filter", "value"),
        Input("conversations-store", "data"),
    )
    def filter_conversations(ownership: str, conversations: list[dict[str, Any]]) -> list[dict[str, Any]]:
        if ownership == "All":
            return conversations
        return [c for c in conversations if c.get("ownership") == ownership]

    return app


def _build_layout(report: dict[str, Any]) -> Any:
    """Build the full dashboard layout from the sync report."""

    # Connection status banner
    if report["connected"]:
        sync_detail = f"Last sync: {report['sync_time']}"
        if report.get("new_emails_synced", 0) > 0:
            sync_detail += f"  •  {report['new_emails_synced']} new emails synced"
        elif report.get("from_cache"):
            sync_detail += "  •  Loaded from cache (instant startup)"
        status_banner = html.Div(
            [
                html.Span("✅ Connected to Outlook", style={"fontWeight": "bold", "color": "#2e7d32"}),
                html.Span(f"  •  {sync_detail}", style={"marginLeft": "16px", "color": "#555"}),
            ],
            style={"padding": "12px 16px", "backgroundColor": "#e8f5e9", "borderRadius": "8px", "marginBottom": "24px"},
        )
    else:
        error_msg = report["error"] or "Unknown error"
        status_banner = html.Div(
            [
                html.Span("❌ Failed to connect to Outlook", style={"fontWeight": "bold", "color": "#c62828"}),
                html.Pre(error_msg, style={"fontSize": "12px", "marginTop": "8px", "whiteSpace": "pre-wrap"}),
            ],
            style={"padding": "12px 16px", "backgroundColor": "#ffebee", "borderRadius": "8px", "marginBottom": "24px"},
        )

    # Data volume section
    data_summary = html.Div(
        [
            html.H2("📊 Data Summary"),
            html.Div(
                [
                    _metric_card("Total Emails Synced", report["total_emails"], "#e3f2fd"),
                    _metric_card("Conversations", report["total_conversations"], "#f3e5f5"),
                    _metric_card("Folders Synced", len(report["folders_synced"]), "#fff3e0"),
                    _metric_card("Unread", report["unread_count"], "#fce4ec"),
                    _metric_card("Flagged", report["flagged_count"], "#fff8e1"),
                ],
                style={"display": "grid", "gridTemplateColumns": "repeat(5, 1fr)", "gap": "12px", "marginBottom": "24px"},
            ),
        ]
    )

    # Folder breakdown
    folder_rows = [{"folder": k, "emails": v} for k, v in sorted(report["emails_by_folder"].items(), key=lambda x: -x[1])]
    folder_section = html.Div(
        [
            html.H3("📁 Emails by Folder"),
            dash_table.DataTable(
                columns=[{"name": "Folder", "id": "folder"}, {"name": "Emails", "id": "emails"}],
                data=folder_rows,
                style_table={"maxWidth": "400px"},
                style_cell={"textAlign": "left", "padding": "8px"},
            ),
        ],
        style={"marginBottom": "24px"},
    )

    # Ownership breakdown
    ownership_counts = report["ownership_counts"]
    ownership_section = html.Div(
        [
            html.H2("🎯 Ownership Breakdown"),
            html.Div(
                [
                    _metric_card("Awaiting My Action", ownership_counts.get("Awaiting My Action", 0), "#ffcdd2"),
                    _metric_card("Waiting For Others", ownership_counts.get("Waiting For Others", 0), "#c8e6c9"),
                    _metric_card("Information Only", ownership_counts.get("Information Only", 0), "#bbdefb"),
                    _metric_card("Unknown", ownership_counts.get("Unknown", 0), "#e0e0e0"),
                ],
                style={"display": "grid", "gridTemplateColumns": "repeat(4, 1fr)", "gap": "12px", "marginBottom": "24px"},
            ),
        ]
    )

    # Conversation type breakdown
    type_counts = report["type_counts"]
    type_section = html.Div(
        [
            html.H3("🌐 Conversation Types"),
            html.Div(
                [
                    _metric_card("Internal", type_counts.get("Internal", 0), "#e8f5e9"),
                    _metric_card("External", type_counts.get("External", 0), "#e3f2fd"),
                    _metric_card("Mixed", type_counts.get("Mixed", 0), "#fff3e0"),
                ],
                style={"display": "grid", "gridTemplateColumns": "repeat(3, 1fr)", "gap": "12px", "marginBottom": "24px"},
            ),
        ]
    )

    # Top senders
    top_senders = report["top_senders"].most_common(10)
    sender_rows = [{"sender": name, "emails": count} for name, count in top_senders]
    sender_section = html.Div(
        [
            html.H3("👤 Top Senders"),
            dash_table.DataTable(
                columns=[{"name": "Sender", "id": "sender"}, {"name": "Emails", "id": "emails"}],
                data=sender_rows,
                style_table={"maxWidth": "500px"},
                style_cell={"textAlign": "left", "padding": "8px"},
            ),
        ],
        style={"marginBottom": "24px"},
    )

    # Top domains
    top_domains = report["top_domains"].most_common(10)
    domain_rows = [{"domain": domain, "emails": count} for domain, count in top_domains]
    domain_section = html.Div(
        [
            html.H3("🏢 Top Domains"),
            dash_table.DataTable(
                columns=[{"name": "Domain", "id": "domain"}, {"name": "Emails", "id": "emails"}],
                data=domain_rows,
                style_table={"maxWidth": "500px"},
                style_cell={"textAlign": "left", "padding": "8px"},
            ),
        ],
        style={"marginBottom": "24px"},
    )

    # Conversations table with filter
    conv_data = []
    for conv in report["conversations"]:
        conv_data.append({
            "topic": conv.topic,
            "ownership": conv.ownership,
            "type": conv.conversation_type,
            "messages": conv.message_count,
            "latest_sender": conv.latest_sender,
            "last_activity": conv.last_activity.strftime("%Y-%m-%d %H:%M") if conv.last_activity else "",
            "action": "⚠️ Yes" if conv.action_required else "",
        })

    conversations_section = html.Div(
        [
            html.H2("💬 Conversations"),
            dcc.Dropdown(
                id="ownership-filter",
                options=[
                    {"label": "All", "value": "All"},
                    {"label": "Awaiting My Action", "value": "Awaiting My Action"},
                    {"label": "Waiting For Others", "value": "Waiting For Others"},
                    {"label": "Information Only", "value": "Information Only"},
                    {"label": "Unknown", "value": "Unknown"},
                ],
                value="All",
                clearable=False,
                style={"maxWidth": "300px", "marginBottom": "12px"},
            ),
            dash_table.DataTable(
                id="conversations-table",
                columns=[
                    {"name": "Topic", "id": "topic"},
                    {"name": "Ownership", "id": "ownership"},
                    {"name": "Type", "id": "type"},
                    {"name": "Messages", "id": "messages"},
                    {"name": "Latest Sender", "id": "latest_sender"},
                    {"name": "Last Activity", "id": "last_activity"},
                    {"name": "Action?", "id": "action"},
                ],
                data=conv_data,
                page_size=20,
                sort_action="native",
                filter_action="native",
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "left", "padding": "8px", "maxWidth": "300px", "overflow": "hidden", "textOverflow": "ellipsis"},
                style_header={"fontWeight": "bold", "backgroundColor": "#f5f5f5"},
            ),
            dcc.Store(id="conversations-store", data=conv_data),
        ]
    )

    return html.Div(
        [
            html.H1("📬 Outlook Intelligence Dashboard"),
            status_banner,
            data_summary,
            folder_section,
            ownership_section,
            type_section,
            sender_section,
            domain_section,
            conversations_section,
        ],
        style={"maxWidth": "1400px", "margin": "0 auto", "padding": "24px", "fontFamily": "Segoe UI, sans-serif"},
    )


def _metric_card(title: str, value: Any, bg_color: str = "#FAFAFA") -> Any:
    return html.Div(
        [
            html.P(title, style={"margin": "0 0 4px 0", "fontSize": "13px", "color": "#555"}),
            html.P(str(value), style={"margin": "0", "fontSize": "32px", "fontWeight": "bold"}),
        ],
        style={
            "border": "1px solid #e0e0e0",
            "borderRadius": "8px",
            "padding": "16px",
            "backgroundColor": bg_color,
            "textAlign": "center",
        },
    )


if __name__ == "__main__":  # pragma: no cover
    print("\n" + "=" * 60)
    print("  📬 Outlook Intelligence Dashboard")
    print("=" * 60)
    print("\n  Syncing with Outlook...\n")
    app = create_app()
    print("\n" + "=" * 60)
    print("  🚀 Dashboard ready at http://127.0.0.1:8050")
    print("=" * 60 + "\n")
    app.run(debug=True)
