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
    from dash import Dash, Input, Output, State, callback_context, dash_table, dcc, html, no_update
except ImportError:  # pragma: no cover
    Dash = None  # type: ignore[assignment]
    Input = Output = State = callback_context = dash_table = dcc = html = no_update = None  # type: ignore[assignment]

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

            # Check if previous sync was incomplete
            # We use entry_id dedup instead of timestamp-based incremental
            # to handle resuming a partial sync of older emails
            existing_entry_ids: set[str] = set()
            with db.connect() as conn:
                rows = conn.execute(
                    "SELECT entry_id FROM emails WHERE folder = ?", (folder_name,)
                ).fetchall()
                existing_entry_ids = {row[0] for row in rows}

            already_synced = len(existing_entry_ids)
            if already_synced > 0:
                print(f"\n  📂 {folder_name} ({already_synced:,} already in database, syncing remaining...)")
            else:
                label = f"up to {limit_per_folder:,}" if limit_per_folder else "all emails"
                print(f"\n  📂 {folder_name} (full sync, {label})")

            start_time = time.time()
            count = 0
            skipped = 0
            folder_records = []

            for item in extractor.iter_folder_items(folder_name, limit=limit_per_folder):
                record = extractor.extract_message(item, folder_name)

                # Skip emails we already have in the database
                if record.entry_id in existing_entry_ids:
                    skipped += 1
                    # Once we hit a run of 50 consecutive already-synced emails, stop
                    # (we've reached the point where previous sync left off)
                    if already_synced > 0 and skipped > already_synced:
                        break
                    continue

                skipped = 0  # Reset consecutive skip counter
                folder_records.append(record)
                count += 1
                elapsed = time.time() - start_time
                rate = count / elapsed if elapsed > 0 else 0
                sys.stdout.write(
                    f"\r  ⏳ {count:,} new | {already_synced + count:,} total | {folder_name} | {rate:.0f}/sec     "
                )
                sys.stdout.flush()

                # Persist in batches of 100 to avoid losing progress on interrupt
                if count % 100 == 0:
                    for rec in folder_records[-100:]:
                        db.upsert_email(rec.to_dict())

            # End of folder
            elapsed = time.time() - start_time
            print(f"\n  ✓ {folder_name}: {count:,} new emails in {elapsed:.1f}s ({already_synced + count:,} total)")

            all_records.extend(folder_records)

            # Persist any remaining records not yet saved
            remaining = folder_records[-(count % 100):] if count % 100 != 0 else []
            for record in remaining:
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


def _metric_card_id(title: str) -> str:
    slug = "".join(char.lower() if char.isalnum() else "-" for char in title)
    while "--" in slug:
        slug = slug.replace("--", "-")
    return f"metric-{slug.strip('-')}"


def _metric_filters() -> dict[str, dict[str, Any]]:
    return {
        _metric_card_id("Total Emails Synced"): {"kind": "all", "value": None, "label": "All Conversations"},
        _metric_card_id("Conversations"): {"kind": "all", "value": None, "label": "All Conversations"},
        _metric_card_id("Folders Synced"): {"kind": "all", "value": None, "label": "All Conversations"},
        _metric_card_id("Unread"): {"kind": "has_unread", "value": True, "label": "Unread"},
        _metric_card_id("Flagged"): {"kind": "has_flagged", "value": True, "label": "Flagged"},
        _metric_card_id("Awaiting My Action"): {
            "kind": "ownership",
            "value": "Awaiting My Action",
            "label": "Awaiting My Action",
        },
        _metric_card_id("Waiting For Others"): {
            "kind": "ownership",
            "value": "Waiting For Others",
            "label": "Waiting For Others",
        },
        _metric_card_id("Information Only"): {
            "kind": "ownership",
            "value": "Information Only",
            "label": "Information Only",
        },
        _metric_card_id("Unknown"): {"kind": "ownership", "value": "Unknown", "label": "Unknown Ownership"},
        _metric_card_id("Internal"): {"kind": "type", "value": "Internal", "label": "Internal"},
        _metric_card_id("External"): {"kind": "type", "value": "External", "label": "External"},
        _metric_card_id("Mixed"): {"kind": "type", "value": "Mixed", "label": "Mixed"},
    }


def _load_conversation_flags() -> dict[str, dict[str, bool]]:
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        rows = conn.execute(
            """
            SELECT
                conversation_id,
                MAX(CASE WHEN read_status = 'Unread' THEN 1 ELSE 0 END) AS has_unread,
                MAX(
                    CASE
                        WHEN flag_status NOT IN ('0', '', 'None') AND flag_status IS NOT NULL THEN 1
                        ELSE 0
                    END
                ) AS has_flagged
            FROM emails
            GROUP BY conversation_id
            """
        ).fetchall()
    return {
        str(row["conversation_id"]): {
            "has_unread": bool(row["has_unread"]),
            "has_flagged": bool(row["has_flagged"]),
        }
        for row in rows
    }


def _filter_conversations(conversations: list[dict[str, Any]], filter_state: dict[str, Any] | None) -> list[dict[str, Any]]:
    if not filter_state or filter_state.get("kind") == "all":
        return conversations

    kind = str(filter_state.get("kind") or "")
    value = filter_state.get("value")
    if kind == "ownership":
        return [conversation for conversation in conversations if conversation.get("ownership") == value]
    if kind == "type":
        return [conversation for conversation in conversations if conversation.get("type") == value]
    if kind in {"has_unread", "has_flagged"}:
        return [conversation for conversation in conversations if bool(conversation.get(kind)) == bool(value)]
    return conversations


def _load_conversation_detail(conversation_id: str) -> dict[str, Any] | None:
    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        conversation = conn.execute(
            """
            SELECT conversation_id, topic, participants, latest_sender, last_activity
            FROM conversations
            WHERE conversation_id = ?
            """,
            (conversation_id,),
        ).fetchone()
        if conversation is None:
            return None

        emails = conn.execute(
            """
            SELECT entry_id, subject, sender_name, received_time, read_status, body_preview
            FROM emails
            WHERE conversation_id = ?
            ORDER BY received_time DESC, sent_time DESC, rowid DESC
            """,
            (conversation_id,),
        ).fetchall()

    email_rows = [dict(row) for row in emails]
    participants = [
        participant.strip()
        for participant in str(conversation["participants"] or "").split(",")
        if participant.strip()
    ]
    latest_email = email_rows[0] if email_rows else {}
    return {
        "conversation_id": str(conversation["conversation_id"]),
        "topic": str(conversation["topic"] or "Untitled Conversation"),
        "participants": participants,
        "latest_sender": str(conversation["latest_sender"] or ""),
        "last_activity": str(conversation["last_activity"] or ""),
        "emails": email_rows,
        "latest_email": latest_email,
    }


def _open_email_in_outlook(entry_id: str) -> tuple[bool, str]:
    try:
        import win32com.client  # type: ignore[import-not-found]

        namespace = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        namespace.GetItemFromID(entry_id).Display()
        return True, "Opened email in Outlook."
    except Exception as exc:  # pragma: no cover
        return False, f"Could not open Outlook email: {exc}"


def create_app(sync_limit: int | None = None) -> "Dash":
    """Create the Dash app with live Outlook data."""

    if Dash is None:
        raise RuntimeError("Dash is not installed. Install requirements.txt to run the dashboard UI.")

    # Sync data from Outlook (incremental if cache exists)
    report = _sync_outlook(limit_per_folder=sync_limit)

    app = Dash(__name__, suppress_callback_exceptions=True)
    app.layout = _build_layout(report)

    @app.callback(
        Output("conversations-table", "data"),
        Output("conversations-table", "active_cell"),
        Output("active-filter-label", "children"),
        Output("clear-filter-button", "disabled"),
        Input("metric-filter-store", "data"),
        Input("conversations-store", "data"),
    )
    def filter_conversations(
        filter_state: dict[str, Any] | None,
        conversations: list[dict[str, Any]] | None,
    ) -> tuple[list[dict[str, Any]], None, str, bool]:
        all_conversations = conversations or []
        applied_filter = filter_state or {"kind": "all", "value": None, "label": "All Conversations"}
        filtered = _filter_conversations(all_conversations, applied_filter)
        label = f"Active Filter: {applied_filter.get('label', 'All Conversations')} ({len(filtered)} conversations)"
        return filtered, None, label, applied_filter.get("kind") == "all"

    @app.callback(
        Output("metric-filter-store", "data"),
        Input(_metric_card_id("Total Emails Synced"), "n_clicks"),
        Input(_metric_card_id("Conversations"), "n_clicks"),
        Input(_metric_card_id("Folders Synced"), "n_clicks"),
        Input(_metric_card_id("Unread"), "n_clicks"),
        Input(_metric_card_id("Flagged"), "n_clicks"),
        Input(_metric_card_id("Awaiting My Action"), "n_clicks"),
        Input(_metric_card_id("Waiting For Others"), "n_clicks"),
        Input(_metric_card_id("Information Only"), "n_clicks"),
        Input(_metric_card_id("Unknown"), "n_clicks"),
        Input(_metric_card_id("Internal"), "n_clicks"),
        Input(_metric_card_id("External"), "n_clicks"),
        Input(_metric_card_id("Mixed"), "n_clicks"),
        Input("clear-filter-button", "n_clicks"),
        prevent_initial_call=True,
    )
    def update_metric_filter(*_: Any) -> dict[str, Any] | Any:
        triggered = callback_context.triggered[0]["prop_id"].split(".")[0] if callback_context.triggered else ""
        if not triggered:
            return no_update
        if triggered == "clear-filter-button":
            return {"kind": "all", "value": None, "label": "All Conversations"}
        return _metric_filters().get(triggered, no_update)

    @app.callback(
        Output("conversation-detail-panel", "children"),
        Output("selected-entry-store", "data"),
        Input("conversations-table", "active_cell"),
        Input("conversations-table", "derived_virtual_data"),
        State("conversations-table", "data"),
    )
    def show_conversation_detail(
        active_cell: dict[str, Any] | None,
        visible_rows: list[dict[str, Any]] | None,
        all_rows: list[dict[str, Any]] | None,
    ) -> tuple[Any, dict[str, Any] | None]:
        rows = visible_rows or all_rows or []
        if not active_cell or active_cell.get("row") is None or active_cell["row"] >= len(rows):
            return (
                html.Div(
                    "Select a conversation row to see thread details and open the latest email in Outlook.",
                    style={"padding": "16px", "color": "#666"},
                ),
                None,
            )

        conversation_id = str(rows[active_cell["row"]].get("conversation_id") or "")
        if not conversation_id:
            return no_update, no_update

        detail = _load_conversation_detail(conversation_id)
        if detail is None:
            return html.Div("Conversation details are unavailable.", style={"padding": "16px", "color": "#c62828"}), None

        email_rows = [
            {
                "subject": email.get("subject") or "(No subject)",
                "sender": email.get("sender_name") or "",
                "date": email.get("received_time") or "",
                "read_status": email.get("read_status") or "",
            }
            for email in detail["emails"]
        ]
        latest_email = detail["latest_email"]
        body_preview = str(latest_email.get("body_preview") or "No preview available.")
        selected_entry = str(latest_email.get("entry_id") or "")

        panel = html.Div(
            [
                html.H3("Conversation Detail", style={"marginTop": "0"}),
                html.Div([html.Strong("Topic: "), html.Span(detail["topic"])]),
                html.Div(
                    [
                        html.Strong("Participants: "),
                        html.Span(", ".join(detail["participants"]) if detail["participants"] else "No participants found"),
                    ],
                    style={"marginTop": "8px"},
                ),
                html.H4("Emails", style={"marginTop": "16px"}),
                dash_table.DataTable(
                    columns=[
                        {"name": "Subject", "id": "subject"},
                        {"name": "Sender", "id": "sender"},
                        {"name": "Date", "id": "date"},
                        {"name": "Read Status", "id": "read_status"},
                    ],
                    data=email_rows,
                    page_size=10,
                    style_table={"overflowX": "auto"},
                    style_cell={"textAlign": "left", "padding": "8px", "maxWidth": "300px", "whiteSpace": "normal"},
                    style_header={"fontWeight": "bold", "backgroundColor": "#f5f5f5"},
                ),
                html.H4("Latest Email Preview", style={"marginTop": "16px"}),
                html.Div(
                    body_preview,
                    style={
                        "backgroundColor": "#fafafa",
                        "border": "1px solid #e0e0e0",
                        "borderRadius": "8px",
                        "padding": "12px",
                        "whiteSpace": "pre-wrap",
                    },
                ),
                html.Button(
                    "Open Latest Email in Outlook",
                    id="open-email-button",
                    n_clicks=0,
                    disabled=not selected_entry,
                    style={
                        "marginTop": "16px",
                        "padding": "10px 16px",
                        "border": "none",
                        "borderRadius": "8px",
                        "backgroundColor": "#1976d2",
                        "color": "white",
                        "cursor": "pointer",
                        "fontWeight": "bold",
                    },
                ),
                html.Div(id="outlook-open-status", style={"marginTop": "12px"}),
            ],
            style={
                "marginTop": "20px",
                "padding": "20px",
                "border": "1px solid #e0e0e0",
                "borderRadius": "10px",
                "backgroundColor": "white",
                "boxShadow": "0 2px 6px rgba(0, 0, 0, 0.06)",
            },
        )
        return panel, {"entry_id": selected_entry}

    @app.callback(
        Output("outlook-open-status", "children"),
        Input("open-email-button", "n_clicks"),
        State("selected-entry-store", "data"),
        prevent_initial_call=True,
    )
    def open_selected_email(n_clicks: int, selected_entry: dict[str, Any] | None) -> Any:
        if not n_clicks:
            return no_update
        entry_id = str((selected_entry or {}).get("entry_id") or "")
        if not entry_id:
            return html.Span("No email selected to open.", style={"color": "#c62828"})
        opened, message = _open_email_in_outlook(entry_id)
        return html.Span(message, style={"color": "#2e7d32" if opened else "#c62828"})

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

    conversation_flags = _load_conversation_flags()

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

    # Conversations table with interactive filters
    conv_data = []
    for conv in report["conversations"]:
        flags = conversation_flags.get(conv.conversation_id, {})
        conv_data.append({
            "topic": conv.topic,
            "ownership": conv.ownership,
            "type": conv.conversation_type,
            "messages": conv.message_count,
            "latest_sender": conv.latest_sender,
            "last_activity": conv.last_activity.strftime("%Y-%m-%d %H:%M") if conv.last_activity else "",
            "action": "⚠️ Yes" if conv.action_required else "",
            "conversation_id": conv.conversation_id,
            "participants": ", ".join(conv.participants),
            "has_unread": flags.get("has_unread", False),
            "has_flagged": flags.get("has_flagged", False),
        })

    conversations_section = html.Div(
        [
            html.H2("💬 Conversations"),
            html.Div(
                [
                    html.Span(
                        "Active Filter: All Conversations",
                        id="active-filter-label",
                        style={"fontWeight": "bold", "color": "#444"},
                    ),
                    html.Button(
                        "Clear Filter",
                        id="clear-filter-button",
                        n_clicks=0,
                        disabled=True,
                        style={
                            "padding": "10px 16px",
                            "border": "1px solid #d0d7de",
                            "borderRadius": "8px",
                            "backgroundColor": "white",
                            "cursor": "pointer",
                        },
                    ),
                ],
                style={
                    "display": "flex",
                    "justifyContent": "space-between",
                    "alignItems": "center",
                    "gap": "12px",
                    "marginBottom": "12px",
                },
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
                    {"name": "Conversation ID", "id": "conversation_id"},
                    {"name": "Participants", "id": "participants"},
                    {"name": "Unread", "id": "has_unread"},
                    {"name": "Flagged", "id": "has_flagged"},
                ],
                data=conv_data,
                hidden_columns=["conversation_id", "participants", "has_unread", "has_flagged"],
                page_size=20,
                sort_action="native",
                filter_action="native",
                cell_selectable=True,
                style_table={"overflowX": "auto"},
                style_cell={"textAlign": "left", "padding": "8px", "maxWidth": "300px", "overflow": "hidden", "textOverflow": "ellipsis"},
                style_header={"fontWeight": "bold", "backgroundColor": "#f5f5f5"},
                style_data_conditional=[
                    {
                        "if": {"state": "active"},
                        "backgroundColor": "#e3f2fd",
                        "border": "1px solid #1976d2",
                    }
                ],
                css=[{"selector": "tr:hover", "rule": "background-color: #f7fbff; cursor: pointer;"}],
            ),
            dcc.Store(id="conversations-store", data=conv_data),
            dcc.Store(id="metric-filter-store", data={"kind": "all", "value": None, "label": "All Conversations"}),
            dcc.Store(id="selected-entry-store"),
            html.Div(
                id="conversation-detail-panel",
                children=html.Div(
                    "Select a conversation row to see thread details and open the latest email in Outlook.",
                    style={"padding": "16px", "color": "#666"},
                ),
            ),
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
    return html.Button(
        html.Div(
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
                "transition": "transform 0.15s ease, box-shadow 0.15s ease, border-color 0.15s ease",
                "boxShadow": "0 2px 4px rgba(0, 0, 0, 0.04)",
            },
        ),
        id=_metric_card_id(title),
        className="metric-card-button",
        n_clicks=0,
        title=f"Filter conversations by {title}",
        style={"padding": "0", "background": "transparent"},
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
