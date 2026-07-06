"""Dash application for the Outlook Intelligence Dashboard.

Connects to Outlook via COM on startup, syncs emails into SQLite,
and displays live data about what has been pulled in.
"""

from __future__ import annotations

import traceback
from collections import Counter
from datetime import datetime
from typing import Any

try:
    from dash import Dash, Input, Output, dash_table, dcc, html
except ImportError:  # pragma: no cover
    Dash = None  # type: ignore[assignment]
    Input = Output = dash_table = dcc = html = None  # type: ignore[assignment]

import sys
from pathlib import Path

# Ensure src is importable when running this file directly.
_SRC_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(_SRC_ROOT))

from src.conversations.engine import ConversationEngine
from src.outlook.extractor import OutlookExtractor
from src.storage.database import DatabaseManager


def _sync_outlook(limit_per_folder: int = 500) -> dict[str, Any]:
    """Connect to Outlook, extract emails, build conversations, and return a status report."""

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
    }

    try:
        extractor = OutlookExtractor()
        extractor.connect()
        report["connected"] = True

        # Extract from default folders (Inbox + Sent Items)
        records = extractor.extract_folders(limit=limit_per_folder)
        report["total_emails"] = len(records)
        report["sync_time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Persist to SQLite
        db = DatabaseManager()
        db.initialize()
        for record in records:
            db.upsert_email(record.to_dict())

        # Folder breakdown
        for record in records:
            folder = record.folder
            report["emails_by_folder"][folder] = report["emails_by_folder"].get(folder, 0) + 1
            if folder not in report["folders_synced"]:
                report["folders_synced"].append(folder)
            report["top_senders"][record.sender_name] += 1
            report["top_domains"][record.sender_domain] += 1
            if record.read_status == "Unread":
                report["unread_count"] += 1
            if record.flag_status and record.flag_status not in ("0", "", "None"):
                report["flagged_count"] += 1

        # Build conversations
        engine = ConversationEngine()
        email_dicts = [r.to_dict() for r in records]
        conversations = engine.build_conversations(email_dicts)
        report["total_conversations"] = len(conversations)
        report["conversations"] = conversations

        for conv in conversations:
            report["ownership_counts"][conv.ownership] += 1
            report["type_counts"][conv.conversation_type] += 1

            # Persist conversation
            db.upsert_conversation(conv.to_dict())

    except Exception as e:
        report["error"] = f"{type(e).__name__}: {e}\n{traceback.format_exc()}"

    return report


def create_app(sync_limit: int = 500) -> "Dash":
    """Create the Dash app with live Outlook data."""

    if Dash is None:
        raise RuntimeError("Dash is not installed. Install requirements.txt to run the dashboard UI.")

    # Sync data from Outlook
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
        status_banner = html.Div(
            [
                html.Span("✅ Connected to Outlook", style={"fontWeight": "bold", "color": "#2e7d32"}),
                html.Span(f"  •  Last sync: {report['sync_time']}", style={"marginLeft": "16px", "color": "#555"}),
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
    print("Connecting to Outlook and syncing data...")
    app = create_app(sync_limit=500)
    print("Dashboard ready. Opening at http://127.0.0.1:8050")
    app.run(debug=True)
