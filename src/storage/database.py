"""SQLite storage layer for dashboard entities and sync state."""

from __future__ import annotations

import sqlite3
from contextlib import contextmanager
from pathlib import Path
from typing import Iterator, Mapping

SCHEMA = """
CREATE TABLE IF NOT EXISTS emails (
    entry_id TEXT PRIMARY KEY,
    conversation_id TEXT NOT NULL,
    conversation_topic TEXT,
    subject TEXT,
    sender_name TEXT,
    sender_email TEXT,
    sender_domain TEXT,
    recipients TEXT,
    cc TEXT,
    received_time TEXT,
    sent_time TEXT,
    body_preview TEXT,
    folder TEXT,
    category TEXT,
    flag_status TEXT,
    read_status TEXT,
    importance TEXT,
    synced_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_emails_conversation_id ON emails (conversation_id);
CREATE INDEX IF NOT EXISTS idx_emails_received_time ON emails (received_time);
CREATE INDEX IF NOT EXISTS idx_emails_sender_domain ON emails (sender_domain);

CREATE TABLE IF NOT EXISTS conversations (
    conversation_id TEXT PRIMARY KEY,
    topic TEXT,
    participants TEXT,
    message_count INTEGER DEFAULT 0,
    start_date TEXT,
    last_activity TEXT,
    latest_sender TEXT,
    folder TEXT,
    conversation_type TEXT,
    ownership TEXT,
    action_required INTEGER DEFAULT 0,
    action_confidence INTEGER DEFAULT 0,
    action_summary TEXT,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS work_items (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    source TEXT NOT NULL,
    status TEXT,
    priority TEXT,
    urgency_score INTEGER DEFAULT 0,
    owner TEXT,
    created_date TEXT,
    last_activity_date TEXT,
    due_date TEXT,
    conversation_id TEXT,
    notes TEXT,
    tags TEXT,
    folder TEXT,
    user_priority TEXT,
    final_priority TEXT,
    rank_order INTEGER,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (conversation_id) REFERENCES conversations (conversation_id)
);

CREATE INDEX IF NOT EXISTS idx_work_items_priority ON work_items (final_priority, urgency_score DESC);
CREATE INDEX IF NOT EXISTS idx_work_items_owner ON work_items (owner);
CREATE INDEX IF NOT EXISTS idx_work_items_status ON work_items (status);

CREATE TABLE IF NOT EXISTS sync_state (
    folder_name TEXT PRIMARY KEY,
    last_synced_at TEXT,
    last_entry_id TEXT,
    sync_mode TEXT DEFAULT 'incremental',
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS dashboard_views (
    name TEXT PRIMARY KEY,
    filters_json TEXT,
    grouping TEXT,
    columns_json TEXT,
    sorting_json TEXT,
    updated_at TEXT DEFAULT CURRENT_TIMESTAMP
);
"""


class DatabaseManager:
    """Thin SQLite wrapper for schema management and common CRUD operations."""

    def __init__(self, database_path: str = "database/dashboard.db") -> None:
        self.database_path = Path(database_path)
        self._shared_connection: sqlite3.Connection | None = None

    @contextmanager
    def connect(self) -> Iterator[sqlite3.Connection]:
        """Open a SQLite connection, ensuring the parent directory exists."""

        if str(self.database_path) == ":memory:":
            if self._shared_connection is None:
                self._shared_connection = sqlite3.connect(":memory:")
                self._shared_connection.row_factory = sqlite3.Row
            connection = self._shared_connection
            should_close = False
        else:
            self.database_path.parent.mkdir(parents=True, exist_ok=True)
            connection = sqlite3.connect(self.database_path)
            connection.row_factory = sqlite3.Row
            should_close = True
        try:
            yield connection
            connection.commit()
        finally:
            if should_close:
                connection.close()

    def initialize(self) -> None:
        """Create all required dashboard tables and indexes."""

        with self.connect() as connection:
            connection.executescript(SCHEMA)

    def upsert_email(self, record: Mapping[str, object]) -> None:
        """Insert or update a normalized email record."""

        payload = self._normalize(record)
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO emails (
                    entry_id, conversation_id, conversation_topic, subject, sender_name,
                    sender_email, sender_domain, recipients, cc, received_time, sent_time,
                    body_preview, folder, category, flag_status, read_status, importance
                ) VALUES (
                    :entry_id, :conversation_id, :conversation_topic, :subject, :sender_name,
                    :sender_email, :sender_domain, :recipients, :cc, :received_time, :sent_time,
                    :body_preview, :folder, :category, :flag_status, :read_status, :importance
                )
                ON CONFLICT(entry_id) DO UPDATE SET
                    conversation_id = excluded.conversation_id,
                    conversation_topic = excluded.conversation_topic,
                    subject = excluded.subject,
                    sender_name = excluded.sender_name,
                    sender_email = excluded.sender_email,
                    sender_domain = excluded.sender_domain,
                    recipients = excluded.recipients,
                    cc = excluded.cc,
                    received_time = excluded.received_time,
                    sent_time = excluded.sent_time,
                    body_preview = excluded.body_preview,
                    folder = excluded.folder,
                    category = excluded.category,
                    flag_status = excluded.flag_status,
                    read_status = excluded.read_status,
                    importance = excluded.importance,
                    synced_at = CURRENT_TIMESTAMP
                """,
                payload,
            )

    def upsert_conversation(self, record: Mapping[str, object]) -> None:
        """Insert or update a conversation summary record."""

        payload = self._normalize(record)
        payload["participants"] = self._stringify(payload.get("participants", ""))
        payload["action_required"] = int(bool(payload.get("action_required", False)))
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO conversations (
                    conversation_id, topic, participants, message_count, start_date, last_activity,
                    latest_sender, folder, conversation_type, ownership, action_required,
                    action_confidence, action_summary
                ) VALUES (
                    :conversation_id, :topic, :participants, :message_count, :start_date, :last_activity,
                    :latest_sender, :folder, :conversation_type, :ownership, :action_required,
                    :action_confidence, :action_summary
                )
                ON CONFLICT(conversation_id) DO UPDATE SET
                    topic = excluded.topic,
                    participants = excluded.participants,
                    message_count = excluded.message_count,
                    start_date = excluded.start_date,
                    last_activity = excluded.last_activity,
                    latest_sender = excluded.latest_sender,
                    folder = excluded.folder,
                    conversation_type = excluded.conversation_type,
                    ownership = excluded.ownership,
                    action_required = excluded.action_required,
                    action_confidence = excluded.action_confidence,
                    action_summary = excluded.action_summary,
                    updated_at = CURRENT_TIMESTAMP
                """,
                payload,
            )

    def upsert_work_item(self, record: Mapping[str, object]) -> None:
        """Insert or update a work item using the shared dashboard schema."""

        payload = self._normalize(record)
        payload["tags"] = self._stringify(payload.get("tags", ""))
        payload.setdefault("user_priority", payload.get("priority"))
        payload.setdefault("final_priority", payload.get("priority"))
        payload.setdefault("rank_order", None)
        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO work_items (
                    id, title, source, status, priority, urgency_score, owner, created_date,
                    last_activity_date, due_date, conversation_id, notes, tags, folder,
                    user_priority, final_priority, rank_order
                ) VALUES (
                    :id, :title, :source, :status, :priority, :urgency_score, :owner, :created_date,
                    :last_activity_date, :due_date, :conversation_id, :notes, :tags, :folder,
                    :user_priority, :final_priority, :rank_order
                )
                ON CONFLICT(id) DO UPDATE SET
                    title = excluded.title,
                    source = excluded.source,
                    status = excluded.status,
                    priority = excluded.priority,
                    urgency_score = excluded.urgency_score,
                    owner = excluded.owner,
                    created_date = excluded.created_date,
                    last_activity_date = excluded.last_activity_date,
                    due_date = excluded.due_date,
                    conversation_id = excluded.conversation_id,
                    notes = excluded.notes,
                    tags = excluded.tags,
                    folder = excluded.folder,
                    user_priority = excluded.user_priority,
                    final_priority = excluded.final_priority,
                    rank_order = excluded.rank_order,
                    updated_at = CURRENT_TIMESTAMP
                """,
                payload,
            )

    def list_work_items(self, status: str | None = None, limit: int = 100) -> list[dict[str, object]]:
        """Return persisted work items, optionally filtered by workflow status."""

        query = "SELECT * FROM work_items"
        params: tuple[object, ...] = ()
        if status:
            query += " WHERE status = ?"
            params = (status,)
        query += " ORDER BY (rank_order IS NULL), rank_order ASC, urgency_score DESC, updated_at DESC LIMIT ?"
        params += (limit,)
        with self.connect() as connection:
            rows = connection.execute(query, params).fetchall()
        return [dict(row) for row in rows]

    def update_sync_state(self, folder_name: str, last_synced_at: str, last_entry_id: str = "") -> None:
        """Persist incremental sync markers for an Outlook folder."""

        with self.connect() as connection:
            connection.execute(
                """
                INSERT INTO sync_state (folder_name, last_synced_at, last_entry_id)
                VALUES (?, ?, ?)
                ON CONFLICT(folder_name) DO UPDATE SET
                    last_synced_at = excluded.last_synced_at,
                    last_entry_id = excluded.last_entry_id,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (folder_name, last_synced_at, last_entry_id),
            )

    @staticmethod
    def _normalize(record: Mapping[str, object]) -> dict[str, object]:
        payload = dict(record)
        for key, value in list(payload.items()):
            if hasattr(value, "isoformat"):
                payload[key] = value.isoformat()
        return payload

    @staticmethod
    def _stringify(value: object) -> str:
        if isinstance(value, (list, tuple, set)):
            return ", ".join(str(item) for item in value)
        return "" if value is None else str(value)
