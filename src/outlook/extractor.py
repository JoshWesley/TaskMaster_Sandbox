"""Outlook COM extraction helpers."""

from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime
from email.utils import getaddresses
from typing import Iterable, Iterator

try:
    import win32com.client  # type: ignore[import-not-found]
except ImportError:  # pragma: no cover - optional dependency for non-Windows CI.
    win32com = None
else:  # pragma: no cover - exercised only when pywin32 is installed.
    win32com = win32com.client


@dataclass(slots=True)
class EmailRecord:
    """Normalized email data captured from Outlook."""

    entry_id: str
    conversation_id: str
    conversation_topic: str
    subject: str
    sender_name: str
    sender_email: str
    sender_domain: str
    recipients: str
    cc: str
    received_time: datetime | None
    sent_time: datetime | None
    body_preview: str
    folder: str
    category: str
    flag_status: str
    read_status: str
    importance: str

    def to_dict(self) -> dict[str, object]:
        """Return a storage-friendly dictionary representation."""

        return asdict(self)


class OutlookExtractor:
    """Read Outlook folders and normalize mail items for downstream processing."""

    def __init__(
        self,
        internal_domain: str = "@motionapplied.com",
        folder_names: Iterable[str] | None = None,
        namespace: object | None = None,
    ) -> None:
        self.internal_domain = internal_domain.lower()
        self.folder_names = list(folder_names or ["Inbox", "Sent Items"])
        self._namespace = namespace

    def connect(self) -> object:
        """Connect to Outlook MAPI and return the namespace object."""

        if self._namespace is not None:
            return self._namespace
        if win32com is None:
            raise RuntimeError(
                "pywin32 is required to use Outlook COM. Install requirements.txt first."
            )
        application = win32com.Dispatch("Outlook.Application")
        self._namespace = application.GetNamespace("MAPI")
        return self._namespace

    def extract_folders(
        self, folder_names: Iterable[str] | None = None, limit: int | None = None
    ) -> list[EmailRecord]:
        """Return normalized messages from the requested Outlook folders."""

        records: list[EmailRecord] = []
        for folder_name in folder_names or self.folder_names:
            for message in self.iter_folder_items(folder_name, limit=limit):
                records.append(self.extract_message(message, folder_name))
        return records

    def iter_folder_items(self, folder_name: str, limit: int | None = None) -> Iterator[object]:
        """Yield raw Outlook items from a folder in descending received order."""

        namespace = self.connect()
        folder = namespace.Folders.Item(1).Folders[folder_name]
        items = folder.Items
        items.Sort("[ReceivedTime]", True)
        count = 0
        for item in items:
            if getattr(item, "Class", None) != 43:
                continue
            yield item
            count += 1
            if limit is not None and count >= limit:
                break

    def extract_message(self, item: object, folder_name: str) -> EmailRecord:
        """Normalize a single Outlook MailItem into an EmailRecord."""

        sender_email = self._safe_get(item, "SenderEmailAddress")
        unread = self._safe_raw(item, "UnRead", False)
        return EmailRecord(
            entry_id=self._safe_get(item, "EntryID"),
            conversation_id=self._safe_get(item, "ConversationID") or self._safe_get(item, "EntryID"),
            conversation_topic=self._safe_get(item, "ConversationTopic") or self._safe_get(item, "Subject"),
            subject=self._safe_get(item, "Subject"),
            sender_name=self._safe_get(item, "SenderName"),
            sender_email=sender_email,
            sender_domain=self._extract_domain(sender_email),
            recipients=self._normalize_recipients(self._safe_get(item, "To")),
            cc=self._normalize_recipients(self._safe_get(item, "CC")),
            received_time=self._safe_datetime(item, "ReceivedTime"),
            sent_time=self._safe_datetime(item, "SentOn"),
            body_preview=self._body_preview(self._safe_get(item, "Body")),
            folder=folder_name,
            category=self._safe_get(item, "Categories"),
            flag_status=str(self._safe_raw(item, "FlagStatus", "")),
            read_status="Unread" if bool(unread) else "Read",
            importance=self._importance_name(str(self._safe_raw(item, "Importance", ""))),
        )

    @staticmethod
    def _safe_get(item: object, attribute: str, default: str = "") -> str:
        value = getattr(item, attribute, default)
        return "" if value is None else str(value)

    @staticmethod
    def _safe_raw(item: object, attribute: str, default: object = "") -> object:
        return getattr(item, attribute, default)

    @staticmethod
    def _safe_datetime(item: object, attribute: str) -> datetime | None:
        value = getattr(item, attribute, None)
        return value if isinstance(value, datetime) else None

    @staticmethod
    def _normalize_recipients(value: str) -> str:
        recipients = [address for _, address in getaddresses([value or ""]) if address]
        return "; ".join(recipients)

    @staticmethod
    def _extract_domain(email_address: str) -> str:
        if "@" not in email_address:
            return ""
        return email_address.rsplit("@", 1)[-1].lower()

    @staticmethod
    def _body_preview(body: str, limit: int = 280) -> str:
        compact = " ".join(body.split())
        return compact[:limit]

    @staticmethod
    def _importance_name(level: str) -> str:
        mapping = {"0": "Low", "1": "Normal", "2": "High"}
        return mapping.get(str(level), str(level) or "Unknown")
