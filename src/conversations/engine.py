"""Conversation grouping, classification, and ownership heuristics."""

from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime
from typing import Iterable, Mapping

ACTION_PHRASES = [
    "can you",
    "could you",
    "please",
    "need",
    "required",
    "review",
    "approve",
    "confirm",
    "provide",
    "action required",
    "response needed",
    "deadline",
    "urgent",
    "asap",
    "blocking",
]

INFO_ONLY_PHRASES = ["fyi", "notification", "status update", "automated", "distribution"]


@dataclass(slots=True)
class ConversationSummary:
    """Aggregated view of a conversation thread."""

    conversation_id: str
    topic: str
    participants: list[str] = field(default_factory=list)
    message_count: int = 0
    start_date: datetime | None = None
    last_activity: datetime | None = None
    latest_sender: str = ""
    folder: str = ""
    conversation_type: str = "Unknown"
    ownership: str = "Unknown"
    action_required: bool = False
    action_confidence: int = 0
    action_summary: str = ""

    def to_dict(self) -> dict[str, object]:
        """Return a JSON-serializable summary."""

        return {
            "conversation_id": self.conversation_id,
            "topic": self.topic,
            "participants": self.participants,
            "message_count": self.message_count,
            "start_date": self.start_date.isoformat() if self.start_date else None,
            "last_activity": self.last_activity.isoformat() if self.last_activity else None,
            "latest_sender": self.latest_sender,
            "folder": self.folder,
            "conversation_type": self.conversation_type,
            "ownership": self.ownership,
            "action_required": self.action_required,
            "action_confidence": self.action_confidence,
            "action_summary": self.action_summary,
        }


class ConversationEngine:
    """Build conversation-level intelligence from normalized email records."""

    def __init__(self, internal_domain: str = "@motionapplied.com") -> None:
        self.internal_domain = internal_domain.lower().lstrip("@")

    def build_conversations(self, emails: Iterable[Mapping[str, object]]) -> list[ConversationSummary]:
        """Group emails by ConversationID and derive summary fields."""

        grouped: dict[str, list[Mapping[str, object]]] = defaultdict(list)
        for email in emails:
            conversation_id = str(
                email.get("conversation_id")
                or email.get("ConversationID")
                or email.get("entry_id")
                or email.get("EntryID")
                or ""
            )
            if conversation_id:
                grouped[conversation_id].append(email)

        return [self._summarize(conversation_id, messages) for conversation_id, messages in grouped.items()]

    def detect_action_request(self, text: str) -> tuple[bool, int, str]:
        """Detect action-oriented phrases in a subject/body snippet."""

        lowered = text.lower()
        matches = [phrase for phrase in ACTION_PHRASES if phrase in lowered]
        if not matches:
            return False, 0, ""
        confidence = min(100, 35 + len(matches) * 12)
        summary = ", ".join(dict.fromkeys(match.title() for match in matches[:3]))
        return True, confidence, summary

    def determine_ownership(self, messages: list[Mapping[str, object]]) -> str:
        """Infer whether the next action is with me, others, or nobody."""

        latest = self._latest_message(messages)
        latest_folder = str(latest.get("folder") or latest.get("Folder") or "").lower()
        text = " ".join(
            [
                str(latest.get("subject") or latest.get("Subject") or ""),
                str(latest.get("body_preview") or latest.get("BodyPreview") or ""),
            ]
        ).lower()
        action_required, _, _ = self.detect_action_request(text)
        if any(marker in text for marker in INFO_ONLY_PHRASES) and not action_required:
            return "Information Only"
        if "sent" in latest_folder:
            return "Waiting For Others"
        if action_required:
            return "Awaiting My Action"
        return "Unknown"

    def classify_conversation_type(self, messages: list[Mapping[str, object]]) -> str:
        """Classify a conversation as Internal, External, or Mixed."""

        domains = {
            self._domain_from_email(str(message.get("sender_email") or message.get("SenderEmail") or ""))
            for message in messages
        }
        domains = {domain for domain in domains if domain}
        if not domains:
            return "Unknown"
        internal = {domain for domain in domains if domain == self.internal_domain}
        if domains == internal:
            return "Internal"
        if internal and len(domains) > len(internal):
            return "Mixed"
        return "External"

    def _summarize(self, conversation_id: str, messages: list[Mapping[str, object]]) -> ConversationSummary:
        sorted_messages = sorted(messages, key=self._message_timestamp)
        latest = sorted_messages[-1]
        text = " ".join(
            [
                str(latest.get("subject") or latest.get("Subject") or ""),
                str(latest.get("body_preview") or latest.get("BodyPreview") or ""),
            ]
        )
        action_required, confidence, summary = self.detect_action_request(text)
        participants = sorted(
            {
                address
                for message in sorted_messages
                for address in self._participant_addresses(message)
            }
        )
        return ConversationSummary(
            conversation_id=conversation_id,
            topic=str(
                latest.get("conversation_topic")
                or latest.get("ConversationTopic")
                or latest.get("subject")
                or latest.get("Subject")
                or "Untitled Conversation"
            ),
            participants=participants,
            message_count=len(sorted_messages),
            start_date=self._message_timestamp(sorted_messages[0]),
            last_activity=self._message_timestamp(latest),
            latest_sender=str(
                latest.get("sender_email")
                or latest.get("SenderEmail")
                or latest.get("sender_name")
                or latest.get("SenderName")
                or ""
            ),
            folder=str(latest.get("folder") or latest.get("Folder") or ""),
            conversation_type=self.classify_conversation_type(sorted_messages),
            ownership=self.determine_ownership(sorted_messages),
            action_required=action_required,
            action_confidence=confidence,
            action_summary=summary,
        )

    @staticmethod
    def _participant_addresses(message: Mapping[str, object]) -> list[str]:
        fields = [
            str(message.get("sender_email") or message.get("SenderEmail") or ""),
            str(message.get("recipients") or message.get("Recipients") or ""),
            str(message.get("cc") or message.get("CC") or ""),
        ]
        participants: list[str] = []
        for field in fields:
            participants.extend([value.strip() for value in field.split(";") if value.strip()])
        return participants

    @staticmethod
    def _latest_message(messages: list[Mapping[str, object]]) -> Mapping[str, object]:
        return max(messages, key=ConversationEngine._message_timestamp)

    @staticmethod
    def _message_timestamp(message: Mapping[str, object]) -> datetime:
        candidates = [
            message.get("received_time") or message.get("ReceivedTime"),
            message.get("sent_time") or message.get("SentTime"),
            message.get("last_activity") or message.get("LastActivity"),
        ]
        for value in candidates:
            if isinstance(value, datetime):
                # Strip timezone to avoid naive vs aware comparison errors
                return value.replace(tzinfo=None)
            if isinstance(value, str) and value:
                try:
                    parsed = datetime.fromisoformat(value)
                    return parsed.replace(tzinfo=None)
                except ValueError:
                    continue
        return datetime.min

    @staticmethod
    def _domain_from_email(email_address: str) -> str:
        if "@" not in email_address:
            return ""
        return email_address.rsplit("@", 1)[-1].lower()
