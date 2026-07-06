"""Priority and urgency management for conversation work items."""

from __future__ import annotations

from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from typing import Mapping
from uuid import uuid4

PRIORITY_LEVELS = ["P1 Critical", "P2 High", "P3 Normal", "P4 Low", "P5 Someday"]
PRIORITY_ORDER = {value: index for index, value in enumerate(PRIORITY_LEVELS, start=1)}


@dataclass(slots=True)
class WorkItem:
    """Common work item representation for emails, manual tasks, and meeting actions."""

    id: str
    title: str
    source: str
    status: str = "Open"
    priority: str = "P3 Normal"
    urgency_score: int = 0
    owner: str = "Unknown"
    created_date: datetime | None = None
    last_activity_date: datetime | None = None
    due_date: datetime | None = None
    conversation_id: str = ""
    notes: str = ""
    tags: list[str] = field(default_factory=list)
    folder: str = ""

    @classmethod
    def create(cls, title: str, source: str, **kwargs: object) -> "WorkItem":
        """Create a work item with an auto-generated identifier."""

        return cls(id=str(uuid4()), title=title, source=source, **kwargs)

    def to_record(self) -> dict[str, object]:
        """Serialize the work item for persistence or dashboard rendering."""

        data = asdict(self)
        data["tags"] = ", ".join(self.tags)
        for key in ("created_date", "last_activity_date", "due_date"):
            value = data[key]
            data[key] = value.isoformat() if isinstance(value, datetime) else None
        return data


class BacklogManager:
    """Apply rule-based urgency scoring and stable backlog ordering."""

    def calculate_urgency_score(self, item: Mapping[str, object]) -> int:
        """Calculate a 0-100 urgency score from common work-item signals."""

        score = 0
        status = str(item.get("status", "")).lower()
        owner = str(item.get("owner", "")).lower()
        importance = str(item.get("importance", item.get("priority", ""))).lower()
        source = str(item.get("source", "")).lower()
        notes = str(item.get("notes", "")).lower()
        folder = str(item.get("folder", "")).lower()

        if "awaiting my action" in owner:
            score += 25
        if "waiting for others" in owner:
            score -= 10
        if "high" in importance or "critical" in importance:
            score += 20
        if "unread" in notes or item.get("read_status") == "Unread":
            score += 10
        if str(item.get("flag_status", "0")) not in {"", "0", "None"}:
            score += 15
        if source == "email" and folder and "sent" not in folder:
            score += 5
        if any(marker in notes for marker in ["deadline", "urgent", "asap", "blocking"]):
            score += 20
        if status in {"completed", "done", "closed"}:
            score -= 40
        if owner == "information only":
            score -= 20

        due_date = self._coerce_datetime(item.get("due_date") or item.get("DueDate"))
        if due_date:
            days_until_due = (due_date.date() - date.today()).days
            if days_until_due < 0:
                score += 25
            elif days_until_due <= 2:
                score += 15
            elif days_until_due <= 7:
                score += 5

        last_activity = self._coerce_datetime(item.get("last_activity_date") or item.get("LastActivityDate"))
        if last_activity:
            age_days = (datetime.now() - last_activity).days
            if age_days <= 1:
                score += 10
            elif age_days > 30:
                score -= 10

        return max(0, min(100, score))

    def assign_priority(self, urgency_score: int, user_priority: str | None = None) -> str:
        """Return the user override or a system-derived priority level."""

        if user_priority in PRIORITY_ORDER:
            return user_priority
        if urgency_score >= 85:
            return "P1 Critical"
        if urgency_score >= 70:
            return "P2 High"
        if urgency_score >= 45:
            return "P3 Normal"
        if urgency_score >= 20:
            return "P4 Low"
        return "P5 Someday"

    def rank_items(self, items: list[WorkItem], user_order: Mapping[str, int] | None = None) -> list[WorkItem]:
        """Sort work items by manual order, then priority, urgency, and activity."""

        order = dict(user_order or {})
        enriched: list[WorkItem] = []
        for item in items:
            if item.urgency_score == 0:
                item.urgency_score = self.calculate_urgency_score(item.to_record())
            if item.priority not in PRIORITY_ORDER:
                item.priority = self.assign_priority(item.urgency_score)
            enriched.append(item)

        return sorted(
            enriched,
            key=lambda item: (
                order.get(item.id, 10**6),
                PRIORITY_ORDER.get(item.priority, 99),
                -item.urgency_score,
                -(item.last_activity_date or item.created_date or datetime.min).timestamp(),
            ),
        )

    @staticmethod
    def _coerce_datetime(value: object) -> datetime | None:
        if isinstance(value, datetime):
            return value
        if isinstance(value, str) and value:
            try:
                return datetime.fromisoformat(value)
            except ValueError:
                return None
        return None
