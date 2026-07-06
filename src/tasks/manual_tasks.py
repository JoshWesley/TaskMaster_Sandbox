"""Manual task and meeting action helpers."""

from __future__ import annotations

from datetime import datetime
from typing import Iterable

from src.backlog.manager import WorkItem


class ManualTaskManager:
    """Create common work-item records for non-email actions."""

    def create_manual_task(
        self,
        title: str,
        description: str = "",
        priority: str = "P3 Normal",
        due_date: datetime | None = None,
        tags: Iterable[str] | None = None,
        category: str = "Manual Task",
        status: str = "Open",
    ) -> WorkItem:
        """Create a user-authored manual task."""

        created_at = datetime.now()
        return WorkItem.create(
            title=title,
            source="Manual Task",
            status=status,
            priority=priority,
            due_date=due_date,
            notes=description,
            tags=self._normalize_tags(tags, category),
            created_date=created_at,
            last_activity_date=created_at,
            owner="Awaiting My Action",
            folder=category,
        )

    def create_meeting_action(
        self,
        action: str,
        owner: str,
        due_date: datetime | None = None,
        notes: str = "",
        tags: Iterable[str] | None = None,
        status: str = "Open",
    ) -> WorkItem:
        """Create a meeting-derived work item using the common model."""

        created_at = datetime.now()
        return WorkItem.create(
            title=action,
            source="Meeting Action",
            status=status,
            priority="P2 High" if due_date else "P3 Normal",
            due_date=due_date,
            notes=notes,
            tags=self._normalize_tags(tags, "Meeting Action"),
            created_date=created_at,
            last_activity_date=created_at,
            owner=owner or "Awaiting My Action",
            folder="Meeting Actions",
        )

    @staticmethod
    def _normalize_tags(tags: Iterable[str] | None, category: str) -> list[str]:
        normalized = {category}
        if tags:
            normalized.update(tag.strip() for tag in tags if tag and tag.strip())
        return sorted(normalized)
