"""Metrics and aggregations for the dashboard."""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime
from statistics import mean
from typing import Iterable, Mapping


class AnalyticsEngine:
    """Compute dashboard-ready metrics from emails and work items."""

    def email_load(self, emails: Iterable[Mapping[str, object]]) -> dict[str, dict[str, int]]:
        """Aggregate email volume by day, ISO week, and month."""

        per_day: Counter[str] = Counter()
        per_week: Counter[str] = Counter()
        per_month: Counter[str] = Counter()
        for email in emails:
            timestamp = self._timestamp(email)
            if timestamp is None:
                continue
            per_day[timestamp.strftime("%Y-%m-%d")] += 1
            iso_year, iso_week, _ = timestamp.isocalendar()
            per_week[f"{iso_year}-W{iso_week:02d}"] += 1
            per_month[timestamp.strftime("%Y-%m")] += 1
        return {
            "per_day": dict(per_day),
            "per_week": dict(per_week),
            "per_month": dict(per_month),
        }

    def workload_analysis(self, work_items: Iterable[Mapping[str, object]]) -> dict[str, dict[str, int]]:
        """Break work down by owner, folder, source, and topic."""

        by_owner: Counter[str] = Counter()
        by_folder: Counter[str] = Counter()
        by_source: Counter[str] = Counter()
        by_topic: Counter[str] = Counter()
        for item in work_items:
            by_owner[str(item.get("owner", "Unknown"))] += 1
            by_folder[str(item.get("folder", "Unassigned"))] += 1
            by_source[str(item.get("source", "Unknown"))] += 1
            by_topic[str(item.get("title", "Untitled"))] += 1
        return {
            "by_owner": dict(by_owner),
            "by_folder": dict(by_folder),
            "by_source": dict(by_source),
            "by_topic": dict(by_topic),
        }

    def response_metrics(self, emails: Iterable[Mapping[str, object]]) -> dict[str, float | int]:
        """Estimate response-time metrics from conversation message order."""

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

        response_times_hours: list[float] = []
        waiting_on_me = 0
        waiting_on_others = 0

        for messages in grouped.values():
            ordered = sorted(messages, key=lambda message: self._timestamp(message) or datetime.min)
            if not ordered:
                continue
            latest = ordered[-1]
            if self._is_outgoing(latest):
                waiting_on_others += 1
            else:
                waiting_on_me += 1
            for current, following in zip(ordered, ordered[1:]):
                if not self._is_outgoing(current) and self._is_outgoing(following):
                    current_time = self._timestamp(current)
                    response_time = self._timestamp(following)
                    if current_time and response_time and response_time >= current_time:
                        response_times_hours.append((response_time - current_time).total_seconds() / 3600)

        return {
            "average_response_time_hours": round(mean(response_times_hours), 2) if response_times_hours else 0.0,
            "waiting_on_me": waiting_on_me,
            "waiting_on_others": waiting_on_others,
        }

    @staticmethod
    def _timestamp(record: Mapping[str, object]) -> datetime | None:
        for key in ("received_time", "ReceivedTime", "sent_time", "SentTime", "last_activity_date", "LastActivityDate"):
            value = record.get(key)
            if isinstance(value, datetime):
                return value
            if isinstance(value, str) and value:
                try:
                    return datetime.fromisoformat(value)
                except ValueError:
                    continue
        return None

    @staticmethod
    def _is_outgoing(record: Mapping[str, object]) -> bool:
        folder = str(record.get("folder") or record.get("Folder") or "").lower()
        return "sent" in folder
