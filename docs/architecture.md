# Architecture

The Outlook Intelligence Dashboard is a local-first Python application that converts Outlook conversations into prioritised work items.

## High-Level Flow

1. **Outlook COM Extractor** reads Inbox, Sent Items, and selected folders through `win32com.client`.
2. **Conversation Engine** groups messages by `ConversationID`, detects action language, and assigns ownership.
3. **Backlog Manager** calculates urgency scores, system priority, and stable ordering for the single backlog.
4. **SQLite Storage** caches normalized emails, conversation summaries, work items, and sync state.
5. **Dash UI** presents executive metrics and backlog views without exposing mailbox data outside the machine.

## Design Notes

- The primary management unit is the **conversation**, not the individual email.
- Processing stays local and uses rule-based heuristics rather than external AI services.
- SQLite supports incremental synchronization, persisted ranking, and fast startup for large mailboxes.
- The structure leaves room for future classification, saved views, and richer analytics modules.
