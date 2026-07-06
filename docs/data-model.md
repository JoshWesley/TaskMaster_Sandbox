# Data Model

## Work Item Model

All actionable records use the same base model:

- `ID`
- `Title`
- `Source`
- `Status`
- `Priority`
- `UrgencyScore`
- `Owner`
- `CreatedDate`
- `LastActivityDate`
- `DueDate`
- `ConversationID`
- `Notes`
- `Tags`
- `Folder`

Supported version 1 sources are **Email**, **Manual Task**, and **Meeting Action**.

## Email Model

Normalized email records capture:

- `EntryID`
- `ConversationID`
- `ConversationTopic`
- `Subject`
- `SenderName`
- `SenderEmail`
- `SenderDomain`
- `Recipients`
- `CC`
- `ReceivedTime`
- `SentTime`
- `BodyPreview`
- `Folder`
- `Category`
- `FlagStatus`
- `ReadStatus`
- `Importance`

## Conversation Summary

Derived conversation records store topic, participants, message count, last activity, latest sender, ownership, action signals, and internal/external classification.

## Storage Tables

The initial SQLite schema includes:

- `emails`
- `conversations`
- `work_items`
- `sync_state`
- `dashboard_views`

These tables provide the baseline for incremental sync, backlog persistence, and dashboard rendering.
