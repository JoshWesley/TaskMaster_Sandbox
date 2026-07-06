# Outlook Intelligence Dashboard

## Personal Work Management System Built on Outlook

Version: 1.0

---

# Vision

Transform Microsoft Outlook from a chronological communication tool into a prioritised work management platform.

The dashboard should help answer:

1. What requires my attention right now?
2. What am I waiting on from others?
3. What can I safely ignore?
4. What should I work on next?
5. Which customers, colleagues or projects are generating the most workload?

This application is not intended to be another email client.

Its purpose is to convert communications into actionable work items and make workload visible, manageable and prioritised.

---

# Background

As a senior manager I receive a large volume of communications from:

- Internal colleagues
- Customers
- Suppliers
- Partners
- Leadership teams

Standard Outlook views are optimized for reading email chronologically.

They do not provide sufficient visibility into:

- Which conversations require action
- Which communications are informational only
- Which topics are becoming urgent
- Which customers are creating the most demand
- Which conversations are blocked
- What should be worked on next

The goal is to create an Inbox Intelligence Platform that behaves more like a personal Jira and Kanban system than an email client.

---

# Scope

## Included In Version 1

### Outlook

- Inbox
- Sent Items
- Custom Outlook folders
- Categories
- Flags
- Conversation Threads
- Read/Unread Status
- Importance

### Manual Tasks

User-created tasks independent of email.

### Meeting Actions

Actions manually created from meetings.

### Dashboarding

- Executive dashboard
- Kanban board
- Prioritised backlog
- Conversation analysis
- Folder dashboards
- Sender views
- Domain views
- Analytics

---

## Explicitly Out Of Scope

Not required for Version 1:

- Microsoft Graph
- Teams integration
- Planner integration
- Azure services
- Cloud services
- External AI services
- LLM-based classification
- Semantic search

The architecture should allow future addition of these capabilities.

---

# Design Principles

## Action-Oriented

Focus on work requiring attention rather than unread emails.

## Conversation-Oriented

Manage conversations rather than individual messages.

## Prioritised

Everything should contribute to understanding what should be worked on next.

## Visual

Users should understand workload immediately.

## Local First

All processing occurs locally.

## Privacy Focused

No external services.

No cloud processing.

No mailbox data leaves the local device.

---

# Technical Constraints

## Data Source

Use Outlook COM.

Required:

```python
win32com.client
```

Supported Client:

```text
Classic Outlook for Windows
```

Must Not Require:

```text
Microsoft Graph
Azure App Registration
OAuth
Cloud Connectivity
External APIs
```

---

# Technology Stack

## Backend

```text
Python 3.11+
```

Recommended Packages:

```text
pywin32
pandas
numpy
sqlite3
rapidfuzz
scikit-learn
python-dateutil
```

---

## Frontend

Preferred:

```text
Dash
```

Alternative:

```text
React
```

---

## Storage

```text
SQLite
```

All data stored locally.

---

# High Level Architecture

```text
Outlook COM
      │
      ▼
Email Extractor
      │
      ▼
Conversation Engine
      │
      ▼
Classification Engine
      │
      ▼
Work Item Engine
      │
      ▼
SQLite Cache
      │
      ▼
Dashboard UI
```

---

# Core Object Model

All work should be represented using a common object.

## Work Item

```text
ID
Title
Source
Status
Priority
UrgencyScore
Owner
CreatedDate
LastActivityDate
DueDate
ConversationID
Notes
Tags
Folder
```

---

# Supported Sources

## Version 1

```text
Email
Manual Task
Meeting Action
```

## Future

```text
Teams
Planner
Graph
AI Generated Actions
```

---

# Email Data Model

Store:

```text
EntryID
ConversationID
ConversationTopic
Subject
SenderName
SenderEmail
SenderDomain
Recipients
CC
ReceivedTime
SentTime
BodyPreview
Folder
Category
FlagStatus
ReadStatus
Importance
```

---

# Internal / External Classification

## Internal

Any address ending with:

```text
@motionapplied.com
```

---

## External

Any other domain.

---

## Conversation Types

```text
Internal
External
Mixed
```

Examples:

```text
Motion Applied only
=> Internal

Motion Applied + Customer
=> Mixed

Customer only
=> External
```

---

# Conversation Engine

The primary unit of management is:

```text
Conversation
```

Not individual emails.

Grouping Key:

```text
ConversationID
```

Metrics:

```text
Topic
Participants
Message Count
Start Date
Last Activity
Latest Sender
Folder
Internal/External Status
Urgency
Priority
```

---

# Ownership Detection

Determine who owns the next action.

---

## Awaiting My Action

Conditions:

- Latest email received by me
- Question asked
- Request made
- Action requested
- No response sent by me

---

## Waiting For Others

Conditions:

- Latest message sent by me
- Waiting for response

---

## Information Only

Indicators:

- FYI
- Notification
- Status update
- Automated email
- Distribution email

No action required.

---

## Unknown

Unable to determine confidently.

---

# Action Detection

Initial implementation should be rule-based.

Search for phrases such as:

```text
Can you
Could you
Please
Need
Required
Review
Approve
Confirm
Provide
Action Required
Response Needed
Deadline
Urgent
ASAP
Blocking
```

Outputs:

```text
Action Required
Confidence Score
Action Summary
```

---

# Priority System

Each work item contains:

```text
System Priority
User Priority
Final Priority
```

---

# User Priority

User priority always takes precedence.

The user must be able to:

```text
Drag
Drop
Rank
Pin
Override
```

items manually.

---

# Priority Levels

```text
P1 Critical
P2 High
P3 Normal
P4 Low
P5 Someday
```

---

# Alternative Workflow States

Optional workflow:

```text
Today
This Week
Waiting
Deferred
Someday
```

---

# Backlog Management

Maintain a single prioritised backlog.

Each conversation appears once.

Example:

```text
1. Customer Escalation
2. Portal Improvement Discussion
3. Telemetry Investigation
4. Software Release Approval
5. KPI Dashboard Development
```

Ranking persists between sessions.

---

# Urgency Engine

Generate:

```text
0-100
```

urgency score.

---

## Positive Signals

```text
Unread
Flagged
High Importance
External Sender
Deadline Mentioned
Urgent Mentioned
Blocking Mentioned
Recent Activity
Repeated Follow-Up
```

---

## Negative Signals

```text
Information Only
Already Replied
Inactive Conversation
Completed
```

---

## Visualisation

```text
Green
Amber
Red
```

based on urgency.

---

# Manual Tasks

Users must be able to create tasks manually.

Examples:

```text
Prepare Leadership Presentation

Review FAE Organisation

AI Strategy Proposal

Portal Roadmap Planning
```

Fields:

```text
Title
Description
Priority
DueDate
Tags
Category
Status
```

---

# Meeting Actions

Support manually created meeting actions.

Example:

```text
Action:
Provide Release Timeline

Owner:
Josh

Due:
Friday
```

Stored using the common Work Item model.

---

# Folder Dashboard System

Users should be able to create dashboards for Outlook folders.

Examples:

```text
Inbox
Customer Support
Product Management
Telemetry
Management
Portal
Waiting For Customer
```

Users can configure:

```text
Filters
Grouping
Columns
Sorting
Views
```

Configuration should persist.

---

# Dashboard Views

## Executive Dashboard

Display:

```text
Awaiting My Action

Waiting For Others

External Conversations

Internal Conversations

Flagged Work

Overdue Work

Manual Tasks

Top Priorities
```

---

## Kanban Board

Columns:

```text
Awaiting My Action

Waiting For Others

Information Only

Scheduled Follow Up

Completed
```

Card Contents:

```text
Subject
Company
Latest Sender
Age
Urgency
Priority
Message Count
```

---

## Backlog View

Single prioritised list.

Supports:

```text
Sorting
Filtering
Drag And Drop
Persistent Ordering
```

---

## Conversation View

Display:

```text
Participants
Messages
Ownership
Priority
Urgency
History
```

---

## Sender View

Metrics:

```text
Conversation Count
Open Actions
Average Urgency
Last Contact Date
```

---

## Domain View

Group by:

```text
motionapplied.com

Customer Domains

Supplier Domains

Partner Domains
```

Display:

```text
Conversation Count
Action Count
Urgency
```

---

# Search

## Standard Search

Search by:

```text
Sender
Subject
Domain
Conversation
Body Preview
```

---

## Saved Searches

Examples:

```text
Urgent External

Awaiting Response

Customer Escalations

Leadership Topics

My Actions
```

---

# Analytics

## Email Load

```text
Emails Per Day
Emails Per Week
Emails Per Month
```

---

## Workload Analysis

```text
Work By Person

Work By Company

Work By Folder

Work By Topic
```

---

## Response Metrics

```text
Average Response Time

Waiting On Me

Waiting On Others
```

---

# Performance Requirements

Target Mailbox Size:

```text
10,000 - 100,000 emails
```

Requirements:

```text
Incremental Synchronisation

SQLite Cache

Background Refresh

Fast Startup

Minimal Outlook Queries
```

The application must never reprocess the entire mailbox during normal operation.

---

# Repository Structure

```text
outlook-intelligence-dashboard/
│
├── README.md
├── requirements.txt
├── src/
│   ├── outlook/
│   ├── conversations/
│   ├── backlog/
│   ├── analytics/
│   ├── dashboard/
│   ├── tasks/
│   └── storage/
│
├── database/
│   └── dashboard.db
│
├── config/
│   └── settings.json
│
└── docs/
    ├── architecture.md
    ├── data-model.md
    ├── roadmap.md
```

---

# Future Roadmap

Potential future enhancements:

```text
Microsoft Graph
Teams Integration
Planner Integration
AI Summaries
LLM Classification
Semantic Search
Vector Database
Meeting Transcript Processing
```

These capabilities should not require redesign of the core Work Item model.

---

# Success Criteria

The system succeeds if the user can immediately answer:

- What requires my action today?
- What should I do next?
- What am I waiting for?
- What is blocked?
- Which customers need attention?
- Which conversations are urgent?
- What can safely be ignored?

The final experience should feel like:

Personal Jira
+
Personal Kanban
+
Inbox Intelligence Platform

rather than a traditional email client.