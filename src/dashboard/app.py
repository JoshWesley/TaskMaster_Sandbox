"""Dash application skeleton for the Outlook Intelligence Dashboard."""

from __future__ import annotations

from typing import Any

try:
    from dash import Dash, Input, Output, dash_table, dcc, html
except ImportError:  # pragma: no cover - dependency may not exist in validation environment.
    Dash = None  # type: ignore[assignment]
    Input = Output = dash_table = dcc = html = None  # type: ignore[assignment]


def create_app(snapshot: dict[str, Any] | None = None) -> "Dash":
    """Create a runnable Dash app with placeholder dashboard sections."""

    if Dash is None:
        raise RuntimeError("Dash is not installed. Install requirements.txt to run the dashboard UI.")

    app = Dash(__name__)
    dashboard_snapshot = snapshot or _default_snapshot()
    app.layout = html.Div(
        [
            html.H1("Outlook Intelligence Dashboard"),
            html.P("Conversation-centric work management for Outlook."),
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
            ),
            html.Div(
                [
                    _metric_card("Awaiting My Action", dashboard_snapshot["counts"]["awaiting_my_action"]),
                    _metric_card("Waiting For Others", dashboard_snapshot["counts"]["waiting_for_others"]),
                    _metric_card("External Conversations", dashboard_snapshot["counts"]["external_conversations"]),
                    _metric_card("Top Priorities", dashboard_snapshot["counts"]["top_priorities"]),
                ],
                style={"display": "grid", "gridTemplateColumns": "repeat(4, 1fr)", "gap": "16px", "marginBottom": "24px"},
            ),
            dash_table.DataTable(
                id="backlog-table",
                columns=[
                    {"name": "Title", "id": "title"},
                    {"name": "Owner", "id": "owner"},
                    {"name": "Priority", "id": "priority"},
                    {"name": "Urgency", "id": "urgency_score"},
                    {"name": "Folder", "id": "folder"},
                ],
                data=dashboard_snapshot["backlog"],
                page_size=10,
                style_table={"overflowX": "auto"},
            ),
            dcc.Store(id="backlog-store", data=dashboard_snapshot["backlog"]),
        ],
        style={"maxWidth": "1200px", "margin": "0 auto", "padding": "24px"},
    )

    @app.callback(Output("backlog-table", "data"), Input("ownership-filter", "value"), Input("backlog-store", "data"))
    def filter_backlog(ownership: str, backlog: list[dict[str, Any]]) -> list[dict[str, Any]]:
        if ownership == "All":
            return backlog
        return [item for item in backlog if item.get("owner") == ownership]

    return app


def _metric_card(title: str, value: Any) -> Any:
    return html.Div(
        [html.H3(title), html.P(str(value), style={"fontSize": "28px", "fontWeight": "bold"})],
        style={"border": "1px solid #D9D9D9", "borderRadius": "8px", "padding": "16px", "backgroundColor": "#FAFAFA"},
    )


def _default_snapshot() -> dict[str, Any]:
    return {
        "counts": {
            "awaiting_my_action": 0,
            "waiting_for_others": 0,
            "external_conversations": 0,
            "top_priorities": 0,
        },
        "backlog": [],
    }


if __name__ == "__main__":  # pragma: no cover - manual runtime entry point.
    create_app().run(debug=True)
