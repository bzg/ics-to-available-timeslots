#!/usr/bin/env python3
"""
Export an ICS calendar file to HTML or plain-text format.

Configuration via environment variables:
- TITLE      Page / document title  (default: "Available Time Slots")
- ICAL_URL   Link shown in the HTML header (default: "https://bzg.fr/agenda.ics")
- VISIO_URL  Visio link shown in the HTML header
              (default: "https://rendez-vous.renater.fr/swh-partnerships")
"""

from __future__ import annotations

import argparse
import html
import logging
import os
import sys
from collections import defaultdict
from datetime import date, datetime, time, timedelta
from typing import NamedTuple, Sequence
from textwrap import dedent

from icalendar import Calendar

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class SlotEvent(NamedTuple):
    start: datetime
    end: datetime
    summary: str

    @property
    def duration_label(self) -> str:
        hours = (self.end - self.start).total_seconds() / 3600
        return f"{hours:.1f}h" if hours >= 1 else f"{int(hours * 60)}min"

    @property
    def time_range(self) -> str:
        return f"{self.start:%H:%M} - {self.end:%H:%M}"


# Week key → sorted dict of date → list of events.
WeeklyGroups = dict[tuple[int, int], dict[date, list[SlotEvent]]]

# ---------------------------------------------------------------------------
# ICS helpers
# ---------------------------------------------------------------------------

def parse_ics(source: str) -> Calendar:
    """Parse an ICS file (or ``-`` for stdin) and return the Calendar."""
    data = sys.stdin.buffer.read() if source == "-" else open(source, "rb").read()
    return Calendar.from_ical(data)


def _normalize_dt(dt: datetime | date) -> datetime:
    return dt if isinstance(dt, datetime) else datetime.combine(dt, time.min)


def extract_events(cal: Calendar) -> list[SlotEvent]:
    """Return sorted events from *cal*."""
    events = [
        SlotEvent(
            start=_normalize_dt(c.get("DTSTART").dt),
            end=_normalize_dt(c.get("DTEND").dt),
            summary=str(c.get("SUMMARY", "Available")),
        )
        for c in cal.walk()
        if c.name == "VEVENT"
    ]
    events.sort(key=lambda e: e.start)
    return events


def group_by_week(events: Sequence[SlotEvent]) -> WeeklyGroups:
    """Group events by ISO week, then by date (both sorted)."""
    weeks: dict[tuple[int, int], dict[date, list[SlotEvent]]] = defaultdict(
        lambda: defaultdict(list)
    )
    for ev in events:
        iso_year, week_num, _ = ev.start.isocalendar()
        weeks[(iso_year, week_num)][ev.start.date()].append(ev)

    return {
        wk: dict(sorted(dates.items()))
        for wk, dates in sorted(weeks.items())
    }


def _week_monday(iso_year: int, week_num: int) -> date:
    """Return the Monday of the given ISO week."""
    return date.fromisocalendar(iso_year, week_num, 1)


# ---------------------------------------------------------------------------
# Plain-text renderer
# ---------------------------------------------------------------------------

_RULE = "=" * 70
_THIN = "-" * 70


def render_ascii(events: Sequence[SlotEvent], title: str) -> str:
    """Render events as plain text."""
    if not events:
        return "No available time slots found.\n"

    lines = [_RULE, title.center(70), _RULE, ""]
    weeks = group_by_week(events)

    for idx, ((iso_y, iso_w), dates) in enumerate(weeks.items()):
        if idx:
            lines.append("")
        monday = _week_monday(iso_y, iso_w)
        lines.append(f"WEEK OF {monday.strftime('%B %d, %Y').upper()}")
        lines.append(_THIN)

        for day_events in dates.values():
            lines.append(f"\n  {day_events[0].start:%A, %B %d}")
            for ev in day_events:
                lines.append(f"    {ev.time_range} ({ev.duration_label})")

    lines.extend(["", _RULE, f"Generated on {datetime.now():%Y-%m-%d at %H:%M}", ""])
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# HTML renderer
# ---------------------------------------------------------------------------

_CSS = dedent("""\
    * { margin: 0; padding: 0; box-sizing: border-box; }
    body {
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto,
                     'Helvetica Neue', Arial, sans-serif;
        line-height: 1.6; color: #333; background: #f5f5f5; padding: 20px;
    }
    .container {
        max-width: 900px; margin: 0 auto; background: white;
        border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        overflow: hidden;
    }
    header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white; padding: 30px; text-align: center;
    }
    header h1 { font-size: 28px; font-weight: 600; margin-bottom: 8px; }
    header p  { opacity: 0.9; font-size: 14px; }
    header a  { color: white; }
    .events   { padding: 20px; }
    .week-group  { margin-bottom: 30px; }
    .week-header {
        font-size: 14px; font-weight: 600; color: #666;
        text-transform: uppercase; letter-spacing: 0.5px;
        margin-bottom: 12px; padding-bottom: 8px;
        border-bottom: 2px solid #e0e0e0;
    }
    .event {
        background: white; border-left: 4px solid #667eea; padding: 16px;
        margin-bottom: 12px; border-radius: 6px; transition: all 0.2s;
        border: 1px solid #e0e0e0;
    }
    .event:hover {
        border-left-color: #764ba2;
        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        transform: translateX(4px);
    }
    .event-date {
        font-weight: 600; color: #333; font-size: 16px;
        margin-bottom: 12px; padding-bottom: 8px;
        border-bottom: 1px solid #f0f0f0;
    }
    .event-time {
        color: #666; font-size: 14px; padding: 8px 0;
        display: flex; align-items: center; justify-content: space-between;
    }
    .event-time:not(:last-child) { border-bottom: 1px dashed #e8e8e8; }
    .time-range { flex: 1; }
    .event-duration {
        display: inline-block; background: #e8eaf6; color: #667eea;
        padding: 2px 8px; border-radius: 4px; font-size: 12px;
        font-weight: 600; margin-left: 8px;
    }
    .no-events {
        text-align: center; padding: 60px 20px; color: #999;
    }
    footer {
        text-align: center; padding: 20px; color: #999;
        font-size: 12px; border-top: 1px solid #e0e0e0;
    }
    @media (max-width: 600px) { header h1 { font-size: 24px; } }
""")


def _render_events_html(events: Sequence[SlotEvent]) -> str:
    """Build the <div class="events">…</div> fragment."""
    if not events:
        return (
            '        <div class="no-events">\n'
            "            <p>No available time slots found.</p>\n"
            "        </div>"
        )

    parts: list[str] = ['        <div class="events">']
    weeks = group_by_week(events)

    for (iso_y, iso_w), dates in weeks.items():
        monday = _week_monday(iso_y, iso_w)
        parts.append(
            f'            <div class="week-group">\n'
            f'                <div class="week-header">'
            f"Week of {monday.strftime('%B %d, %Y')}</div>"
        )

        for day_events in dates.values():
            date_str = html.escape(f"{day_events[0].start:%A, %B %d}")
            parts.append(
                f'                <div class="event">\n'
                f'                    <div class="event-date">{date_str}</div>'
            )
            for ev in day_events:
                parts.append(
                    f'                    <div class="event-time">\n'
                    f'                        <span class="time-range">'
                    f"{html.escape(ev.time_range)}</span>\n"
                    f'                        <span class="event-duration">'
                    f"{html.escape(ev.duration_label)}</span>\n"
                    f"                    </div>"
                )
            parts.append("                </div>")
        parts.append("            </div>")

    parts.append("        </div>")
    return "\n".join(parts)


def render_html(events: Sequence[SlotEvent], title: str) -> str:
    """Render a full HTML page."""
    ical_url = os.environ.get("ICAL_URL", "https://bzg.fr/agenda.ics")
    visio_url = os.environ.get(
        "VISIO_URL", "https://rendez-vous.renater.fr/swh-partnerships"
    )
    safe_title = html.escape(title)
    generated_at = f"{datetime.now():%Y-%m-%d at %H:%M}"
    events_html = _render_events_html(events)

    return dedent(f"""\
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>{safe_title}</title>
        <style>
    {_CSS}    </style>
    </head>
    <body>
        <div class="container">
            <header>
                <h1>{safe_title}</h1>
                <p><strong><a href="{html.escape(ical_url)}">iCal file</a> \
    - <a href="{html.escape(visio_url)}">Visio link</a></strong></p>
            </header>
    {events_html}
            <footer>
                Generated on {generated_at}
            </footer>
        </div>
    </body>
    </html>
    """)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

_FORMAT_ALIASES = {"text": "ascii", "txt": "ascii"}
_RENDERERS = {
    "html": render_html,
    "ascii": render_ascii,
}


def _resolve_format(name: str) -> str:
    name = name.lower()
    name = _FORMAT_ALIASES.get(name, name)
    if name not in _RENDERERS:
        raise argparse.ArgumentTypeError(
            f"Unknown format '{name}'. Choose from: html, ascii (text/txt)."
        )
    return name


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Export an ICS calendar to HTML or plain text.",
    )
    p.add_argument("input", help="Input .ics file (use '-' for stdin)")
    p.add_argument(
        "output",
        nargs="?",
        default="-",
        help="Output file (default: stdout)",
    )
    p.add_argument(
        "-f",
        "--format",
        type=_resolve_format,
        default="html",
        dest="fmt",
        help="Output format: html, ascii/text/txt (default: html)",
    )
    p.add_argument(
        "-t",
        "--title",
        default=None,
        help='Title (default: $TITLE or "Available Time Slots")',
    )
    p.add_argument("-v", "--verbose", action="store_true")
    return p.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_args(argv)

    logging.basicConfig(
        format="%(message)s",
        level=logging.DEBUG if args.verbose else logging.INFO,
        stream=sys.stderr,
    )

    title = args.title or os.environ.get("TITLE", "Available Time Slots")

    log.info("Reading %s…", args.input)
    cal = parse_ics(args.input)

    events = extract_events(cal)
    log.info("Found %d events", len(events))

    render = _RENDERERS[args.fmt]
    content = render(events, title)

    if args.output == "-":
        sys.stdout.write(content)
    else:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(content)
        log.info("✓ Written to %s", args.output)


if __name__ == "__main__":
    main()
