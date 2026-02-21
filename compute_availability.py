#!/usr/bin/env python3
"""
Compute available timeslots from an ICS calendar file.

Assumptions:
- Not available on weekends (Saturday and Sunday).
- Working hours: Monday-Friday, 1:30 PM - 5:00 PM (13:30-17:00).

Configuration via environment variables:
- WORK_START       Working day start time (default: "13:30")
- WORK_END         Working day end time (default: "17:00")
- LEAD_DAYS        Working days to skip before first available slot (default: 3)
- BUFFER_MINUTES   Break around each appointment in minutes (default: 10)
- MIN_SLOT_MINUTES Minimum slot duration to keep in minutes (default: 44)
- TZ               Timezone (default: "Europe/Paris")
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
from bisect import bisect_left
from dataclasses import dataclass
from datetime import datetime, date, time, timedelta
from typing import Generator, Iterable, Sequence
from zoneinfo import ZoneInfo

from dateutil.rrule import rrulestr
from icalendar import Calendar, Event

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Type alias
# ---------------------------------------------------------------------------
Interval = tuple[datetime, datetime]


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
def _parse_time(val: str) -> time:
    """Parse 'HH:MM' into a time object."""
    h, m = val.split(":")
    return time(hour=int(h), minute=int(m))


@dataclass(frozen=True)
class Config:
    """All tuneable knobs, loaded from environment with sensible defaults."""

    tz: ZoneInfo
    work_start: time
    work_end: time
    lead_days: int
    buffer_minutes: int
    min_slot_minutes: int

    @classmethod
    def from_env(cls) -> Config:
        return cls(
            tz=ZoneInfo(os.environ.get("TZ", "Europe/Paris")),
            work_start=_parse_time(os.environ.get("WORK_START", "13:30")),
            work_end=_parse_time(os.environ.get("WORK_END", "17:00")),
            lead_days=int(os.environ.get("LEAD_DAYS", "3")),
            buffer_minutes=int(os.environ.get("BUFFER_MINUTES", "10")),
            min_slot_minutes=int(os.environ.get("MIN_SLOT_MINUTES", "44")),
        )


# ---------------------------------------------------------------------------
# ICS helpers
# ---------------------------------------------------------------------------
def parse_ics(filename: str) -> Calendar:
    """Parse an ICS file and return the Calendar object."""
    with open(filename, "rb") as f:
        return Calendar.from_ical(f.read())


def _ensure_datetime(dt: datetime | date, tz: ZoneInfo) -> datetime:
    """Normalise a date or naive datetime to a timezone-aware datetime."""
    if not isinstance(dt, datetime):
        dt = datetime.combine(dt, time.min)
    return dt if dt.tzinfo else dt.replace(tzinfo=tz)


def _expand_event(
    event: Event, window_start: datetime, window_end: datetime, tz: ZoneInfo
) -> Generator[Interval, None, None]:
    """Yield (start, end) pairs for every occurrence of *event* inside the window."""
    dtstart = _ensure_datetime(event.get("DTSTART").dt, tz)
    dtend = _ensure_datetime(event.get("DTEND").dt, tz)
    rrule = event.get("RRULE")

    if not rrule:
        if window_start <= dtstart <= window_end:
            yield (dtstart, dtend)
        return

    duration = dtend - dtstart
    try:
        rule = rrulestr(rrule.to_ical().decode("utf-8"), dtstart=dtstart)
        for occ in rule:
            if occ > window_end:
                break
            if occ >= window_start:
                yield (occ, occ + duration)
    except Exception:
        log.warning("Could not parse recurrence rule for %s", event.get("SUMMARY", "?"))
        if window_start <= dtstart <= window_end:
            yield (dtstart, dtend)


def collect_busy_times(
    cal: Calendar, window_start: datetime, window_end: datetime, tz: ZoneInfo
) -> list[Interval]:
    """Return sorted busy intervals from *cal* within the given window."""
    busy: list[Interval] = []
    for component in cal.walk():
        if component.name == "VEVENT":
            try:
                busy.extend(_expand_event(component, window_start, window_end, tz))
            except Exception:
                log.warning("Skipping unparseable event: %s", component.get("SUMMARY", "?"))
    busy.sort(key=lambda iv: iv[0])
    return busy


# ---------------------------------------------------------------------------
# Interval arithmetic
# ---------------------------------------------------------------------------
def merge_intervals(intervals: Iterable[Interval]) -> list[Interval]:
    """Merge overlapping or adjacent intervals (input must be sorted)."""
    merged: list[Interval] = []
    for start, end in intervals:
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return merged


def _working_hours(
    start_date: datetime, end_date: datetime, cfg: Config
) -> Generator[Interval, None, None]:
    """Yield one (start, end) working-hour block per weekday in the range."""
    day = start_date.date()
    last = end_date.date()
    while day <= last:
        if day.weekday() < 5:  # Mon-Fri
            yield (
                datetime.combine(day, cfg.work_start, tzinfo=cfg.tz),
                datetime.combine(day, cfg.work_end, tzinfo=cfg.tz),
            )
        day += timedelta(days=1)


def add_working_days(base: datetime, n: int) -> datetime:
    """Advance *base* by *n* working days (skipping weekends).

    Examples (with n=3):
        Tuesday  → Friday   (Wed, Thu, Fri)
        Saturday → Wednesday (Mon, Tue, Wed)
        Sunday   → Wednesday (Mon, Tue, Wed)
    """
    current = base.date()
    added = 0
    while added < n:
        current += timedelta(days=1)
        if current.weekday() < 5:
            added += 1
    return datetime.combine(current, time.min, tzinfo=base.tzinfo)


def compute_available_slots(
    working_blocks: Sequence[Interval],
    busy_times: Sequence[Interval],
    cfg: Config,
) -> list[Interval]:
    """Subtract buffered busy times from working blocks.

    Returns available intervals whose duration strictly exceeds
    *cfg.min_slot_minutes*.
    """
    if not working_blocks:
        return []

    # Expand busy times by the buffer on each side, then merge.
    buffer = timedelta(minutes=cfg.buffer_minutes)
    if buffer:
        padded = sorted(
            ((s - buffer, e + buffer) for s, e in busy_times),
            key=lambda iv: iv[0],
        )
    else:
        padded = sorted(busy_times, key=lambda iv: iv[0])
    merged_busy = merge_intervals(padded)

    min_dur = timedelta(minutes=cfg.min_slot_minutes)

    if not merged_busy:
        return [(s, e) for s, e in working_blocks if e - s > min_dur]

    busy_starts = [iv[0] for iv in merged_busy]
    slots: list[Interval] = []

    for work_start, work_end in working_blocks:
        idx = max(bisect_left(busy_starts, work_start) - 1, 0)
        cursor = work_start

        while idx < len(merged_busy):
            b_start, b_end = merged_busy[idx]

            if b_start >= work_end:
                break
            if b_end <= cursor:
                idx += 1
                continue

            # Free gap before this busy block
            if cursor < b_start:
                gap_end = min(b_start, work_end)
                if gap_end - cursor > min_dur:
                    slots.append((cursor, gap_end))

            cursor = max(cursor, b_end)
            idx += 1

        # Trailing free time
        if cursor < work_end and work_end - cursor > min_dur:
            slots.append((cursor, work_end))

    return slots


# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------
def build_availability_calendar(slots: Sequence[Interval]) -> Calendar:
    """Create a new ICS calendar advertising the given available slots."""
    cal = Calendar()
    cal.add("prodid", "-//Available Timeslots//bzg//")
    cal.add("version", "2.0")
    cal.add("calscale", "GREGORIAN")
    cal.add("x-wr-calname", "Available Slots")
    cal.add("x-wr-timezone", "Europe/Paris")

    now = datetime.now(ZoneInfo("UTC"))

    for idx, (start, end) in enumerate(slots):
        event = Event()
        event.add("summary", "Available")
        event.add("dtstart", start)
        event.add("dtend", end)
        event.add("dtstamp", now)
        event.add("uid", f"available-{idx}-{start:%Y%m%d%H%M%S}@bzg")
        event.add("status", "TENTATIVE")
        event.add("transp", "TRANSPARENT")
        cal.add_component(event)

    return cal


def print_summary(slots: Sequence[Interval], preview: int = 5) -> None:
    """Log a human-friendly summary of the available slots."""
    total_h = sum((e - s).total_seconds() / 3600 for s, e in slots)
    log.info("Available slots: %d (%.1f h total)", len(slots), total_h)
    for start, end in slots[:preview]:
        dur = (end - start).total_seconds() / 3600
        log.info(
            "  • %s – %s (%.1fh)",
            f"{start:%a %Y-%m-%d %H:%M}",
            f"{end:%H:%M}",
            dur,
        )


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Compute available timeslots from an ICS calendar.",
    )
    p.add_argument("input", help="Input .ics file")
    p.add_argument(
        "output",
        nargs="?",
        default="-",
        help="Output .ics file (default: stdout)",
    )
    p.add_argument(
        "-w",
        "--weeks",
        type=int,
        default=4,
        help="How many weeks ahead to scan (default: 4)",
    )
    p.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Increase log verbosity",
    )
    return p.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> None:
    args = parse_args(argv)
    cfg = Config.from_env()

    logging.basicConfig(
        format="%(message)s",
        level=logging.DEBUG if args.verbose else logging.INFO,
        stream=sys.stderr,
    )

    today = datetime.now(cfg.tz).replace(hour=0, minute=0, second=0, microsecond=0)
    start_date = add_working_days(today, cfg.lead_days)
    end_date = start_date + timedelta(weeks=args.weeks)

    log.info("Today:  %s", today.date())
    log.info(
        "Window: %s → %s (%d working-day lead)",
        start_date.date(),
        end_date.date(),
        cfg.lead_days,
    )
    log.info(
        "Hours:  %s–%s, buffer %d min, min slot %d min",
        cfg.work_start.strftime("%H:%M"),
        cfg.work_end.strftime("%H:%M"),
        cfg.buffer_minutes,
        cfg.min_slot_minutes,
    )

    cal = parse_ics(args.input)
    busy = collect_busy_times(cal, start_date, end_date, cfg.tz)
    log.info("Busy slots found: %d", len(busy))

    working = list(_working_hours(start_date, end_date, cfg))
    log.debug("Working blocks: %d", len(working))

    available = compute_available_slots(working, busy, cfg)
    print_summary(available)

    ical_data = build_availability_calendar(available).to_ical()

    if args.output == "-":
        sys.stdout.buffer.write(ical_data)
    else:
        with open(args.output, "wb") as f:
            f.write(ical_data)
        log.info("✓ Written to %s", args.output)


if __name__ == "__main__":
    main()
