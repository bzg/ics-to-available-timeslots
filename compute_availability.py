#!/usr/bin/env python3
"""
Compute available timeslots from an ICS calendar file.

Assumptions:
- Not available on weekends (Saturday-Sunday)
- Working hours: Monday-Friday, 1:30 PM - 5:00 PM (13:30-17:00)
"""

import sys
from bisect import bisect_left
from datetime import datetime, timedelta, time
from typing import List, Tuple
from zoneinfo import ZoneInfo

from icalendar import Calendar, Event
from dateutil.rrule import rrulestr

# Constants
DEFAULT_TZ = ZoneInfo('Europe/Paris')
WORK_START_TIME = time(hour=13, minute=30)
WORK_END_TIME = time(hour=17, minute=0)
AVAILABILITY_LEAD_DAYS = 3  # Working days before availability starts


def parse_ics(filename: str) -> Calendar:
    """Parse an ICS file and return the Calendar object."""
    with open(filename, 'rb') as f:
        return Calendar.from_ical(f.read())


def get_datetime_with_tz(dt, default_tz: ZoneInfo = DEFAULT_TZ) -> datetime:
    """Ensure datetime has timezone information."""
    if not isinstance(dt, datetime):
        # It's a date object, convert to datetime at start of day
        dt = datetime.combine(dt, time.min)
    
    return dt if dt.tzinfo else dt.replace(tzinfo=default_tz)


def expand_recurring_events(event, start_date: datetime, end_date: datetime) -> List[Tuple[datetime, datetime]]:
    """Expand a recurring event into individual occurrences."""
    dtstart = get_datetime_with_tz(event.get('DTSTART').dt)
    dtend = get_datetime_with_tz(event.get('DTEND').dt)
    rrule = event.get('RRULE')

    if not rrule:
        # Not a recurring event
        if start_date <= dtstart <= end_date:
            return [(dtstart, dtend)]
        return []

    # Parse the recurrence rule
    rrule_str = rrule.to_ical().decode('utf-8')
    duration = dtend - dtstart

    # Generate occurrences
    try:
        rule = rrulestr(rrule_str, dtstart=dtstart)
        occurrences = []
        for occurrence_start in rule:
            if occurrence_start > end_date:
                break
            if occurrence_start >= start_date:
                occurrences.append((occurrence_start, occurrence_start + duration))
        return occurrences
    except Exception as e:
        print(f"Warning: Could not parse recurrence rule: {e}", file=sys.stderr)
        # Fall back to single occurrence
        if start_date <= dtstart <= end_date:
            return [(dtstart, dtend)]
        return []


def get_busy_times(cal: Calendar, start_date: datetime, end_date: datetime) -> List[Tuple[datetime, datetime]]:
    """Extract all busy time periods from the calendar."""
    busy_times = []

    for component in cal.walk():
        if component.name == "VEVENT":
            try:
                busy_times.extend(expand_recurring_events(component, start_date, end_date))
            except Exception as e:
                print(f"Warning: Could not process event: {e}", file=sys.stderr)

    # Sort by start time
    busy_times.sort(key=lambda x: x[0])
    return busy_times


def merge_overlapping_intervals(intervals: List[Tuple[datetime, datetime]]) -> List[Tuple[datetime, datetime]]:
    """Merge overlapping time intervals."""
    if not intervals:
        return []

    merged = [intervals[0]]

    for current_start, current_end in intervals[1:]:
        last_start, last_end = merged[-1]

        if current_start <= last_end:
            # Overlapping or adjacent intervals
            merged[-1] = (last_start, max(last_end, current_end))
        else:
            merged.append((current_start, current_end))

    return merged


def generate_working_hours(start_date: datetime, end_date: datetime) -> List[Tuple[datetime, datetime]]:
    """Generate all working hour blocks (Mon-Fri, 13:30-17:00)."""
    working_blocks = []
    current_date = start_date.date()
    end = end_date.date()
    tz = start_date.tzinfo or DEFAULT_TZ

    while current_date <= end:
        # Skip weekends (5=Saturday, 6=Sunday)
        if current_date.weekday() < 5:
            work_start = datetime.combine(current_date, WORK_START_TIME, tzinfo=tz)
            work_end = datetime.combine(current_date, WORK_END_TIME, tzinfo=tz)
            working_blocks.append((work_start, work_end))
        current_date += timedelta(days=1)

    return working_blocks


def calculate_availability_start_date(base_date: datetime) -> datetime:
    """Calculate availability start date as today + AVAILABILITY_LEAD_DAYS working days.

    Working days exclude weekends (Saturday and Sunday).

    Examples (with AVAILABILITY_LEAD_DAYS=3):
        - If today is Tuesday -> Friday (Tue+1=Wed, Wed+1=Thu, Thu+1=Fri)
        - If today is Saturday -> Wednesday (skip Sat/Sun, Mon+1=Tue, Tue+1=Wed)
        - If today is Sunday -> Wednesday (skip Sun, Mon+1=Tue, Tue+1=Wed)
    """
    current = base_date.date()
    working_days_added = 0

    while working_days_added < AVAILABILITY_LEAD_DAYS:
        current += timedelta(days=1)
        if current.weekday() < 5:
            working_days_added += 1

    # Return as datetime at start of day with timezone
    return datetime.combine(current, time.min, tzinfo=base_date.tzinfo)


def compute_available_slots(working_blocks: List[Tuple[datetime, datetime]],
                           busy_times: List[Tuple[datetime, datetime]],
                           min_duration_minutes: int = 44) -> List[Tuple[datetime, datetime]]:
    """Compute available time slots by subtracting busy times from working hours.

    Args:
        working_blocks: List of working hour time blocks
        busy_times: List of busy time periods
        min_duration_minutes: Minimum duration in minutes for a slot to be included (default: 44)

    Returns:
        List of available time slots that meet the minimum duration requirement
    """
    if not working_blocks:
        return []

    # Merge overlapping busy times
    busy_times = merge_overlapping_intervals(busy_times)
    
    if not busy_times:
        # No busy times: all working blocks are available (if they meet min duration)
        min_duration = timedelta(minutes=min_duration_minutes)
        return [(start, end) for start, end in working_blocks if end - start > min_duration]

    # Pre-compute for binary search
    busy_starts = [b[0] for b in busy_times]
    min_duration = timedelta(minutes=min_duration_minutes)
    available_slots = []

    for work_start, work_end in working_blocks:
        # Find first busy slot that might overlap with this working block
        # We want busy slots where busy_end > work_start
        idx = bisect_left(busy_starts, work_start)
        # Check the previous slot too, as it might extend into our work block
        if idx > 0:
            idx -= 1

        current_start = work_start

        while idx < len(busy_times):
            busy_start, busy_end = busy_times[idx]
            
            # If busy time is completely after current work block, we're done
            if busy_start >= work_end:
                break

            # If busy time is completely before current position, skip it
            if busy_end <= current_start:
                idx += 1
                continue

            # If there's a gap between current position and busy start
            if current_start < busy_start:
                slot_end = min(busy_start, work_end)
                if slot_end - current_start > min_duration:
                    available_slots.append((current_start, slot_end))

            # Move current position to after the busy time
            current_start = max(current_start, busy_end)
            idx += 1

        # Add remaining time in work block if any
        if current_start < work_end and work_end - current_start > min_duration:
            available_slots.append((current_start, work_end))

    return available_slots


def create_availability_calendar(available_slots: List[Tuple[datetime, datetime]]) -> Calendar:
    """Create a new ICS calendar with available time slots."""
    cal = Calendar()
    cal.add('prodid', '-//Available Timeslots//bzg//')
    cal.add('version', '2.0')
    cal.add('calscale', 'GREGORIAN')
    cal.add('x-wr-calname', 'Available Slots')
    cal.add('x-wr-timezone', 'Europe/Paris')

    now = datetime.now(ZoneInfo('UTC'))
    
    for idx, (start, end) in enumerate(available_slots):
        event = Event()
        event.add('summary', 'Available')
        event.add('dtstart', start)
        event.add('dtend', end)
        event.add('dtstamp', now)
        event.add('uid', f'available-{idx}-{start.strftime("%Y%m%d%H%M%S")}@bzg')
        event.add('status', 'TENTATIVE')
        event.add('transp', 'TRANSPARENT')
        cal.add_component(event)

    return cal


def main():
    if len(sys.argv) < 2:
        print("Usage: python compute_availability.py <input.ics> [output.ics] [weeks_ahead]", file=sys.stderr)
        print("\nDefaults:", file=sys.stderr)
        print("  output.ics = stdout (use '-' or omit for stdout)", file=sys.stderr)
        print("  weeks_ahead = 4", file=sys.stderr)
        print("\nExamples:", file=sys.stderr)
        print("  python compute_availability.py input.ics              # Output to stdout, 4 weeks", file=sys.stderr)
        print("  python compute_availability.py input.ics 8            # Output to stdout, 8 weeks", file=sys.stderr)
        print("  python compute_availability.py input.ics out.ics     # Output to file, 4 weeks", file=sys.stderr)
        print("  python compute_availability.py input.ics out.ics 8   # Output to file, 8 weeks", file=sys.stderr)
        sys.exit(1)

    input_file = sys.argv[1]

    # Parse arguments intelligently
    output_file = None
    weeks_ahead = 4

    if len(sys.argv) > 2:
        second_arg = sys.argv[2]
        try:
            weeks_ahead = int(second_arg)
        except ValueError:
            output_file = second_arg if second_arg != '-' else None

        if len(sys.argv) > 3:
            weeks_ahead = int(sys.argv[3])

    # Set up date range
    today = datetime.now(DEFAULT_TZ).replace(hour=0, minute=0, second=0, microsecond=0)
    start_date = calculate_availability_start_date(today)
    end_date = start_date + timedelta(weeks=weeks_ahead)

    print(f"Today: {today.date()}", file=sys.stderr)
    print(f"Availability starts on: {start_date.date()} (today + {AVAILABILITY_LEAD_DAYS} working days)", file=sys.stderr)
    print(f"Processing calendar from {start_date.date()} to {end_date.date()}", file=sys.stderr)

    # Parse input calendar
    print(f"Reading {input_file}...", file=sys.stderr)
    cal = parse_ics(input_file)

    # Get busy times
    print("Extracting busy times...", file=sys.stderr)
    busy_times = get_busy_times(cal, start_date, end_date)
    print(f"Found {len(busy_times)} busy time slots", file=sys.stderr)

    # Generate working hours
    print("Generating working hours (Mon-Fri, 13:30-17:00)...", file=sys.stderr)
    working_blocks = generate_working_hours(start_date, end_date)
    print(f"Generated {len(working_blocks)} working hour blocks", file=sys.stderr)

    # Compute available slots
    print("Computing available slots (skipping slots ≤44 minutes)...", file=sys.stderr)
    available_slots = compute_available_slots(working_blocks, busy_times)
    print(f"Found {len(available_slots)} available time slots", file=sys.stderr)

    # Create output calendar
    print("Creating availability calendar...", file=sys.stderr)
    availability_cal = create_availability_calendar(available_slots)

    # Write to file or stdout
    ical_data = availability_cal.to_ical()

    if output_file is None:
        sys.stdout.buffer.write(ical_data)
    else:
        with open(output_file, 'wb') as f:
            f.write(ical_data)

        total_hours = sum((end - start).total_seconds() / 3600 for start, end in available_slots)
        print(f"\n✓ Successfully created {output_file}", file=sys.stderr)
        print(f"\nSummary:", file=sys.stderr)
        print(f"  - Total available hours: {total_hours:.1f}h", file=sys.stderr)
        print(f"  - Available slots: {len(available_slots)}", file=sys.stderr)

        if available_slots:
            print(f"\nNext few available slots:", file=sys.stderr)
            for start, end in available_slots[:5]:
                duration = (end - start).total_seconds() / 3600
                print(f"  • {start.strftime('%a %Y-%m-%d %H:%M')} - {end.strftime('%H:%M')} ({duration:.1f}h)", file=sys.stderr)


if __name__ == "__main__":
    main()
