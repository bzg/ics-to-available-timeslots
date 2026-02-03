#!/usr/bin/env python3
"""
Export ICS calendar file to HTML or ASCII text format.
"""

import sys
from datetime import datetime, timedelta, time
from typing import List, Dict
from collections import defaultdict

from icalendar import Calendar


def parse_ics(filename: str) -> Calendar:
    """Parse an ICS file and return the Calendar object."""
    if filename == '-':
        return Calendar.from_ical(sys.stdin.buffer.read())
    with open(filename, 'rb') as f:
        return Calendar.from_ical(f.read())


def get_events(cal: Calendar) -> List[Dict]:
    """Extract events from calendar and sort by start time."""
    def normalize_dt(dt):
        return dt if isinstance(dt, datetime) else datetime.combine(dt, time.min)

    events = [
        {
            'start': normalize_dt(c.get('DTSTART').dt),
            'end': normalize_dt(c.get('DTEND').dt),
            'summary': str(c.get('SUMMARY', 'Available'))
        }
        for c in cal.walk() if c.name == "VEVENT"
    ]
    return sorted(events, key=lambda x: x['start'])


def format_duration(start: datetime, end: datetime) -> str:
    """Format duration in a human-readable way."""
    hours = (end - start).total_seconds() / 3600
    if hours >= 1:
        return f"{hours:.1f}h"
    return f"{int(hours * 60)}min"


def group_events_by_week(events: List[Dict]) -> Dict:
    """Group events by ISO week and then by date.
    
    Returns a sorted dict: {(iso_year, week_num): {date: [events]}}
    """
    weeks = defaultdict(lambda: defaultdict(list))
    for event in events:
        start = event['start']
        iso_year, week_num, _ = start.isocalendar()
        weeks[(iso_year, week_num)][start.date()].append(event)
    
    # Sort by week, then sort dates within each week
    return {
        week_key: dict(sorted(dates.items()))
        for week_key, dates in sorted(weeks.items())
    }


def get_week_start(dates: Dict) -> datetime:
    """Get the Monday of the week from a dict of dates."""
    first_date = min(dates.keys())
    first_event = dates[first_date][0]
    return first_event['start'] - timedelta(days=first_event['start'].weekday())


def generate_ascii(events: List[Dict], title: str = "Bastien's Available Time Slots") -> str:
    """Generate ASCII text representation of events."""
    if not events:
        return "No available time slots found.\n"

    lines = [
        "=" * 70,
        title.center(70),
        "=" * 70,
        ""
    ]

    weeks = group_events_by_week(events)

    for week_idx, (week_key, dates) in enumerate(weeks.items()):
        if week_idx > 0:
            lines.append("")

        week_start = get_week_start(dates)
        lines.append(f"WEEK OF {week_start.strftime('%B %d, %Y').upper()}")
        lines.append("-" * 70)

        for date_key, day_events in dates.items():
            date_str = day_events[0]['start'].strftime('%A, %B %d')
            lines.append(f"\n  {date_str}")

            for event in day_events:
                start, end = event['start'], event['end']
                duration = format_duration(start, end)
                lines.append(f"    {start.strftime('%H:%M')} - {end.strftime('%H:%M')} ({duration})")

    generated_at = datetime.now().strftime('%Y-%m-%d at %H:%M')
    lines.extend(["", "=" * 70, f"Generated on {generated_at}", ""])

    return "\n".join(lines)


def generate_html(events: List[Dict], title: str = "Bastien's Available Time Slots") -> str:
    """Generate HTML page from events and return as string."""
    generated_at = datetime.now().strftime('%Y-%m-%d at %H:%M')

    # Build events HTML
    if not events:
        events_html = '''
        <div class="no-events">
            <p>No available time slots found.</p>
        </div>'''
    else:
        weeks = group_events_by_week(events)
        events_parts = ['        <div class="events">']

        for week_key, dates in weeks.items():
            week_start = get_week_start(dates)
            events_parts.append(f'''
            <div class="week-group">
                <div class="week-header">Week of {week_start.strftime('%B %d, %Y')}</div>''')

            for date_key, day_events in dates.items():
                date_str = day_events[0]['start'].strftime('%A, %B %d')
                events_parts.append(f'''
                <div class="event">
                    <div class="event-date">{date_str}</div>''')

                for event in day_events:
                    start, end = event['start'], event['end']
                    duration = format_duration(start, end)
                    time_str = f"{start.strftime('%H:%M')} - {end.strftime('%H:%M')}"
                    events_parts.append(f'''
                    <div class="event-time">
                        <span class="time-range">{time_str}</span>
                        <span class="event-duration">{duration}</span>
                    </div>''')

                events_parts.append('''
                </div>''')

            events_parts.append('''
            </div>''')

        events_parts.append('''
        </div>''')
        events_html = '\n'.join(events_parts)

    return f'''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            line-height: 1.6; color: #333; background: #f5f5f5; padding: 20px;
        }}
        .container {{
            max-width: 900px; margin: 0 auto; background: white;
            border-radius: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); overflow: hidden;
        }}
        header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white; padding: 30px; text-align: center;
        }}
        header h1 {{ font-size: 28px; font-weight: 600; margin-bottom: 8px; }}
        header p {{ opacity: 0.9; font-size: 14px; }}
        header a {{ color: white; }}
        .events {{ padding: 20px; }}
        .week-group {{ margin-bottom: 30px; }}
        .week-header {{
            font-size: 14px; font-weight: 600; color: #666;
            text-transform: uppercase; letter-spacing: 0.5px;
            margin-bottom: 12px; padding-bottom: 8px; border-bottom: 2px solid #e0e0e0;
        }}
        .event {{
            background: white; border-left: 4px solid #667eea; padding: 16px;
            margin-bottom: 12px; border-radius: 6px; transition: all 0.2s; border: 1px solid #e0e0e0;
        }}
        .event:hover {{
            border-left-color: #764ba2; box-shadow: 0 2px 8px rgba(0,0,0,0.1); transform: translateX(4px);
        }}
        .event-date {{
            font-weight: 600; color: #333; font-size: 16px;
            margin-bottom: 12px; padding-bottom: 8px; border-bottom: 1px solid #f0f0f0;
        }}
        .event-time {{
            color: #666; font-size: 14px; padding: 8px 0;
            display: flex; align-items: center; justify-content: space-between;
        }}
        .event-time:not(:last-child) {{ border-bottom: 1px dashed #e8e8e8; }}
        .time-range {{ flex: 1; }}
        .event-duration {{
            display: inline-block; background: #e8eaf6; color: #667eea;
            padding: 2px 8px; border-radius: 4px; font-size: 12px; font-weight: 600; margin-left: 8px;
        }}
        .no-events {{ text-align: center; padding: 60px 20px; color: #999; }}
        footer {{
            text-align: center; padding: 20px; color: #999;
            font-size: 12px; border-top: 1px solid #e0e0e0;
        }}
        @media (max-width: 600px) {{ header h1 {{ font-size: 24px; }} }}
    </style>
</head>
<body>
    <div class="container">
        <header>
            <h1>{title}</h1>
            <p><strong><a href="https://bzg.fr/agenda.ics">iCal file</a></strong></p>
        </header>
{events_html}
        <footer>
            Generated on {generated_at}
        </footer>
    </div>
</body>
</html>
'''


def main():
    if len(sys.argv) < 2:
        print("Usage: python ics_export.py <input.ics> [--format html|ascii] [output_file] [title]", file=sys.stderr)
        print("\nArguments:", file=sys.stderr)
        print("  input.ics       Input ICS file (use '-' for stdin)", file=sys.stderr)
        print("  --format        Output format: html or ascii (default: html)", file=sys.stderr)
        print("  output_file     Output file (default: stdout, use '-' for stdout)", file=sys.stderr)
        print("  title           Title for the output (default: \"Bastien's Available Time Slots\")", file=sys.stderr)
        print("\nExamples:", file=sys.stderr)
        print("  python ics_export.py input.ics                          # HTML to stdout", file=sys.stderr)
        print("  python ics_export.py input.ics --format ascii           # ASCII to stdout", file=sys.stderr)
        print("  python ics_export.py input.ics --format html out.html  # HTML to file", file=sys.stderr)
        print("  python ics_export.py input.ics --format ascii out.txt  # ASCII to file", file=sys.stderr)
        print("  cat input.ics | python ics_export.py - --format ascii  # Pipe stdin to ASCII stdout", file=sys.stderr)
        sys.exit(1)

    # Parse arguments
    input_file = sys.argv[1]
    output_format = "html"
    output_file = None
    title = "Bastien's Available Time Slots"

    i = 2
    while i < len(sys.argv):
        arg = sys.argv[i]
        if arg in ('--format', '-f'):
            if i + 1 >= len(sys.argv):
                print("Error: --format requires an argument", file=sys.stderr)
                sys.exit(1)
            output_format = sys.argv[i + 1].lower()
            if output_format in ('text', 'txt'):
                output_format = 'ascii'
            elif output_format not in ('html', 'ascii'):
                print(f"Error: Invalid format '{output_format}'. Use 'html' or 'ascii'.", file=sys.stderr)
                sys.exit(1)
            i += 2
        elif not arg.startswith('-'):
            if output_file is None:
                output_file = arg if arg != '-' else None
            else:
                title = arg
            i += 1
        else:
            i += 1

    print(f"Reading {input_file}...", file=sys.stderr)
    cal = parse_ics(input_file)

    print("Extracting events...", file=sys.stderr)
    events = get_events(cal)
    print(f"Found {len(events)} events", file=sys.stderr)

    print(f"Generating {output_format.upper()} output...", file=sys.stderr)

    output_content = generate_html(events, title) if output_format == 'html' else generate_ascii(events, title)

    if output_file is None:
        print(output_content)
    else:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(output_content)
        print(f"\nSuccessfully created {output_file}", file=sys.stderr)
        if output_format == 'html':
            print(f"\nOpen in browser: file://{output_file}", file=sys.stderr)


if __name__ == "__main__":
    main()
