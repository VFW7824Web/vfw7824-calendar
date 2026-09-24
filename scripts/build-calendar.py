"""
VFW Post 7824 - Microsoft 365 Calendar Builder
Version: 1.0.0

Purpose:
- Download the published Microsoft 365 ICS calendar.
- Save a public copy as calendar.ics.
- Expand recurring events.
- Generate calendar-data.json for the public website.

The Microsoft ICS URL is supplied through the GitHub Actions secret:
M365_CALENDAR_ICS_URL
"""

import hashlib
import json
import os
import sys
import urllib.request
from datetime import date, datetime, time, timedelta, timezone
from pathlib import Path
from zoneinfo import ZoneInfo

from icalendar import Calendar
import recurring_ical_events


TIME_ZONE_NAME = "America/Los_Angeles"
LOCAL_TIME_ZONE = ZoneInfo(TIME_ZONE_NAME)

ROOT = Path(__file__).resolve().parent.parent
ICS_OUTPUT = ROOT / "calendar.ics"
JSON_OUTPUT = ROOT / "calendar-data.json"


def download_calendar(url):
    """Download the Microsoft 365 published ICS calendar."""
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "VFW-Post-7824-Calendar/1.0"
        },
    )

    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            data = response.read()
    except Exception as exc:
        raise RuntimeError(
            f"Unable to download Microsoft 365 calendar: {exc}"
        ) from exc

    if not data:
        raise RuntimeError("Microsoft 365 calendar returned no data.")

    if b"BEGIN:VCALENDAR" not in data:
        raise RuntimeError(
            "Downloaded data does not appear to be an ICS calendar."
        )

    return data


def text_value(component, name):
    """Safely convert an iCalendar field to plain text."""
    value = component.get(name)

    if value is None:
        return ""

    try:
        return str(value)
    except Exception:
        return ""


def get_categories(component):
    """Return Outlook/iCalendar categories as a normal Python list."""
    category_property = component.get("categories")

    if category_property is None:
        return []

    try:
        if hasattr(category_property, "cats"):
            return [
                str(item).strip()
                for item in category_property.cats
                if str(item).strip()
            ]
    except Exception:
        pass

    try:
        raw = category_property.to_ical().decode("utf-8")
    except Exception:
        raw = str(category_property)

    categories = []

    for value in raw.split(","):
        value = value.strip()

        if value:
            categories.append(value)

    return categories


def determine_source(categories):
    """
    Map Outlook categories to the four VFW calendar sources.

    Recommended Outlook categories:
    Post 7824
    District 6
    Department
    National
    """
    normalized = {
        item.strip().lower()
        for item in categories
    }

    if (
        "national" in normalized
        or "vfw national" in normalized
    ):
        return "VFW National"

    if (
        "department" in normalized
        or "department of washington" in normalized
        or "washington department" in normalized
    ):
        return "Department of Washington"

    if (
        "district" in normalized
        or "district 6" in normalized
        or "vfw district 6" in normalized
    ):
        return "District 6"

    return "Post 7824"


def normalize_datetime(value):
    """
    Convert an iCalendar date/datetime into Pacific time.

    Returns:
      ISO formatted date/time
      True/False indicating whether this is an all-day value
    """
    if isinstance(value, datetime):
        if value.tzinfo is None:
            value = value.replace(tzinfo=LOCAL_TIME_ZONE)
        else:
            value = value.astimezone(LOCAL_TIME_ZONE)

        return value.isoformat(), False

    if isinstance(value, date):
        value_datetime = datetime.combine(
            value,
            time.min,
            tzinfo=LOCAL_TIME_ZONE,
        )

        return value_datetime.isoformat(), True

    return "", False


def event_sort_datetime(event):
    """Provide a safe datetime used to sort generated events."""
    try:
        return datetime.fromisoformat(event["start"])
    except Exception:
        return datetime.max.replace(tzinfo=timezone.utc)


def make_event_id(uid, start):
    """Create a stable unique identifier for each occurrence."""
    raw = f"{uid}|{start}".encode("utf-8")

    return hashlib.sha256(raw).hexdigest()[:20]


def process_event(component):
    """Convert one expanded VEVENT component into website JSON."""
    dtstart_property = component.get("dtstart")

    if dtstart_property is None:
        return None

    start_value = component.decoded("dtstart")
    start_iso, all_day = normalize_datetime(start_value)

    if not start_iso:
        return None

    dtend_property = component.get("dtend")

    if dtend_property is not None:
        end_value = component.decoded("dtend")
        end_iso, _ = normalize_datetime(end_value)
    else:
        if all_day:
            end_datetime = (
                datetime.fromisoformat(start_iso)
                + timedelta(days=1)
            )
        else:
            end_datetime = (
                datetime.fromisoformat(start_iso)
                + timedelta(hours=1)
            )

        end_iso = end_datetime.isoformat()

    title = text_value(component, "summary").strip()

    if not title:
        title = "VFW Post 7824 Event"

    location = text_value(component, "location").strip()
    description = text_value(component, "description").strip()
    url = text_value(component, "url").strip()
    uid = text_value(component, "uid").strip()

    categories = get_categories(component)
    source = determine_source(categories)

    event_id = make_event_id(
        uid or title,
        start_iso,
    )

    return {
        "id": event_id,
        "uid": uid,
        "title": title,
        "start": start_iso,
        "end": end_iso,
        "allDay": all_day,
        "location": location,
        "description": description,
        "url": url,
        "source": source,
        "categories": categories,
    }


def build_calendar():
    """Main calendar build process."""
    ics_url = os.environ.get(
        "M365_CALENDAR_ICS_URL",
        "",
    ).strip()

    html_url = os.environ.get(
        "M365_CALENDAR_HTML_URL",
        "",
    ).strip()

    if not ics_url:
        raise RuntimeError(
            "M365_CALENDAR_ICS_URL is not configured."
        )

    print("Downloading Microsoft 365 calendar...")

    ics_data = download_calendar(ics_url)

    # Publish a local copy for the website's Subscribe button.
    ICS_OUTPUT.write_bytes(ics_data)

    print("Saved calendar.ics")

    calendar = Calendar.from_ical(ics_data)

    now = datetime.now(LOCAL_TIME_ZONE)

    # Keep enough history for recent-event viewing while supplying
    # approximately one year of upcoming events.
    range_start = now - timedelta(days=180)
    range_end = now + timedelta(days=400)

    print(
        "Expanding events from "
        f"{range_start.date()} through {range_end.date()}..."
    )

    try:
        components = recurring_ical_events.of(
            calendar
        ).between(
            range_start,
            range_end,
        )
    except Exception as exc:
        raise RuntimeError(
            f"Unable to expand recurring calendar events: {exc}"
        ) from exc

    events = []

    for component in components:
        try:
            event = process_event(component)

            if event:
                events.append(event)

        except Exception as exc:
            print(
                f"Warning: skipped an event: {exc}",
                file=sys.stderr,
            )

    events.sort(key=event_sort_datetime)

    sources = sorted(
        {
            event["source"]
            for event in events
            if event["source"]
        }
    )

    categories = sorted(
        {
            category
            for event in events
            for category in event["categories"]
            if category
        },
        key=str.lower,
    )

    output = {
        "version": "1.0.0",
        "generatedAt": datetime.now(
            timezone.utc
        ).isoformat(),
        "timeZone": TIME_ZONE_NAME,
        "calendarName": "VFW Post 7824 Events Calendar",
        "calendarViewUrl": html_url,
        "subscriptionUrl": "calendar.ics",
        "eventCount": len(events),
        "sources": sources,
        "categories": categories,
        "events": events,
    }

    JSON_OUTPUT.write_text(
        json.dumps(
            output,
            indent=2,
            ensure_ascii=False,
        ),
        encoding="utf-8",
    )

    print(
        f"Created calendar-data.json with "
        f"{len(events)} events."
    )

    print("Calendar build completed successfully.")


if __name__ == "__main__":
    try:
        build_calendar()
    except Exception as exc:
        print(
            f"Calendar build failed: {exc}",
            file=sys.stderr,
        )
        sys.exit(1)
