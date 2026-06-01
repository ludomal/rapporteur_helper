"""Fetch 'Meeting in Focus' details from the ITU-T Study Group page."""

import re
from dataclasses import dataclass
from datetime import date

from ..data.constants import hostname
from ..html import get_html_tree

# Months lookup for parsing
MONTHS = {
    "January": 1,
    "February": 2,
    "March": 3,
    "April": 4,
    "May": 5,
    "June": 6,
    "July": 7,
    "August": 8,
    "September": 9,
    "October": 10,
    "November": 11,
    "December": 12,
}


@dataclass
class MeetingInfo:
    place: str
    start_date: date
    end_date: date

    @property
    def meeting_date(self) -> str:
        """YYMMDD format for the start date."""
        return self.start_date.strftime("%y%m%d")

    @property
    def meeting_details(self) -> str:
        """Formatted string like 'Geneva, 9-17 June 2026'."""
        day1 = self.start_date.day
        day2 = self.end_date.day
        if self.start_date.month == self.end_date.month:
            month = self.start_date.strftime("%B")
            return f"{self.place}, {day1}-{day2} {month} {self.start_date.year}"
        else:
            month1 = self.start_date.strftime("%B")
            month2 = self.end_date.strftime("%B")
            return f"{self.place}, {day1} {month1} - {day2} {month2} {self.start_date.year}"


def get_sg_page_url(study_group: int, study_period: str = "2025-2028") -> str:
    return f"{hostname}/en/ITU-T/studygroups/{study_period}/{study_group}/Pages/default.aspx"


def fetch_meeting_info(study_group: int = 12, study_period: str = "2025-2028") -> MeetingInfo:
    """Scrape the 'Meeting in Focus' section from the ITU-T SG page.

    Returns a MeetingInfo with place, start_date, and end_date.
    """
    url = get_sg_page_url(study_group, study_period)
    tree = get_html_tree(url)

    # The meeting info is in a bold element after the "Meeting in Focus" heading
    # Look for bold/strong text matching the pattern "Place, D-D Month YYYY"
    bold_texts = tree.xpath("//strong/text() | //b/text()")

    # Pattern: "Geneva, 9-17 June 2026" or "Geneva, 28 November - 6 December 2025"
    # Match hyphen or en-dash as separator
    dash = r"[-\u2013]"
    same_month = re.compile(rf"^(.+),\s+(\d{{1,2}})\s*{dash}\s*(\d{{1,2}})\s+(\w+)\s+(\d{{4}})$")
    cross_month = re.compile(rf"^(.+),\s+(\d{{1,2}})\s+(\w+)\s*{dash}\s*(\d{{1,2}})\s+(\w+)\s+(\d{{4}})$")

    for text in bold_texts:
        text = text.strip()
        if m := same_month.match(text):
            place, day1, day2, month, year = m.groups()
            return MeetingInfo(
                place=place.strip(),
                start_date=date(int(year), MONTHS[month], int(day1)),
                end_date=date(int(year), MONTHS[month], int(day2)),
            )
        if m := cross_month.match(text):
            place, day1, month1, day2, month2, year = m.groups()
            return MeetingInfo(
                place=place.strip(),
                start_date=date(int(year), MONTHS[month1], int(day1)),
                end_date=date(int(year), MONTHS[month2], int(day2)),
            )

    raise ValueError(f"Could not find 'Meeting in Focus' details on {url}")
