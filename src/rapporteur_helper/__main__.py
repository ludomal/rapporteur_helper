import logging
from datetime import date
from pathlib import Path

import click

from rapporteur_helper.generate_reports import main
from rapporteur_helper.itut.meeting import MeetingInfo, fetch_meeting_info


@click.command()
@click.option("--questions", "-q", default="1-20", help="Question numbers: range '1-20', list '1,2,7', or single '5' (default: 1-20)")
@click.option("--study-group", "-s", default=12, type=int, help="ITU-T Study Group number (default: 12)")
@click.option("--study-period-id", default=18, type=int, help="Study period ID (default: 18)")
@click.option("--study-period-start", default=25, type=int, help="Study period start year, 2-digit (default: 25)")
@click.option("--meeting-date", "-d", default=None, help="Override meeting start date (YYMMDD). If omitted, fetched from ITU-T website.")
@click.option("--meeting-place", "-p", default=None, help="Override meeting location. If omitted, fetched from ITU-T website.")
@click.option("--meeting-end-date", default=None, help="Override meeting end date (YYMMDD). If omitted, fetched from ITU-T website.")
@click.option("--add-qall/--no-add-qall", default=False, help="Include documents for QALL in each report (default: False)")
@click.option("--output-dir", "-o", type=click.Path(path_type=Path), help="Output directory (default: current directory)")
@click.option("--verbose/--no-verbose", "-v", default=False, help="Enable verbose output (default: False)")
def cli(
    questions: str,
    study_group: int,
    study_period_id: int,
    study_period_start: int,
    meeting_date: str | None,
    meeting_place: str | None,
    meeting_end_date: str | None,
    add_qall: bool,
    output_dir: Path,
    verbose: bool,
):
    """ITU-T Rapporteur's status report generator.

    Automatically fetches meeting details (place, dates) from the ITU-T Study Group page.
    Use --meeting-date, --meeting-place, --meeting-end-date to override.
    """
    log_level = logging.INFO if verbose else logging.WARNING
    logging.basicConfig(level=log_level, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    question_list = parse_questions(questions)

    # Build MeetingInfo: auto-fetch or use overrides
    meeting_info: MeetingInfo | None = None
    if meeting_date and meeting_place and meeting_end_date:
        # All overrides provided — build manually
        start = _parse_yymmdd(meeting_date)
        end = _parse_yymmdd(meeting_end_date)
        meeting_info = MeetingInfo(place=meeting_place, start_date=start, end_date=end)
    elif meeting_date or meeting_place or meeting_end_date:
        # Partial overrides — fetch then override individual fields
        meeting_info = fetch_meeting_info(study_group)
        if meeting_date:
            meeting_info.start_date = _parse_yymmdd(meeting_date)
        if meeting_end_date:
            meeting_info.end_date = _parse_yymmdd(meeting_end_date)
        if meeting_place:
            meeting_info.place = meeting_place
    # else: None → generate_reports.main() will auto-fetch

    if verbose:
        info = meeting_info or fetch_meeting_info(study_group)
        click.echo(f"Meeting: {info.meeting_details}")
        click.echo(f"Questions: {question_list}")
        click.echo(f"Study Group: {study_group}")
        if meeting_info is None:
            meeting_info = info  # reuse the fetched info

    main(
        questions=question_list,
        meeting_info=meeting_info,
        studyGroup=study_group,
        studyPeriodId=study_period_id,
        studyPeriodStart=study_period_start,
        add_qall=add_qall,
        output_dir=output_dir,
        verbose=verbose,
    )


def _parse_yymmdd(s: str) -> date:
    """Parse a YYMMDD string into a date."""
    return date(2000 + int(s[:2]), int(s[2:4]), int(s[4:6]))


def parse_questions(questions: str) -> list[int]:
    """Parse questions parameter into a list of integers.

    Supports: range '1-20', comma-separated '1,2,7,14', or single '5'.
    """
    if "-" in questions and "," not in questions:
        start, end = questions.split("-")
        return list(range(int(start), int(end) + 1))
    elif "," in questions:
        return [int(q.strip()) for q in questions.split(",")]
    else:
        return [int(questions)]


if __name__ == "__main__":
    cli()
