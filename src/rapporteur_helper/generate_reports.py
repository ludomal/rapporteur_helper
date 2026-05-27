import logging
from collections.abc import Iterable
from datetime import datetime, timedelta
from pathlib import Path

from docx import Document as open_docx
from docx.document import Document
from docxtpl import DocxTemplate

from rapporteur_helper.content.contacts import get_chair_text, insert_contacts
from rapporteur_helper.content.documents import insert_documents
from rapporteur_helper.data.constants import template_file
from rapporteur_helper.itut.endpoints import get_endpoint
from rapporteur_helper.itut.questions import get_questions_details
from rapporteur_helper.itut.work_programme import get_work_program, insert_work_program
from rapporteur_helper.word_docx.paragraph import find_element

logger = logging.getLogger(__name__)


def main(
    meetingDate: str,
    questions: Iterable[int],
    meeting_place: str = "Geneva",
    meeting_duration_days: int = 9,
    studyGroup: int = 12,
    studyPeriodId: int = 18,
    studyPeriodStart: int = 25,
    add_qall: bool = True,
    output_dir: Path | None = None,
    verbose: bool = False,
):
    # parse/check parameters
    output_dir = Path.cwd() if output_dir is None else output_dir
    output_dir /= meetingDate
    output_dir.mkdir(parents=True, exist_ok=True)

    # Meeting details
    md_start = datetime.strptime("20" + meetingDate, "%Y%m%d")  # validate format
    md_end = md_start + timedelta(days=meeting_duration_days)
    year = md_start.strftime("%Y")
    month = md_start.strftime("%B")
    day1 = md_start.strftime("%d")
    day2 = md_end.strftime("%d")

    if md_start.month != md_end.month:
        month2 = md_end.strftime("%B")
        meetingDetails = f"{meeting_place}, {day1} {month} - {day2} {month2} {year}"
    else:
        meetingDetails = f"{meeting_place}, {day1} - {day2} {month} {year}"

    try:
        questionInfo = get_questions_details(studyGroup, studyPeriodId)
    except Exception as e:
        raise RuntimeError(f"Error - Cannot fetch question details from ITU-T website: {e}") from e

    for question in questions:
        if question not in questionInfo:
            logger.warning(f"Question {question} not found in questionInfo")
            continue

        logger.info(f"Generating report for Q{question}")
        endpoints_c = []
        endpoints_td = []

        if add_qall:
            endpoints_c.append(get_endpoint(studyGroup, None, studyPeriodStart, meetingDate, "C"))
            endpoints_td.append(get_endpoint(studyGroup, None, studyPeriodStart, meetingDate, "TD"))

        endpoints_c.append(get_endpoint(studyGroup, question, studyPeriodStart, meetingDate, "C"))
        endpoints_td.append(get_endpoint(studyGroup, question, studyPeriodStart, meetingDate, "TD"))
        context_vars = {
            "q": str(question),
            "place_date": meetingDetails,
            "sg": str(studyGroup),
            "q_title": questionInfo[question]["title"],
            "wp": questionInfo[question]["wp"],
            "rapporteur_ident": "Co-Rapporteurs" if len(questionInfo[question]["rapporteurs"]) > 1 else "Rapporteur",
            "chairing_persons": get_chair_text(questionInfo[question]["rapporteurs"]),
        }
        abstract = f'This document contains the Status report of Question {question}/{studyGroup}: "{questionInfo[question]["title"]}" for the meeting in {meetingDetails}.'
        context_vars["abstract"] = abstract

        try:
            with template_file.open("rb") as f:
                document: Document = open_docx(f)

            # Meeting date
            # replace(document, "[place, dates]", meetingDetails)

            # Abstract
            # replace(document, "[Insert an abstract]", abstract)

            # Insert contributions
            if docSection := find_element(document, "Copy table of contributions."):
                insert_documents(docSection, endpoints_c, verbose=verbose)

            # Insert temporary documents
            # print("  Inserting temporary documents")
            if docSection := find_element(document, "Copy the TD table"):
                insert_documents(docSection, endpoints_td, verbose=verbose)

            # Replace question number
            # replace(document, f"X/{studyGroup}", f"{question}/{studyGroup}")
            # replace(document, f"t{studyPeriodStart}sg{studyGroup}qX@lists.itu.int", f"t{studyPeriodStart}sg{studyGroup}q{question}@lists.itu.int")

            # Replace working party number
            # replace(document, f"Working Party y/{studyGroup}", f"Working Party {questionInfo[question]['wp']}/{studyGroup}")

            # Replace question title
            # replace(document, "[title of question]", questionInfo[question]["title"])
            # replace(document, "Title of question", questionInfo[question]["title"])

            # Insert contact information of rapporteur(s)
            insert_contacts(document, questionInfo[question])

            # Insert work programme
            workProgram = get_work_program(question)
            # pprint(workProgram)
            insert_work_program(document, workProgram)

            output_file = output_dir / f"Q{question}_status_report.docx"
            document.save(str(output_file))

            # replacements of remaining variables via docxtpl
            doc = DocxTemplate(str(output_file))
            doc.render(context_vars)
            doc.save(str(output_file))
        except Exception:
            logger.exception(f"Error generating report for Q{question}")
            # traceback.print_stack()
            # pprint(questionInfo)


if __name__ == "__main__":
    # Update these parameters for each meeting
    studyGroup = 12
    meetingDate = "250114"  # meetingDate = "220607"
    questions = list(range(1, 21))
    # questions = [1,2, 7, 14]
    meeting_duration_days = 9  # TODO: parse via web API?

    # Update these parameters to the current study period
    # TODO: derive values programmatically from the meeting date
    studyPeriodId = 18
    studyPeriodStart = 25

    # misc. parameters
    meeting_place = "Geneva"  # TODO: parse via web API?
    add_qall = False
    verbose = True

    main(
        meetingDate=meetingDate,
        questions=questions,
        meeting_place=meeting_place,
        meeting_duration_days=meeting_duration_days,
        studyGroup=studyGroup,
        studyPeriodId=studyPeriodId,
        studyPeriodStart=studyPeriodStart,
        add_qall=add_qall,
        verbose=verbose,
    )

