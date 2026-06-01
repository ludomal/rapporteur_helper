import contextlib
import logging
import re

from lxml.html import HtmlElement

from ..html import get_html_tree

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class QuestionDetailsParseException(Exception):
    def __init__(self, url):
        super().__init__(f"get_questions_details - Could not parse question details from {url}")


def find_span_in_row(row: HtmlElement, span_id: str) -> str | None:
    matched_item = row.xpath(f".//span[contains(@id,'{span_id}')]/text()")
    return matched_item[0] if matched_item else None


def get_questions_details(studyGroup: int, studyPeriodId: int) -> dict:
    info = {}
    qNum = -1
    url = f"https://www.itu.int/net4/ITU-T/lists/loqr.aspx?Group={studyGroup}&Period={studyPeriodId}"

    tree = get_html_tree(url)

    # Find and parse all rows (<tr>) in the document
    rows = tree.xpath("//tr")

    # Extract WP number and Question title
    for row in rows:
        # Question number and WP
        if len(tmp := row.xpath(".//span[contains(@id,'lblQWP')]/text()")) > 0:
            tmp = tmp[0]
            try:
                res = re.search(rf"Q(\d+)/{studyGroup}.*WP(\d+)/{studyGroup}", tmp)
                qNum = int(res.group(1))
                wpNum = int(res.group(2))
            except Exception as e:
                # If it fails, check that it is because there is no WP number
                res = re.search(rf"Q(\d+)/{studyGroup}.*", tmp)
                qNum = int(res.group(1))
                wpNum = -1
                logger.warning(f"Could not parse WP number for question {qNum}: {e} -> Plenary Question?")

            # try:
            # Question title
            qTitle = row.xpath(".//span[contains(@id,'lblQuestion')]/text()")[0]

            if qNum not in info:
                info[qNum] = {"rapporteurs": []}

            info[qNum].update({"wp": wpNum, "title": qTitle})

            # except Exception as exc:
            #    # This is not a row with a question
            # logger.exception("Exception occurred while parsing question row")

        # Extract Rapporteurs contact details, if question number was parsed out successfully
        if qNum > 0:
            # if lastname cannot be parsed, it's an invalid entry
            lastName = find_span_in_row(row, "dtlRappQues_lblLName")
            if lastName is not None:
                tmp = {}
                tmp["lastName"] = lastName.upper()
                tmp["firstName"] = find_span_in_row(row, "dtlRappQues_lblFName")
                tmp["role"] = find_span_in_row(row, "dtlRappQues_lblRole")
                tmp["company"] = find_span_in_row(row, "dtlRappQues_lblCompany")
                tmp["address"] = " ".join(row.xpath(".//span[contains(@id,'dtlRappQues_lblAddress')]/text()"))
                tmp["country"] = row.xpath(".//span[contains(@id,'dtlRappQues_lblAddress')]/text()")[-1]

                with contextlib.suppress(Exception):
                    # Some Rapporteurs do not have a telephone number available
                    tmp["tel"] = row.xpath(".//span[contains(@id,'dtlRappQues_telLabel')]/text()")[0]
                tmp["email"] = row.xpath(".//a[contains(@id,'dtlRappQues_linkemail')]/text()")[0].replace("[at]", "@")

                info[qNum]["rapporteurs"].append(tmp)

    if len(info) < 1:
        raise QuestionDetailsParseException(url)

    return info


if __name__ == "__main__":
    pass
