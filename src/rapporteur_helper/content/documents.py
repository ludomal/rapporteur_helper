import logging
import re
from typing import Any

import requests
from docx.text.paragraph import Paragraph
from lxml import html

from ..data.constants import hostname
from ..word_docx.links import add_hyperlink

logger = logging.getLogger(__name__)


def insert_documents(docSection: Paragraph, endpoints: list | Any, verbose: bool = False, studyGroup: int = 12):
    if not isinstance(endpoints, list):
        endpoints = [endpoints]

    rows = []
    for endpoint in endpoints:
        if verbose:
            logger.info(f"Retrieving documents from: {endpoint['url']}")
        x = requests.get(endpoint["url"], timeout=30)
        tree = html.fromstring(x.content)

        # Find and parse all rows (<tr>) in the document
        rows += tree.xpath("//tr")

    docSection.text = ""

    # Parse in descending order
    for i in range(len(rows) - 1, 0, -1):
        row = rows[i]
        columns = row.xpath(".//td")

        try:
            # A document row should have the attributes below
            # If not, then the row is ignored
            if len(columns) < 4:
                continue

            # Link and document number should be in the second column
            if (len(href := columns[1].xpath(".//a")) > 0) and (a := href[0].attrib.get("href")):
                link = hostname + "/" + a.strip()
            else:
                continue

            text = columns[1].xpath(".//a/strong/text()")
            if not text:
                # Handle case where documents have not yet been uploaded -> no strong text
                text = columns[1].xpath(".//a/text()")
                if not text:
                    continue

            number = text[0].strip().replace("[ ", endpoint["prefix"]).replace(" ]", "")

            if len(revision := columns[1].xpath(".//font/text()")) > 0 and (x := re.search(r"([\d]+)\)", revision[0])):
                revision = x.group(1)
                number = f"{number}r{revision}"

            # Title should be in third row
            title = columns[2].xpath(".//text()")[0].strip()
            sources = columns[3].xpath(".//a")
            src = []
            for source in sources:
                src.append({"link": f"{hostname}/{source.attrib['href']}", "text": source.text.strip()})

            # Relevant questions should be in fourth column
            questions = columns[4].xpath(".//a")
            q = []
            for quest in questions:
                q.append({"link": f"{hostname}/{quest.attrib['href']}", "text": quest.text.strip().replace(f"/{studyGroup}", "")})

            # Generate word document block for this document
            # p = document.add_paragraph()
            p = docSection.insert_paragraph_before()

            tmpNumber = number.replace("-GEN", "")
            add_hyperlink(p, f"{tmpNumber} - {title}", link, format="bold")

            p.add_run("\nSources: ")
            for item in src:
                add_hyperlink(p, item["text"], item["link"], format="hyperlink")
                if src[-1] != item:
                    p.add_run(" | ")

            p.add_run("\nQuestions: ")
            for item in q:
                add_hyperlink(p, item["text"], item["link"], format="hyperlink")
                if q[-1] != item:
                    p.add_run(", ")

            # Do not include a summary section for documents addressed to Q.ALL
            if q[0]["link"].find("QALL") < 0:
                p.add_run("\nSummary:\n")

            p.add_run("\n")

        except Exception as e:
            logger.exception(f"Exception occurred while processing document row: {e}")


if __name__ == "__main__":
    pass
