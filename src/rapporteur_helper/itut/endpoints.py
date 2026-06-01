from typing import Any

from ..data.constants import hostname
from . import VALID_DOCTYPES, DocType


def get_endpoint(studyGroup: int, question: int | str | None, studyPeriodStart: Any, meetingDate: Any, endpoint_type: DocType) -> dict[str, str]:
    question = "ALL" if question is None else question
    prefix = f"SG{studyGroup}-{endpoint_type}"
    url = f"{hostname}/md/meetingdoc.asp?lang=en&parent=T{studyPeriodStart}-SG{studyGroup}-{meetingDate}-{endpoint_type}&question=Q{question}/{studyGroup}"

    return {"url": url, "prefix": prefix, "title": VALID_DOCTYPES[endpoint_type]}


if __name__ == "__main__":
    pass
