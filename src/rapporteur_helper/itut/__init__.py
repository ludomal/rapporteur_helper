from typing import Literal

DocType = Literal["TD", "C"]

VALID_DOCTYPES: dict[DocType, str] = {"TD": "Temporary Documents", "C": "Contributions"}

if __name__ == "__main__":
    pass
