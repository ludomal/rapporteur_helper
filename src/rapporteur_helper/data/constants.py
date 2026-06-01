"""
Module for constants, definitions, fixed paths, etc.
"""

from pathlib import Path

data_dir = Path(__file__).parent

root_dir = data_dir.parent
template_file = data_dir / "template.docx"


hostname: str = "https://www.itu.int"


if __name__ == "__main__":
    pass
