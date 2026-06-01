from logging import getLogger

import requests
from lxml import html

logger = getLogger("html")


def get_html_tree(url):
    try:
        x = requests.get(url, timeout=30)
        return html.fromstring(x.content)
    except Exception:
        logger.exception(f"Error fetching {url}")
        raise


if __name__ == "__main__":
    pass
