from io import BytesIO
from pathlib import Path
from typing import Union

import requests

from .exceptions import OperationalError  # ConnectionError -> OperationalError


def fetch_remote_file(url: Union[str, Path]) -> BytesIO:
    """Fetch a remote file and return it as BytesIO."""
    try:
        response = requests.get(str(url), timeout=10)
        response.raise_for_status()
        return BytesIO(response.content)
    except requests.RequestException as e:
        raise OperationalError(f"Failed to fetch remote file {url}: {e}")
