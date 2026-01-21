from dataclasses import dataclass
from typing import Iterable, List, Optional, Sequence, Tuple


Description = Sequence[Tuple[Optional[str], Optional[str], Optional[int], Optional[int], Optional[int], Optional[int], Optional[bool]]]


@dataclass
class ExecutionResult:
    action: str
    rows: List[Tuple]
    description: Description
    rowcount: int
    lastrowid: Optional[int] = None
