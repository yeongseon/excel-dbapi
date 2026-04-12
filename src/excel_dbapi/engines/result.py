from dataclasses import dataclass
from typing import Any, List, Optional, Sequence, Tuple


Description = Sequence[
    Tuple[
        Optional[str],
        Optional[str],
        Optional[int],
        Optional[int],
        Optional[int],
        Optional[int],
        Optional[bool],
    ]
]


@dataclass
class ExecutionResult:
    action: str
    rows: List[Tuple[Any, ...]]
    description: Description
    rowcount: int
    lastrowid: Optional[int] = None
