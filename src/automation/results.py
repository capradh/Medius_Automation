from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class JobResult:
    """What a job returns: a human-readable summary plus any files it produced."""
    summary: str
    files: list[Path] = field(default_factory=list)
