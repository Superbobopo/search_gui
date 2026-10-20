"""A saved trace is identified by its source path AND its generation."""

from dataclasses import dataclass, field
import sqlite3


@dataclass(frozen=True)
class TraceRef:
    path: str = field(compare=False)
    version_id: str
    start: float = field(default=0, compare=False)
    end: float = field(default=0, compare=False)
    historical: bool = field(default=False, compare=False)

    @property
    def key(self):
        return 'omega:' + self.version_id

    def __fspath__(self):
        return self.path

    def __str__(self):
        return self.path


sqlite3.register_adapter(TraceRef, lambda ref: ref.key)
