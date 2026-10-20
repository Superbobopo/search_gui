"""The displayed occurrence is owned by the application, never by a module import."""

from dataclasses import dataclass


@dataclass(frozen=True)
class OccurrenceSelection:
    path: str
    number: int
    reference: str | None
    modbus: str | None
    details: dict

    @property
    def version_id(self):
        return getattr(self.path, 'version_id', None)

    @classmethod
    def from_details(cls, details):
        row = details['row']
        return cls(details['path'], row['number'], row['ref'], row['modbus'], details)
