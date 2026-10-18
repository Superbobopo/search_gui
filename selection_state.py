"""The displayed occurrence is owned by the application, never by a module import."""

from dataclasses import dataclass


@dataclass(frozen=True)
class OccurrenceSelection:
    path: str
    number: int
    reference: str | None
    modbus: str | None
    details: dict

    @classmethod
    def from_details(cls, details):
        row = details['row']
        return cls(details['path'], row['number'], row['ref'], row['modbus'], details)
