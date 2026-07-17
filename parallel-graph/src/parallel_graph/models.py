"""Shared immutable data models."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
from pathlib import Path


@dataclass(frozen=True)
class SourceSpec:
    path: Path
    name: str


@dataclass(frozen=True)
class Registration:
    timestamp: datetime
    source: str
    species: str


@dataclass(frozen=True)
class MinuteCount:
    minute: datetime
    source: str
    species: str
    count: int


@dataclass(frozen=True)
class ImportReport:
    source: str
    path: Path
    sheet: str
    rows_examined: int
    registrations_loaded: int
    blank_species_skipped: int
    noise_skipped: int
    invalid_time_skipped: int

