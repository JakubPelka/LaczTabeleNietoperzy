"""Minute-resolution aggregation."""

from __future__ import annotations

from collections import Counter
from typing import Iterable

from .models import MinuteCount, Registration


def aggregate_by_minute(registrations: Iterable[Registration]) -> list[MinuteCount]:
    counts = Counter(
        (item.timestamp.replace(second=0, microsecond=0), item.source, item.species)
        for item in registrations
    )
    return [
        MinuteCount(minute, source, species, count)
        for (minute, source, species), count in sorted(counts.items())
    ]

