"""Error type shared by the styled report builder."""
from __future__ import annotations


class ReportError(ValueError):
    """A report spec, input file, or output target is not usable safely."""
