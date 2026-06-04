"""
Pre-import filters (date, size, regex, attachments). Enable via conversion_options['email_filter'].
"""

from __future__ import annotations

import re
from datetime import datetime
from email.utils import parsedate_to_datetime
from typing import Any, Callable


class EmailFilter:
    def __init__(self):
        self._filters: list[Callable[[dict[str, Any]], bool]] = []

    def add_date_filter(self, start_date=None, end_date=None) -> None:
        def date_filter(email_data: dict) -> bool:
            date_str = email_data.get("date") or email_data.get("Date") or ""
            if not date_str:
                return True
            try:
                email_date = parsedate_to_datetime(date_str)
                if email_date.tzinfo is not None:
                    email_date = email_date.replace(tzinfo=None)
                if start_date and email_date < start_date:
                    return False
                if end_date and email_date > end_date:
                    return False
                return True
            except (TypeError, ValueError, OverflowError):
                return True

        self._filters.append(date_filter)

    def add_size_filter(self, max_size_bytes: int) -> None:
        def size_filter(email_data: dict) -> bool:
            return int(email_data.get("size") or 0) <= max_size_bytes

        self._filters.append(size_filter)

    def add_regex_filter(self, field: str, pattern: str, *, exclude: bool = False) -> None:
        regex = re.compile(pattern, re.IGNORECASE)

        def regex_filter(email_data: dict) -> bool:
            value = str(email_data.get(field) or "")
            match = bool(regex.search(value))
            return not match if exclude else match

        self._filters.append(regex_filter)

    def add_attachment_filter(
        self, *, has_attachments: bool | None = None, min_attachments: int = 0
    ) -> None:
        def attachment_filter(email_data: dict) -> bool:
            count = int(email_data.get("attachment_count") or len(email_data.get("attachments") or []))
            if has_attachments is True and count == 0:
                return False
            if has_attachments is False and count > 0:
                return False
            return count >= min_attachments

        self._filters.append(attachment_filter)

    def apply(self, email_data: dict[str, Any]) -> bool:
        return all(f(email_data) for f in self._filters)

    @property
    def enabled(self) -> bool:
        return bool(self._filters)
