"""
JIRA Exporter — PowerPoint template filler.

Provides:
  - PPTExporter: Transform tickets and fill PPT template
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from src.definition import JiraTicket

# Output filenames
_PPTX_FILENAME = "jira_export_pptx.pptx"
_PPT_JSON_FILENAME = "jira_export_pptx.json"


class PPTExporter:
    """Transform JiraTicket objects and fill PPT template."""

    def __init__(self, template_path: Path, output_dir: Path):
        self.template_path = Path(template_path)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def transform_for_ppt(self, tickets: list[JiraTicket]) -> Path:
        """Transform JiraTicket objects into a PPT-friendly JSON format.

        Retains only: key, summary, module (from labels[0]),
        attachments, and custom fields 10370/10432 with media_alts
        resolved to attachment URLs.
        """
        result: list[dict[str, Any]] = []

        for ticket in tickets:
            # Build filename → URL lookup from attachments
            att_lookup: dict[str, str] = {
                att.filename: att.url for att in ticket.attachments
            }

            def resolve_media_alts(field_id: str) -> list[str] | None:
                cf = ticket.custom_fields.get(field_id)
                if not cf:
                    return None
                val = cf.get("value")
                if not isinstance(val, dict):
                    return None
                alts = val.get("_media_alts", [])
                if not alts:
                    return None
                return [att_lookup.get(alt, alt) for alt in alts]

            entry: dict[str, Any] = {
                "key": ticket.key,
                "summary": ticket.summary,
                "module": ticket.labels[0] if ticket.labels else None,
                "attachments": [
                    {
                        "filename": att.filename,
                        "mime_type": att.mime_type,
                        "url": att.url,
                        "size": att.size,
                        "created": att.created,
                        "author": att.author,
                    }
                    for att in ticket.attachments
                ],
                "customfield_10370": resolve_media_alts("customfield_10370"),
                "Solution": resolve_media_alts("customfield_10370"),
                "customfield_10432": resolve_media_alts("customfield_10432"),
                "Diff": resolve_media_alts("customfield_10432"),
            }
            result.append(entry)

        with open(self.output_dir / _PPT_JSON_FILENAME, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
        print(f"PPT JSON written to: {(self.output_dir / _PPT_JSON_FILENAME).resolve()}")
        return self.output_dir / _PPT_JSON_FILENAME

    def fill(
        self,
        tickets: list[dict],
        attachments_map: dict[str, dict[str, str]],
        client: Any,
    ) -> Path:
        """Fill the PPT template with ticket data. Returns the output path."""
        # TODO: Implement PPT template filling
        return self.output_dir / _PPTX_FILENAME
