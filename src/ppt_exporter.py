"""JIRA Exporter — PowerPoint template filler.

Provides:
  - PPTExporter: Transform tickets and fill PPT template
"""

from __future__ import annotations

import json
import math
import shutil
from pathlib import Path
from typing import Any

import win32com.client

from src.definition import JiraClient, JiraTicket

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
                "Self Test Report": resolve_media_alts("customfield_10432"),
            }
            result.append(entry)

        with open(self.output_dir / _PPT_JSON_FILENAME, "w", encoding="utf-8") as f:
            json.dump(result, f, indent=2, ensure_ascii=False)
        print(f"PPT JSON written to: {(self.output_dir / _PPT_JSON_FILENAME).resolve()}")
        return self.output_dir / _PPT_JSON_FILENAME

    def fill(self, tickets: list[dict], client: JiraClient) -> Path:
        """Fill the PPT template with ticket data. Returns the output path."""
        output_path = self.output_dir / _PPTX_FILENAME
        shutil.copy2(self.template_path, output_path)

        ppt_app = win32com.client.Dispatch("PowerPoint.Application")
        ppt_app.Visible = True
        pres = ppt_app.Presentations.Open(str(output_path.resolve()), WithWindow=False)

        try:
            self._fill_overview_table(pres, tickets)
            self._fill_ticket_slides(pres, tickets, client)
            pres.Save()
            pres.Close()
        except Exception as e:
            print(f"  [ERROR] {e}")
            try:
                pres.Close()
            except:
                pass
            raise
        finally:
            ppt_app.Quit()

        print(f"PPT written to: {output_path.resolve()}")
        return output_path

    # ── Slide 1: Overview table ────────────────────────────────────────────

    def _fill_overview_table(self, pres: Any, tickets: list[dict]) -> None:
        """Group tickets by module and fill the Slide-1 table."""
        slide = pres.Slides(1)
        table = None
        for shape in slide.Shapes:
            if shape.HasTable:
                table = shape.Table
                break
        if table is None:
            print("  [WARN] No table found on Slide 1")
            return

        # Group keys by module
        module_keys: dict[str, list[str]] = {}
        for t in tickets:
            mod = t.get("module") or "Other"
            module_keys.setdefault(mod, []).append(t["key"])

        # The template table has 1 header row; add rows for each module
        for i, (mod, keys) in enumerate(module_keys.items()):
            row = table.Rows.Add()
            row.Cells(1).Shape.TextFrame.TextRange.Text = mod
            row.Cells(2).Shape.TextFrame.TextRange.Text = ", ".join(keys)

    # ── Slides 2+3: Per-ticket slides ──────────────────────────────────────

    def _fill_ticket_slides(
        self, pres: Any, tickets: list[dict], client: JiraClient
    ) -> None:
        """Fill Solution (slide 2) and Self Test Report (slide 3) for each ticket."""
        # Hardcoded placeholder dimensions from Slide 2 template
        # Based on actual template analysis: L=379, T=436, W=1196, H=380
        ph_left, ph_top, ph_width, ph_height = 379, 436, 1196, 380

        # Duplicate slides 2 & 3 (template pair) for all tickets beyond the first.
        # Must be done BEFORE any text filling to avoid copying filled content.
        for _ in range(len(tickets) - 1):
            self._duplicate_slide_pair(pres, 2, 3)

        # Now fill each ticket's pair of slides
        for i, ticket in enumerate(tickets):
            key = ticket["key"]
            summary = ticket["summary"]
            module = ticket.get("module") or ""
            title_text = f"{key} {summary}"

            if i == 0:
                sol_slide = pres.Slides(2)
                str_slide = pres.Slides(3)
            else:
                # Slide pairs are in order: [1][2,3][4,5][6,7]...
                # Ticket i uses slides at 1-based positions: 2+2*i and 3+2*i
                sol_slide = pres.Slides(2 + 2 * i)
                str_slide = pres.Slides(3 + 2 * i)

            # Fill text placeholders
            self._replace_text_in_slide(sol_slide, title_text, module)
            self._replace_text_in_slide(str_slide, title_text, module)

            # Download and insert images (passing explicit placeholder dimensions)
            sol_urls = ticket.get("Solution") or []
            str_urls = ticket.get("Self Test Report") or []

            self._insert_images_on_slide(sol_slide, sol_urls, client, ph_left, ph_top, ph_width, ph_height)
            self._insert_images_on_slide(str_slide, str_urls, client, ph_left, ph_top, ph_width, ph_height)

    def _duplicate_slide_pair(self, pres: Any, idx1: int, idx2: int) -> None:
        """Duplicate two template slides (1-based indices) for a new ticket."""
        for idx in (idx1, idx2):
            slide = pres.Slides(idx)
            slide.Copy()
            pres.Slides.Paste()

    def _replace_text_in_slide(
        self, slide: Any, title_text: str, module: str
    ) -> None:
        """Replace [key] [summary] and [Module] placeholders in slide shapes."""
        for shape in slide.Shapes:
            if not shape.HasTextFrame:
                continue
            tf = shape.TextFrame
            tr = tf.TextRange
            old = tr.Text
            if "[key] [summary]" in old:
                tr.Text = old.replace("[key] [summary]", title_text)
            elif "[Module]" in old:
                tr.Text = old.replace("[Module]", module)

    # ── Image insertion helpers ────────────────────────────────────────────

    def _insert_images_on_slide(
        self, slide: Any, urls: list[str], client: JiraClient,
        ph_left: float, ph_top: float, ph_width: float, ph_height: float
    ) -> None:
        """Download images and insert them in a grid on the slide.

        Replaces any existing picture placeholder shape.
        """
        if not urls:
            return

        # Download images to temp files
        local_paths: list[str] = []
        for url in urls:
            ext = ".png"
            path = client.download(url, suffix=ext)
            if path:
                local_paths.append(path)
        if not local_paths:
            return

        # Find and remove existing picture placeholder (if any)
        placeholder = self._find_picture_placeholder(slide)
        if placeholder is not None:
            placeholder.Delete()

        n = len(local_paths)
        cols = min(n, 2) if n <= 2 else math.ceil(math.sqrt(n))
        rows = math.ceil(n / cols)

        # Calculate cell sizes with proportional spacing
        gap_x = ph_width * 0.15 if n == 2 else 10  # 15% width gap for 2 images
        gap_y = 10
        cell_w = (ph_width - gap_x * (cols - 1)) / cols
        cell_h = (ph_height - gap_y * (rows - 1)) / rows

        for idx, img_path in enumerate(local_paths):
            r, c = divmod(idx, cols)
            x = ph_left + c * (cell_w + gap_x)
            y = ph_top + r * (cell_h + gap_y)

            # Insert at original size, then constrain to cell while keeping aspect ratio
            pic = slide.Shapes.AddPicture(
                img_path,
                LinkToFile=False,
                SaveWithDocument=True,
                Left=x,
                Top=y,
                Width=-1,
                Height=-1,
            )
            pic.LockAspectRatio = True
            # Shrink to fit within the cell, preserving aspect ratio
            if pic.Width / cell_w > pic.Height / cell_h:
                pic.Width = cell_w
            else:
                pic.Height = cell_h
            # Center within cell
            pic.Left = x + (cell_w - pic.Width) / 2
            pic.Top = y + (cell_h - pic.Height) / 2

    @staticmethod
    def _find_picture_placeholder(slide: Any) -> Any | None:
        """Return the first picture shape on the slide, or None."""
        for shape in slide.Shapes:
            # Type 13 = msoPicture
            if shape.Type == 13:
                return shape
        return None
