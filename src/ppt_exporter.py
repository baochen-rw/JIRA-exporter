"""
Generate a PowerPoint report from JIRA ticket data using a template.
Uses win32com.client to drive PowerPoint via COM automation.
"""
from __future__ import annotations

import math
import shutil
import tempfile
from pathlib import Path
from typing import Any

import win32com.client as win32com
import pythoncom

from definition import JiraAttachment


TEMPLATE_PATH = Path(__file__).parent.parent / "PPTTemplate" / "Template.pptx"


# ── Helpers ──────────────────────────────────────────────────────────────────

def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
    return int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)


def _apply_color(font: Any, hex_color: str) -> None:
    r, g, b = _hex_to_rgb(hex_color)
    font.Color.RGB = win32com.RGB(r, g, b)


def _set_shape_text(shape: Any, text: str, color: str | None = None) -> None:
    tf = shape.TextFrame
    tr = tf.TextRange
    if not tr.Text:
        tr.Text = text
        if color:
            _apply_color(tr.Font, color)
        return
    first = tr.Characters(1, 1).Font
    font_name, font_size, font_bold, font_italic = first.Name, first.Size, first.Bold, first.Italic
    tr.Text = text
    new_font = tr.Font
    new_font.Name, new_font.Size, new_font.Bold, new_font.Italic = font_name, font_size, font_bold, font_italic
    if color:
        _apply_color(new_font, color)


def _set_shape_text_mixed(shape: Any, segments: list[tuple[str, str | None]]) -> None:
    tf = shape.TextFrame
    tr = tf.TextRange
    first = tr.Characters(1, 1).Font
    font_name, font_size, font_bold, font_italic = first.Name, first.Size, first.Bold, first.Italic
    tr.Text = "".join(t for t, _ in segments)
    tr.Font.Name, tr.Font.Size, tr.Font.Bold, tr.Font.Italic = font_name, font_size, font_bold, font_italic
    pos = 1
    for text, color in segments:
        if color:
            _apply_color(tr.Characters(pos, len(text)).Font, color)
        pos += len(text)


def _set_cell_text(cell: Any, text: str, color: str | None = None) -> None:
    tr = cell.Shape.TextFrame.TextRange
    tr.Text = text
    if color:
        _apply_color(tr.Font, color)


# ── Core class ───────────────────────────────────────────────────────────────

class PPTTemplateFiller:
    """Fills a PPTX template with JIRA ticket data via PowerPoint COM."""

    def __init__(self, template_path: Path | str | None = None):
        self.template_path = Path(template_path or TEMPLATE_PATH)
        self._download_cache: dict[str, str] = {}
        self._session: Any = None

    # -- public entry point -----------------------------------------------

    def fill(
        self,
        tickets: list[dict[str, Any]],
        output_path: Path | str | None = None,
        attachments_map: dict[str, list[JiraAttachment]] | None = None,
        session: Any = None,
    ) -> Path:
        if attachments_map is None:
            attachments_map = {}
        self._download_cache = {}
        self._session = session

        if not self.template_path.exists():
            raise FileNotFoundError(
                f"PPTX template not found: {self.template_path}. "
                "Please create a PPTTemplate/Template.pptx file or update "
                "ppt_template in config.json."
            )

        tmp_dir = Path(tempfile.mkdtemp())
        tmp_copy = tmp_dir / "Template.pptx"
        shutil.copy(self.template_path, tmp_copy)

        ppt_app = win32com.Dispatch("PowerPoint.Application")
        prs = ppt_app.Presentations.Open(str(tmp_copy), WithWindow=False)

        self._fill_summary_slide(prs.Slides(1), tickets)

        original_slide3 = prs.Slides(3)
        for ticket in tickets:
            dup_sol = prs.Slides(2).Duplicate()
            self._fill_ticket_slide(dup_sol.Item(1), ticket, "Solution", attachments_map)

            dup_str = original_slide3.Duplicate()
            self._fill_ticket_slide(dup_str.Item(1), ticket, "Self Test Report", attachments_map)

        n = len(tickets)
        if n:
            prs.Slides(3 + 2 * n).Delete()
            prs.Slides(2).Delete()

        prs.Save()
        prs.Close()
        ppt_app.Quit()

        final_path = Path(output_path) if output_path else self.template_path.parent.parent / "jira_export_pptx.pptx"

        import time
        for attempt in range(3):
            try:
                if final_path.exists():
                    final_path.unlink()
                shutil.copy(str(tmp_copy), str(final_path))
                break
            except OSError:
                if attempt < 2:
                    time.sleep(1)
                else:
                    raise
        shutil.rmtree(tmp_dir, ignore_errors=True)
        return final_path

    # -- slide filling ----------------------------------------------------

    def _fill_summary_slide(self, slide: Any, tickets: list[dict[str, Any]]) -> None:
        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes.Item(i)
            if not shape.HasTable:
                continue
            table = shape.Table
            for ticket in tickets:
                table.Rows.Add()
                is_bug = ticket.get("issue_type", "").lower() == "bug"
                bug_color = "FF0000" if is_bug else None
                _set_cell_text(table.Cell(table.Rows.Count, 1), ticket.get("Module", ""), color=bug_color)
                _set_cell_text(table.Cell(table.Rows.Count, 2), ticket.get("key", ""), color=bug_color)
            return

    def _fill_ticket_slide(
        self, slide: Any, ticket: dict[str, Any], role: str,
        attachments_map: dict[str, list[JiraAttachment]],
    ) -> None:
        key = ticket.get("key", "")
        summary = ticket.get("summary", "")
        module = ticket.get("Module", "")
        is_bug = ticket.get("issue_type", "").lower() == "bug"
        bug_color = "FF0000" if is_bug else None
        ticket_atts = attachments_map.get(key, [])

        for i in range(1, slide.Shapes.Count + 1):
            shape = slide.Shapes.Item(i)
            if not shape.HasTextFrame:
                self._replace_picture(slide, shape, ticket, role, ticket_atts)
                continue
            raw = shape.TextFrame.TextRange.Text.strip()
            if raw == "[Module]":
                _set_shape_text(shape, module, color=bug_color)
            elif raw in ("[key] [summary]", "[key][summary]"):
                _set_shape_text_mixed(shape, [(key, bug_color), (" " + summary, None)])
            elif raw.startswith("[") and raw.endswith("]"):
                val = ticket.get(raw[1:-1], "")
                if isinstance(val, list):
                    val = ", ".join(str(v) for v in val)
                if val:
                    _set_shape_text(shape, str(val))

    # -- picture handling -------------------------------------------------

    def _replace_picture(
        self, slide: Any, picture_shape: Any, ticket: dict[str, Any],
        role: str, ticket_atts: list[JiraAttachment],
    ) -> None:
        """Replace a picture placeholder — handles single or multi-image."""
        field_urls = ticket.get(role)
        if not field_urls:
            return

        # Normalise to list of URL strings
        if isinstance(field_urls, list):
            urls = [str(u).strip() for u in field_urls if str(u).strip().startswith("http")]
        else:
            urls = [str(field_urls).strip()] if str(field_urls).strip().startswith("http") else []
        if not urls:
            return

        # Single image (Solution or only one URL)
        if role == "Solution" or len(urls) == 1:
            url = urls[0]
            att = next((a for a in ticket_atts if url in a.url or a.url in url), None)
            local_path = self._download(att, url)
            if local_path and Path(local_path).exists():
                from PIL import Image as PILImage
                img = PILImage.open(local_path)
                img_w, img_h = img.size
                img.close()
                max_w, max_h = picture_shape.Width, picture_shape.Height
                scale = min(max_w / img_w, max_h / img_h)
                new_w, new_h = int(img_w * scale), int(img_h * scale)
                left = picture_shape.Left + (max_w - new_w) // 2
                top = picture_shape.Top + (max_h - new_h) // 2
                picture_shape.Delete()
                slide.Shapes.AddPicture(str(local_path), False, True, left, top, new_w, new_h)
            return

        # Multiple images -> grid layout
        local_paths: list[str] = []
        for url in urls:
            att = next((a for a in ticket_atts if url in a.url or a.url in url), None)
            lp = self._download(att, url)
            if lp and Path(lp).exists():
                local_paths.append(lp)
        if local_paths:
            self._place_images_grid(slide, picture_shape, local_paths)

    def _place_images_grid(self, slide: Any, placeholder_shape: Any, local_paths: list[str]) -> None:
        n = len(local_paths)
        cols = math.ceil(math.sqrt(n))
        rows = math.ceil(n / cols)
        area_left = placeholder_shape.Left
        area_top = placeholder_shape.Top
        area_w, area_h = placeholder_shape.Width, placeholder_shape.Height
        cell_w, cell_h = area_w / cols, area_h / rows
        placeholder_shape.Delete()

        for idx, path in enumerate(local_paths):
            r, c = idx // cols, idx % cols
            from PIL import Image as PILImage
            img = PILImage.open(path)
            img_w, img_h = img.size
            img.close()
            scale = min(cell_w / img_w, cell_h / img_h)
            new_w, new_h = int(img_w * scale), int(img_h * scale)
            cx = area_left + c * cell_w
            cy = area_top + r * cell_h
            slide.Shapes.AddPicture(
                str(path), False, True,
                cx + (cell_w - new_w) / 2, cy + (cell_h - new_h) / 2,
                new_w, new_h,
            )

    # -- download ---------------------------------------------------------

    def _download(self, att: JiraAttachment | None, url: str) -> str | None:
        """Unified download: try attachment URL first, then raw URL. Cached."""
        cache_key = att.url if att else url
        if cache_key in self._download_cache:
            cached = self._download_cache[cache_key]
            if Path(cached).exists():
                return cached

        session = self._session
        if session is None:
            import requests
            session = requests

        try:
            download_url = att.url if att else url
            resp = session.get(download_url, timeout=30)
            if not resp.ok:
                return None
            ext = Path(att.filename).suffix if att else self._ext_from_content_type(resp)
            tmp = tempfile.NamedTemporaryFile(suffix=ext or ".png", delete=False)
            tmp.write(resp.content)
            tmp.close()
            self._download_cache[cache_key] = tmp.name
            return tmp.name
        except Exception:
            return None

    @staticmethod
    def _ext_from_content_type(resp: Any) -> str:
        ct = resp.headers.get("Content-Type", "")
        if "jpeg" in ct or "jpg" in ct:
            return ".jpg"
        if "gif" in ct:
            return ".gif"
        if "webp" in ct:
            return ".webp"
        return ".png"
