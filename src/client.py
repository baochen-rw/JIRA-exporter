"""
Low-level JIRA API client using requests directly.
Handles authentication and raw API calls for maximum control.
"""

import re
from dataclasses import dataclass, field
from typing import Any

import requests

from .config import JiraConfig


IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp"}

# Shared cache: field ID -> field name (fetched once from JIRA /field API)
FIELD_NAMES: dict[str, str] = {}
IMG_TAG_RE = re.compile(
    r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>', re.IGNORECASE
)
SRC_SET_RE = re.compile(r'srcset=["\']([^"\']+)["\']', re.IGNORECASE)


@dataclass
class JiraAttachment:
    filename: str
    mime_type: str
    url: str
    size: int
    created: str
    author: str

    @property
    def is_image(self) -> bool:
        ext = "." + self.filename.rsplit(".", 1)[-1].lower()
        return ext in IMAGE_EXTENSIONS


@dataclass
class JiraTicket:
    key: str
    id: int
    summary: str
    description: str | None
    status: str
    priority: str | None
    issue_type: str
    project: str
    reporter: str | None
    assignee: str | None
    created: str | None
    updated: str | None
    labels: list[str] = field(default_factory=list)
    components: list[str] = field(default_factory=list)
    fix_versions: list[str] = field(default_factory=list)
    attachments: list[JiraAttachment] = field(default_factory=list)
    image_urls: list[str] = field(default_factory=list)
    custom_fields: dict[str, Any] = field(default_factory=dict)

    @classmethod
    def _extract_text(cls, raw: Any) -> Any:
        """Convert ADF dict to text, or return value as-is."""
        if raw is None:
            return None
        if isinstance(raw, str):
            return raw
        if isinstance(raw, dict):
            rendered = raw.get("renderedBody", "")
            if isinstance(rendered, str) and rendered:
                return rendered
            text, _ = cls._adf_to_text(raw)
            return text if text.strip() else None
        if isinstance(raw, list):
            return [cls._extract_text(item) for item in raw]
        return raw

    @classmethod
    def _extract_display_name(cls, raw: Any) -> str | None:
        if raw is None:
            return None
        if isinstance(raw, str):
            return raw
        if isinstance(raw, dict):
            return raw.get("displayName") or raw.get("name")
        return str(raw)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "JiraTicket":
        fields = data.get("fields", {})

        attachments = []
        for att in fields.get("attachment", []):
            attachments.append(
                JiraAttachment(
                    filename=att.get("filename", ""),
                    mime_type=att.get("mimeType", ""),
                    url=att.get("content", ""),
                    size=att.get("size", 0),
                    created=att.get("created", ""),
                    author=att.get("author", {}).get("displayName", "Unknown"),
                )
            )

        description, image_urls = cls._normalize_description(fields.get("description"))

        # Collect all custom fields
        custom_fields: dict[str, Any] = {}
        for key, value in fields.items():
            if key.startswith("customfield_"):
                display_name = FIELD_NAMES.get(key, key)
                extracted_text = cls._extract_text(value)
                custom_fields[display_name] = {
                    "id": key,
                    "value": extracted_text,
                    "raw": value,
                }

        return cls(
            key=data.get("key", ""),
            id=data.get("id", 0),
            summary=fields.get("summary", ""),
            description=description,
            status=fields.get("status", {}).get("name", ""),
            priority=fields.get("priority", {}).get("name"),
            issue_type=fields.get("issuetype", {}).get("name", ""),
            project=fields.get("project", {}).get("key", ""),
            reporter=fields.get("reporter", {}).get("displayName"),
            assignee=fields.get("assignee", {}).get("displayName"),
            created=fields.get("created"),
            updated=fields.get("updated"),
            labels=fields.get("labels", []),
            components=[c.get("name", "") for c in fields.get("components", [])],
            fix_versions=[v.get("name", "") for v in fields.get("fixVersions", [])],
            attachments=attachments,
            image_urls=image_urls,
            custom_fields=custom_fields,
        )

    @classmethod
    def _adf_to_text(cls, adf: dict) -> str:
        """Flatten Atlassian Document Format (ADF) to plain text, extracting image URLs."""
        urls: list[str] = []
        lines: list[str] = []

        def walk(node: Any) -> None:
            if isinstance(node, dict):
                if node.get("type") == "text":
                    text = node.get("text", "")
                    marks = node.get("marks", [])
                    if any(m.get("type") == "code" for m in marks):
                        text = f"`{text}`"
                    lines.append(text)
                elif node.get("type") == "inlineCard":
                    attrs = node.get("attrs", {})
                    url = attrs.get("url") or attrs.get("data", {}).get("url", "")
                    if url:
                        urls.append(url)
                elif node.get("type") == "media":
                    attrs = node.get("attrs", {})
                    url = attrs.get("url", "")
                    if url:
                        urls.append(url)
                for child in node.get("content", []):
                    walk(child)
            elif isinstance(node, list):
                for item in node:
                    walk(item)

        walk(adf)
        return "\n".join(lines), urls

    @classmethod
    def _normalize_description(cls, raw: Any) -> tuple[str | None, list[str]]:
        """Handle description as plain HTML string, ADF dict, or None."""
        if raw is None:
            return None, []
        if isinstance(raw, str):
            return raw, cls._extract_image_urls(raw, [])
        if isinstance(raw, dict):
            rendered = raw.get("renderedBody", "")
            if isinstance(rendered, str) and rendered:
                return rendered, cls._extract_image_urls(rendered, [])
            text, adf_urls = cls._adf_to_text(raw)
            html_urls = cls._extract_image_urls(text, [])
            combined = list(dict.fromkeys(adf_urls + html_urls))
            return text or None, combined
        return str(raw), []

    @classmethod
    def _extract_image_urls(cls, description: str | None, attachments: list[JiraAttachment]) -> list[str]:
        urls: list[str] = []

        if not description:
            return urls

        for match in IMG_TAG_RE.finditer(description):
            src = match.group(1).strip()
            if src:
                urls.append(src)

        for match in SRC_SET_RE.finditer(description):
            for part in match.group(1).split(","):
                url = part.strip().split()[0]
                if url:
                    urls.append(url)

        attachment_urls = {att.url for att in attachments if att.is_image}
        urls = [u for u in urls if u not in attachment_urls]

        seen: set[str] = set()
        unique = []
        for u in urls:
            if u not in seen:
                seen.add(u)
                unique.append(u)

        return unique

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "key": self.key,
            "summary": self.summary,
            "issue_type": self.issue_type,
        }
        for name, data in self.custom_fields.items():
            if name == "Review Results":
                continue
            val = data.get("value")
            if name == "Solution" and val:
                parts = val.strip().split(maxsplit=1)
                result["Module"] = parts[0] if len(parts) > 0 else None
                result["Diff"] = parts[1] if len(parts) > 1 else None
            elif val is not None or name == "Self Test Report":
                result[name] = val
        return result


class JiraClient:
    def __init__(self, config: JiraConfig):
        self.url = config.url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update(
            {
                "Accept": "application/json",
                "Content-Type": "application/json",
                "X-Atlassian-Token": "no-check",
            }
        )
        self.session.auth = (config.email, config.api_token)
        self._rest_base = f"{self.url}/rest/api/3"

    def _get(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        url = f"{self._rest_base}{path}"
        resp = self.session.get(url, params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def load_field_names(self) -> None:
        """Fetch all field names from JIRA and cache them."""
        if FIELD_NAMES:
            return
        try:
            fields = self._get("/field")
            for f in fields:
                FIELD_NAMES[f["id"]] = f["name"]
        except Exception:
            pass

    def search(self, jql: str, fields: list[str], max_results: int = 100) -> list[JiraTicket]:
        tickets: list[JiraTicket] = []
        start_at = 0

        while True:
            data = self._get(
                "/search",
                {
                    "jql": jql,
                    "fields": ",".join(fields),
                    "startAt": start_at,
                    "maxResults": min(max_results, 100),
                },
            )

            for item in data.get("issues", []):
                tickets.append(JiraTicket.from_dict(item))

            total = data.get("total", 0)
            fetched = start_at + len(data.get("issues", []))
            if fetched >= total:
                break

            start_at = fetched

        return tickets

    def get_ticket(self, key: str, fields: list[str]) -> JiraTicket:
        data = self._get(f"/issue/{key}", {"fields": ",".join(fields)})
        return JiraTicket.from_dict(data)
