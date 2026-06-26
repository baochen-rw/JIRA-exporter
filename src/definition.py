"""
All class definitions for the JIRA Exporter project.
"""

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import requests


# ── Shared constants ────────────────────────────────────────────────────────

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp"}

# Shared cache: field ID -> field name (fetched once from JIRA /field API)
FIELD_NAMES: dict[str, str] = {}

IMG_TAG_RE = re.compile(
    r'<img[^>]+src=["\']([^"\']+)["\'][^>]*>', re.IGNORECASE
)
SRC_SET_RE = re.compile(r'srcset=["\']([^"\']+)["\']', re.IGNORECASE)


# ── JiraConfig ──────────────────────────────────────────────────────────────

@dataclass
class JiraConfig:
    url: str
    email: str
    api_token: str
    project_key: str
    output_file: str = "jira_export.json"
    ppt_export: bool = False
    ppt_template: str = "PPTTemplate/Template.pptx"
    fields: list[str] = field(
        default_factory=lambda: [
            "summary", "description", "status", "priority",
            "issuetype", "project", "reporter", "assignee",
            "created", "updated", "labels", "components",
            "fixVersions", "attachment",
        ]
    )

    @classmethod
    def from_file(cls, path: str | Path = "config.json") -> "JiraConfig":
        cfg_path = Path(path)
        if not cfg_path.exists():
            raise FileNotFoundError(f"Config file not found: {cfg_path.resolve()}")
        with open(cfg_path, encoding="utf-8") as f:
            raw = json.load(f)
        return cls(**{k: v for k, v in raw.items() if k in cls.__annotations__})


# ── JiraAttachment ──────────────────────────────────────────────────────────

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


# ── JiraTicket ──────────────────────────────────────────────────────────────

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

    # -- Parsing helpers --------------------------------------------------

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
            text, media_alts, _ = cls._adf_to_text(raw)
            if media_alts:
                return {"_text": text if text.strip() else None, "_media_alts": media_alts}
            return text if text.strip() else None
        if isinstance(raw, list):
            return [cls._extract_text(item) for item in raw]
        return raw

    @classmethod
    def _adf_to_text(cls, adf: dict) -> tuple[str, list[str], list[str]]:
        """Flatten ADF to plain text, extracting image URLs and media alt texts."""
        urls: list[str] = []
        media_alts: list[str] = []
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
                    alt = attrs.get("alt", "")
                    if url:
                        urls.append(url)
                    if alt:
                        media_alts.append(alt)
                for child in node.get("content", []):
                    walk(child)
            elif isinstance(node, list):
                for item in node:
                    walk(item)

        walk(adf)
        return "\n".join(lines), media_alts, urls

    @staticmethod
    def _extract_image_urls(description: str | None) -> list[str]:
        """Extract unique <img> URLs from an HTML description string."""
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
        seen: set[str] = set()
        unique = []
        for u in urls:
            if u not in seen:
                seen.add(u)
                unique.append(u)
        return unique

    # -- Construction / serialisation -------------------------------------

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "JiraTicket":
        fields = data.get("fields", {})

        attachments = [
            JiraAttachment(
                filename=att.get("filename", ""),
                mime_type=att.get("mimeType", ""),
                url=att.get("content", ""),
                size=att.get("size", 0),
                created=att.get("created", ""),
                author=att.get("author", {}).get("displayName", "Unknown"),
            )
            for att in fields.get("attachment", [])
        ]

        # Normalise description (handles plain HTML, ADF dict, or None)
        raw_desc = fields.get("description")
        if raw_desc is None:
            description, desc_image_urls = None, []
        elif isinstance(raw_desc, str):
            description, desc_image_urls = raw_desc, cls._extract_image_urls(raw_desc)
        elif isinstance(raw_desc, dict):
            rendered = raw_desc.get("renderedBody", "")
            if isinstance(rendered, str) and rendered:
                description, desc_image_urls = rendered, cls._extract_image_urls(rendered)
            else:
                text, _, adf_urls = cls._adf_to_text(raw_desc)
                html_urls = cls._extract_image_urls(text)
                desc_image_urls = list(dict.fromkeys(adf_urls + html_urls))
                description = text or None
        else:
            description, desc_image_urls = str(raw_desc), []

        # Filter out attachment URLs from description image URLs
        att_urls = {att.url for att in attachments if att.is_image}
        image_urls = [u for u in desc_image_urls if u not in att_urls]

        # Collect custom fields
        custom_fields: dict[str, Any] = {}
        for key, value in fields.items():
            if key.startswith("customfield_"):
                display_name = FIELD_NAMES.get(key, key)
                custom_fields[display_name] = {
                    "id": key,
                    "value": cls._extract_text(value),
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

    def to_dict(self) -> dict[str, Any]:
        result: dict[str, Any] = {
            "key": self.key,
            "summary": self.summary,
            "issue_type": self.issue_type,
        }
        attachment_map: dict[str, str] = {
            att.filename: att.url for att in self.attachments
        }
        for name, data in self.custom_fields.items():
            if name == "Review Results":
                continue
            val = data.get("value")
            media_alts: list[str] = []
            if isinstance(val, dict):
                media_alts = val.get("_media_alts", [])
                val = val.get("_text")
            if name == "Solution" and val:
                parts = val.strip().split(maxsplit=1)
                result["Module"] = parts[0] if len(parts) > 0 else None
                result["Solution"] = parts[1] if len(parts) > 1 else None
                continue
            if media_alts:
                urls = [attachment_map[alt] for alt in media_alts if alt in attachment_map]
                result[name] = urls if urls else media_alts
            elif val is not None:
                result[name] = val
            elif name in ("Solution", "Self Test Report"):
                result[name] = None
        return result


# ── JiraClient ──────────────────────────────────────────────────────────────

class JiraClient:
    def __init__(self, config: JiraConfig):
        self.url = config.url.rstrip("/")
        self.session = requests.Session()
        self.session.headers.update({
            "Accept": "application/json",
            "Content-Type": "application/json",
            "X-Atlassian-Token": "no-check",
        })
        self.session.auth = (config.email, config.api_token)
        self._rest_base = f"{self.url}/rest/api/3"

    def _get(self, path: str, params: dict[str, Any] | None = None) -> dict[str, Any]:
        resp = self.session.get(f"{self._rest_base}{path}", params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def _post(self, path: str, json_body: dict[str, Any]) -> dict[str, Any]:
        resp = self.session.post(f"{self._rest_base}{path}", json=json_body, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def load_field_names(self) -> None:
        """Fetch all field names from JIRA and cache them."""
        if FIELD_NAMES:
            return
        try:
            for f in self._get("/field"):
                FIELD_NAMES[f["id"]] = f["name"]
        except Exception:
            pass

    def search(self, jql: str, fields: list[str], max_results: int = 100) -> list[JiraTicket]:
        tickets: list[JiraTicket] = []
        next_page_token: str | None = None
        while True:
            payload: dict[str, Any] = {
                "jql": jql,
                "fields": fields,
                "maxResults": min(max_results, 100),
            }
            if next_page_token:
                payload["nextPageToken"] = next_page_token
            data = self._post("/search/jql", payload)
            for item in data.get("issues", []):
                tickets.append(JiraTicket.from_dict(item))
            next_page_token = data.get("nextPageToken")
            if not next_page_token or not data.get("issues"):
                break
        return tickets

    def get_ticket(self, key: str, fields: list[str]) -> JiraTicket:
        data = self._get(f"/issue/{key}", {"fields": ",".join(fields)})
        return JiraTicket.from_dict(data)


# ── JiraExporter ────────────────────────────────────────────────────────────

class JiraExporter:
    def __init__(self, config: JiraConfig):
        self.config = config
        self.client = JiraClient(config)
        self._last_tickets: list[JiraTicket] | None = None

    def run(self, jql: str) -> list[dict]:
        print(f"Querying JIRA with JQL: {jql}")
        self.client.load_field_names()
        tickets = self.client.search(jql=jql, fields=self.config.fields)
        print(f"Fetched {len(tickets)} ticket(s).")
        results = [t.to_dict() for t in tickets]
        self._last_tickets = tickets
        self._write_output(results)
        self._print_summary(results)
        return results

    def export_tickets(self, keys: list[str]) -> list[dict]:
        print(f"Exporting {len(keys)} specific ticket(s): {', '.join(keys)}")
        self.client.load_field_names()
        results: list[dict] = []
        raw_tickets: list[JiraTicket] = []
        for key in keys:
            try:
                ticket = self.client.get_ticket(key, self.config.fields)
                results.append(ticket.to_dict())
                raw_tickets.append(ticket)
            except Exception as exc:
                print(f"  [WARN] Failed to fetch '{key}': {exc}")
        self._write_output(results)
        self._last_tickets = raw_tickets
        self._print_summary(results)
        return results

    def _write_output(self, results: list[dict]) -> None:
        path = Path(self.config.output_file)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"Output written to: {path.resolve()}")

    def _print_summary(self, results: list[dict]) -> None:
        total_images = sum(
            len(t.get("image_urls", []))
            + len([
                a for a in t.get("attachments", [])
                if a.get("filename", "").lower().endswith(
                    (".png", ".jpg", ".jpeg", ".gif", ".webp", ".svg", ".bmp")
                )
            ])
            for t in results
        )
        print(f"Total tickets: {len(results)}")
        print(f"Total image references: {total_images}")
