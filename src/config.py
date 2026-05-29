"""
Configuration management — loads settings from config.json.
"""

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional


@dataclass
class JiraConfig:
    url: str
    email: str
    api_token: str
    project_key: str
    jql: Optional[str] = None
    output_file: str = "jira_export.json"
    fields: list[str] = field(
        default_factory=lambda: [
            "summary",
            "description",
            "status",
            "priority",
            "issuetype",
            "project",
            "reporter",
            "assignee",
            "created",
            "updated",
            "labels",
            "components",
            "fixVersions",
            "attachment",
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

    def build_jql(self) -> str:
        base = f'project = "{self.project_key}"'
        if self.jql:
            return f"{base} AND {self.jql}"
        return base
