"""
Export JIRA tickets to structured JSON, collecting image URLs from both
ticket descriptions and the attachments list.
"""

import json
from pathlib import Path

from .client import JiraClient, JiraTicket
from .config import JiraConfig


class JiraExporter:
    def __init__(self, config: JiraConfig):
        self.config = config
        self.client = JiraClient(config)

    def run(self) -> list[dict]:
        jql = self.config.build_jql()
        print(f"Querying JIRA with JQL: {jql}")

        self.client.load_field_names()
        tickets = self.client.search(
            jql=jql,
            fields=self.config.fields,
        )
        print(f"Fetched {len(tickets)} ticket(s).")

        results = [t.to_dict() for t in tickets]
        self._write_output(results)

        self._print_summary(results)
        return results

    def export_tickets(self, keys: list[str]) -> list[dict]:
        print(f"Exporting {len(keys)} specific ticket(s): {', '.join(keys)}")
        self.client.load_field_names()

        results = []
        for key in keys:
            try:
                ticket = self.client.get_ticket(key, self.config.fields)
                results.append(ticket.to_dict())
            except Exception as exc:  # pragma: no cover — surface per-ticket errors
                print(f"  [WARN] Failed to fetch '{key}': {exc}")

        self._write_output(results)
        self._print_summary(results)
        return results

    def _write_output(self, results: list[dict]) -> None:
        path = Path(self.config.output_file)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"Output written to: {path.resolve()}")

    def _print_summary(self, results: list[dict]) -> None:
        total_images = 0
        for t in results:
            total_images += len(t.get("image_urls", [])) + len(
                [a for a in t.get("attachments", []) if a.get("filename", "").lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.webp', '.svg', '.bmp'))]
            )
        print(f"Total tickets: {len(results)}")
        print(f"Total image references: {total_images}")
