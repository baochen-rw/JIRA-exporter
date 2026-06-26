import argparse
import json
import sys
from pathlib import Path

from src.definition import JiraConfig, JiraExporter
from src.ppt_exporter import PPTExporter


def main() -> int:
    parser = argparse.ArgumentParser(description="JIRA Exporter")
    parser.add_argument("keys", nargs="+", help="JIRA ticket keys to export")
    args = parser.parse_args()

    config = JiraConfig.from_file("config.json")

    # ── Step 1 & 2: Fetch tickets from JIRA API → Store as JSON ──
    exporter = JiraExporter(config)
    tickets_obj = exporter.export_tickets(args.keys)

    # ── Step 3: Transform to PPT-friendly JSON ──
    if config.ppt_export:
        ppt_exporter = PPTExporter(Path(config.ppt_template), Path(config.output_dir))
        ppt_json = ppt_exporter.transform_for_ppt(tickets_obj)
        with open(ppt_json, encoding="utf-8") as f:
            ppt_data = json.load(f)
        ppt_exporter.fill(ppt_data, exporter.client)

    return 0


if __name__ == "__main__":
    sys.exit(main())
