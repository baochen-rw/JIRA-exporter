"""
Command-line interface for the JIRA exporter.
"""

import argparse
import json
import sys
from pathlib import Path

from definition import JiraConfig, JiraExporter
from ppt_exporter import PPTTemplateFiller


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Export JIRA tickets to structured JSON and/or PowerPoint.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python -m src.main                          # JSON export only\n"
            "  python -m src.main -k PROJ-1 PROJ-2 PROJ-3 # specific tickets\n"
            "  python -m src.main --ppt-only               # PPTX from existing JSON\n"
            "\n"
            "PPT export is controlled via config.json (ppt_export / ppt_template).\n"
        ),
    )
    parser.add_argument(
        "-k", "--keys",
        nargs="+",
        metavar="KEY",
        help="Export specific ticket keys instead of running JQL query",
    )
    parser.add_argument(
        "--jql",
        metavar="JQL",
        help="JQL filter to apply (e.g. 'sprint in openSprints()')",
    )
    parser.add_argument(
        "--ppt-only",
        action="store_true",
        dest="ppt_only",
        help="Generate PowerPoint from an existing JSON file without querying JIRA",
    )
    parser.add_argument(
        "--json-file",
        metavar="PATH",
        dest="json_file",
        help="Path to the JSON file for --ppt-only (default: output_file from config.json)",
    )
    args = parser.parse_args()

    try:
        config = JiraConfig.from_file("config.json")
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Error loading config: {exc}", file=sys.stderr)
        return 1

    # ── PPT-only mode: skip JIRA, read from existing JSON ──────────────────
    if args.ppt_only:
        return _ppt_only(args, config)

    exporter = JiraExporter(config)

    try:
        if args.keys:
            tickets = exporter.export_tickets(args.keys)
        elif args.jql:
            jql = f'project = "{config.project_key}" AND {args.jql}'
            tickets = exporter.run(jql)
        else:
            print(
                "Error: No ticket keys or JQL query provided.\n"
                "  Use -k KEY [KEY ...] to specify tickets, or\n"
                '  use --jql "sprint in openSprints()" to query by JQL.',
                file=sys.stderr,
            )
            return 1
    except Exception as exc:
        print(f"Export failed: {exc}", file=sys.stderr)
        return 1

    if config.ppt_export and tickets:
        try:
            filler = PPTTemplateFiller(
                template_path=config.ppt_template,
            )
            attachments_map: dict[str, list] = {}
            if exporter._last_tickets:
                for raw_ticket in exporter._last_tickets:
                    attachments_map[raw_ticket.key] = raw_ticket.attachments
            pptx_path = filler.fill(
                tickets,
                attachments_map=attachments_map,
                session=exporter.client.session,
            )
            print(f"PowerPoint written to: {pptx_path}")
        except Exception as exc:
            print(f"PowerPoint generation failed: {exc}", file=sys.stderr)
            import traceback
            traceback.print_exc()
            return 1

    return 0


def _ppt_only(args: argparse.Namespace, config: JiraConfig) -> int:
    """Generate PowerPoint from an existing JSON file without querying JIRA."""
    json_path = Path(args.json_file) if args.json_file else Path(config.output_file)

    if not json_path.exists():
        print(f"Error: JSON file not found: {json_path.resolve()}", file=sys.stderr)
        return 1

    try:
        with open(json_path, encoding="utf-8") as f:
            tickets = json.load(f)
    except Exception as exc:
        print(f"Error reading JSON file: {exc}", file=sys.stderr)
        return 1

    if not tickets:
        print("Error: JSON file is empty.", file=sys.stderr)
        return 1

    print(f"Loaded {len(tickets)} ticket(s) from {json_path.resolve()}")

    try:
        filler = PPTTemplateFiller(
            template_path=config.ppt_template,
        )
        pptx_path = filler.fill(tickets)
        print(f"PowerPoint written to: {pptx_path}")
    except Exception as exc:
        print(f"PowerPoint generation failed: {exc}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
