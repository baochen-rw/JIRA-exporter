"""
Command-line interface for the JIRA exporter.
"""

import argparse
import sys
from pathlib import Path

from src.config import JiraConfig
from src.exporter import JiraExporter
from src.ppt_exporter import PPTTemplateFiller


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Export JIRA tickets to structured JSON and/or PowerPoint.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python -m src.main                          # JSON export only\n"
            "  python -m src.main --ppt                    # JSON + PPTX export\n"
            "  python -m src.main -c my_config.json        # custom config\n"
            "  python -m src.main -k PROJ-1 PROJ-2 PROJ-3 # specific tickets\n"
            "  python -m src.main -j 'status = Done'       # custom JQL filter\n"
            "  python -m src.main --ppt -o report.json    # custom output\n"
        ),
    )
    parser.add_argument(
        "-c", "--config",
        default="config.json",
        help="Path to config.json (default: config.json)",
    )
    parser.add_argument(
        "-k", "--keys",
        nargs="+",
        metavar="KEY",
        help="Export specific ticket keys instead of running JQL query",
    )
    parser.add_argument(
        "-j", "--jql",
        metavar="JQL",
        help="Override the JQL clause in the config (appended to project filter)",
    )
    parser.add_argument(
        "-o", "--output",
        metavar="FILE",
        help="Override output file path (JSON or PPTX base name when used with --ppt)",
    )
    parser.add_argument(
        "--ppt",
        action="store_true",
        help="Also generate a PowerPoint report from PPTTemplate/Template.pptx",
    )
    parser.add_argument(
        "--ppt-template",
        metavar="PATH",
        dest="ppt_template",
        help="Path to a custom PPTX template (default: PPTTemplate/Template.pptx)",
    )
    args = parser.parse_args()

    try:
        config = JiraConfig.from_file(args.config)
    except FileNotFoundError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1
    except Exception as exc:
        print(f"Error loading config: {exc}", file=sys.stderr)
        return 1

    if args.jql:
        config.jql = args.jql
    if args.output:
        config.output_file = args.output

    exporter = JiraExporter(config)

    try:
        if args.keys:
            tickets = exporter.export_tickets(args.keys)
        else:
            tickets = exporter.run()
    except Exception as exc:
        print(f"Export failed: {exc}", file=sys.stderr)
        return 1

    if args.ppt and tickets:
        try:
            filler = PPTTemplateFiller(
                template_path=args.ppt_template if args.ppt_template else None,
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


if __name__ == "__main__":
    sys.exit(main())
