"""
Command-line interface for the JIRA exporter.
"""

import argparse
import sys
from pathlib import Path

from src.config import JiraConfig
from src.exporter import JiraExporter


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Export JIRA tickets to structured JSON.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python -m src.main                          # uses config.json\n"
            "  python -m src.main -c my_config.json        # custom config\n"
            "  python -m src.main -k PROJ-1 PROJ-2 PROJ-3  # specific tickets\n"
            "  python -m src.main -j 'status = Done'       # custom JQL filter\n"
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
        help="Override output file path",
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
            exporter.export_tickets(args.keys)
        else:
            exporter.run()
    except Exception as exc:
        print(f"Export failed: {exc}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
