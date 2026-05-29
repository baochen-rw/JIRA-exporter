"""
Discover JIRA projects, fields, and metadata.
Run this first to find your project key and custom field IDs.

Usage:
    python -m src.discover
"""

import sys

from src.config import JiraConfig
from src.client import JiraClient


def discover(config: JiraConfig) -> None:
    client = JiraClient(config)

    print("=" * 60)
    print("JIRA DISCOVERY REPORT")
    print("=" * 60)

    # 1. Test connection & list projects
    print("\n[1] PROJECTS")
    print("-" * 40)
    try:
        projects = client._get("/project")
        for p in projects:
            print(f"  Key: {p['key']:<10} Name: {p.get('name', 'N/A')}")
    except Exception as e:
        print(f"  ERROR: {e}")
        return

    # 2. Get field metadata
    print("\n[2] STANDARD FIELDS (with IDs)")
    print("-" * 40)
    try:
        fields = client._get("/field")
        for f in fields:
            ftype = f.get("fieldType", "")
            schema = f.get("schema", {})
            stype = schema.get("type", "")
            scustom = schema.get("custom", "")
            custom_id = schema.get("customId", "")
            print(f"  {f['id']:<35}  {f['name']}")
            if scustom:
                print(f"    -> type={stype}, custom={scustom}, customId={custom_id}")
    except Exception as e:
        print(f"  ERROR: {e}")

    # 3. Try fetching one issue to find custom fields
    print("\n[3] CUSTOM FIELDS FROM A REAL TICKET")
    print("-" * 40)
    if not config.project_key or config.project_key == "YOUR-PROJECT-KEY":
        print("  Set 'project_key' in config.json first, then re-run this script.")
        print("  Or try with a known ticket key: python -m src.discover -k VEL-16204")
    else:
        try:
            jql = f'project = "{config.project_key}" ORDER BY created DESC'
            issues = client._get("/search", {
                "jql": jql,
                "maxResults": 1,
                "fields": "*all",
            })
            issues_raw = issues.get("issues", [])
            if not issues_raw:
                print(f"  No tickets found for project '{config.project_key}'.")
            else:
                sample = issues_raw[0]
                print(f"  Sample ticket: {sample['key']}")
                raw_fields = sample.get("fields", {})
                for field_id, value in sorted(raw_fields.items()):
                    if field_id.startswith("customfield_") and value is not None:
                        preview = str(value)[:60]
                        print(f"  {field_id:<30}  {preview}")
        except Exception as e:
            print(f"  ERROR: {e}")

    print("\n" + "=" * 60)
    print("NEXT STEP: Fill in config.json with the project key and")
    print("custom field IDs discovered above, then run:")
    print("  python -m src.main -k PROJECT-123")
    print("=" * 60)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Discover JIRA projects and fields.")
    parser.add_argument("-c", "--config", default="config.json", help="Config file path")
    parser.add_argument("-k", "--key", help="Ticket key to inspect for custom fields")
    args = parser.parse_args()

    try:
        config = JiraConfig.from_file(args.config)
    except FileNotFoundError:
        print(f"Error: config.json not found at '{args.config}'")
        sys.exit(1)

    if args.key:
        config.project_key = args.key.split("-")[0]

    discover(config)
