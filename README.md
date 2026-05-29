# JIRA Exporter

Export JIRA Cloud tickets to structured JSON, including image URLs from both ticket descriptions and attachments.

## Setup

1. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

2. **Configure credentials**

   Copy the example file and fill in your values:

   ```bash
   copy config.example.json config.json
   ```

   Then edit `config.json`:

   ```json
   {
     "url": "https://your-domain.atlassian.net",
     "email": "your-login-email@example.com",
     "api_token": "YOUR-API-TOKEN",
     "project_key": "PROJ"
   }
   ```

3. **Get your JIRA API token**

   1. Go to [id.atlassian.com/manage-profile/security/api-tokens](https://id.atlassian.com/manage-profile/security/api-tokens)
   2. Click **Create API token**, give it a label, and click **Create**
   3. Copy the token immediately (you won't see it again)
   4. Paste it into `config.json` under the `api_token` field

## Usage

**Export by JQL query (default from config):**

```bash
python -m src.main
```

**Export by specific ticket keys:**

```bash
python -m src.main -k PROJ-1 PROJ-2 PROJ-3
```

**Add a custom JQL filter (appended to the project filter):**

```bash
python -m src.main -j "status = Done AND issuetype = Bug"
```

**Custom config and output path:**

```bash
python -m src.main -c config.json -o my_export.json
```

## Output Format

Each ticket in the JSON output includes:

| Field | Description |
|---|---|
| `key` | Ticket key (e.g. PROJ-123) |
| `summary` | Ticket title |
| `description` | Full ticket description (HTML) |
| `status` | Current status |
| `priority` | Priority level |
| `issue_type` | Bug, Story, Task, etc. |
| `project` | Project key |
| `reporter` | Creator display name |
| `assignee` | Assignee display name |
| `created` / `updated` | ISO timestamps |
| `labels` | List of labels |
| `components` | List of components |
| `fix_versions` | List of fix versions |
| `attachments` | Full attachment metadata (filename, URL, size, etc.) |
| `image_urls` | Image URLs found in description HTML (inline `src` and `srcset`) |
| `custom_fields` | All custom fields with `{ value, raw, display_name }` — ADF text converted to plain text |

## Auto-Discover Projects & Fields

If you don't know your project key or custom field IDs, run the discover script first:

```bash
pip install -r requirements.txt
python -m src.discover
```

It will print:
- All available projects (with keys)
- All standard and custom field IDs
- Sample custom field values from a real ticket

Then fill in `config.json` with what it finds, and run the exporter:

```bash
python -m src.main -k PROJECT-123
```

## Adding Custom Fields

Edit the `fields` list in `config.json`. JIRA custom fields use IDs like `customfield_10020`. You can add them alongside standard fields:

```json
"fields": [
  "summary",
  "description",
  "...",
  "customfield_10020"
]
```
