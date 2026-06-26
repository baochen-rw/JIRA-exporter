# JIRA Exporter

Export JIRA Cloud tickets to structured JSON, including image URLs from both ticket descriptions and attachments.

## Setup

1. **Install dependencies**

   ```bash
   pip install -r requirements.txt
   ```

2. **Configure credentials**

   Edit `config.json`:

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

## Usage

**Export by JQL query (default from config):**

```bash
python -m src.main
```

**Export by specific ticket keys:**

```bash
python -m src.main -k PROJ-1 PROJ-2 PROJ-3
```


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

