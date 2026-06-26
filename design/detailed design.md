# JIRA Exporter — Detailed Class Design


## 1. Class Overview

```mermaid
classDiagram
    direction LR

    class JiraConfig {
        +str url
        +str email
        +str api_token
        +str project_key
        +str output_file
        +bool ppt_export
        +str ppt_template
        +list~str~ fields
        +from_file(path) JiraConfig
    }

    class JiraAttachment {
        +str filename
        +str mime_type
        +str url
        +int size
        +str created
        +str author
        +is_image bool
    }

    class JiraTicket {
        +str key
        +int id
        +str summary
        +str|None description
        +str status
        +str|None priority
        +str issue_type
        +str project
        +str|None reporter
        +str|None assignee
        +str|None created
        +str|None updated
        +list~str~ labels
        +list~str~ components
        +list~str~ fix_versions
        +list~JiraAttachment~ attachments
        +list~str~ image_urls
        +dict custom_fields
        +from_dict(data) JiraTicket
        -_extract_text(raw) Any
        -_adf_to_text(adf) tuple
        -_extract_image_urls(desc) list~str~
    }

    class JiraClient {
        -str url
        +Session session
        -str _rest_base
        -_get(path, params) dict
        +load_field_names() void
        +get_ticket(key, fields) JiraTicket
        +download(url, suffix) str|None
    }

    class JiraExporter {
        +JiraConfig config
        +JiraClient client
        +export_tickets(keys) list~JiraTicket~
        -_write_output(tickets) void
    }

    class PPTTemplateFiller {
        +Path template_path
        -dict _download_cache
        -Any _client
        +fill(tickets, attachments_map, client) Path
        -_fill_summary_slide(slide, tickets) void
        -_fill_ticket_slide(slide, ticket, role, attachments_map) void
        -_replace_picture(slide, shape, ticket, role, atts) void
        -_place_images_grid(slide, shape, paths) void
        -_download(att, url) str|None
    }

    class MainModule {
        <<module>>
        +main() int
    }
```

## 2. Pipeline Sequence (Detailed)

```mermaid
sequenceDiagram
    participant U as User
    participant M as main.main()
    participant C as JiraConfig
    participant E as JiraExporter
    participant J as JiraClient
    participant API as JIRA API
    participant F as File System
    participant P as PPTTemplateFiller

    U->>M: python -m src.main KEYS
    M->>C: JiraConfig.from_file("config.json")
    C-->>M: JiraConfig object

    rect rgb(230, 245, 255)
    Note over E,F: Step 1 — Fetch tickets from JIRA API
    M->>E: JiraExporter(config)
    M->>E: exporter.export_tickets(keys)
    E->>J: JiraClient(config)
    J->>API: client.load_field_names() → GET /field
    API-->>J: Field metadata
    loop For each key
        J->>API: client.get_ticket(key, fields) → GET /issue/{key}
        API-->>J: Issue JSON
        J->>J: JiraTicket.from_dict(data)
    end
    J-->>E: list[JiraTicket]
    end

    rect rgb(230, 255, 230)
    Note over E,F: Step 2 — Store as JSON on File System
    E->>E: exporter._write_output(tickets)
    E->>F: json.dump(asdict(tickets)) → jira_export.json
    F-->>E: Success
    end

    rect rgb(255, 245, 230)
    Note over M,P: Step 3 — Generate PowerPoint from JSON
    M->>F: json.load(jira_export.json)
    F-->>M: tickets (list[dict])
    M->>M: Build attachments_map from tickets_obj
    M->>P: PPTTemplateFiller(template_path)
    M->>P: filler.fill(tickets, attachments_map, client)
    P->>P: _fill_summary_slide(slide, tickets)
    loop For each ticket
        P->>P: _determine_role(ticket)
        P->>P: _duplicate_slide(prs, tmpl)
        P->>P: _fill_ticket_slide(slide, ticket, role, atts)
        P->>P: _replace_picture(slide, shape, ticket, role, atts)
        P->>J: client.download(url, ext)
        J->>API: GET attachment/image URL
        API-->>J: Binary content
        J-->>P: Local temp file path
        P->>P: _place_images_grid(slide, shape, paths)
    end
    P->>F: prs.save() → jira_export_pptx.pptx
    F-->>P: Success
    P-->>M: Output Path
    end

    M-->>U: Exit with status
```


## 3. Data Transformation Through the Pipeline

The three-step pipeline transforms data through distinct representations at each stage.

```mermaid
graph LR
    subgraph "Step 1 — Fetch"
        A[JIRA REST API<br/>JSON response]
        B[JiraTicket<br/>domain object]
    end

    subgraph "Step 2 — Store"
        C[dict<br/>flat structure]
        D[jira_export.json<br/>file on disk]
    end

    subgraph "Step 3 — Generate"
        E[dict<br/>consumed as-is]
        F[jira_export_pptx.pptx<br/>PowerPoint file]
    end

    A -->|"JiraTicket.from_dict()"| B
    B -->|"asdict(ticket)"| C
    C -->|"json.dump()"| D
    D -->|"json.load()"| E
    E -->|"PPTTemplateFiller.fill()"| F
```

