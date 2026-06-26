
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
        +str output_dir
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

    class PPTExporter {
        +Path template_path
        +Path output_dir
        +transform_for_ppt(tickets) Path
        +fill(tickets, client) Path
        -_fill_overview_table(pres, tickets) void
        -_fill_ticket_slides(pres, tickets, client) void
        -_duplicate_slide_pair(pres, idx1, idx2) void
        -_replace_text_in_slide(slide, title_text, module) void
        -_insert_images_on_slide(slide, urls, client) void
        -_find_picture_placeholder(slide) Shape|None
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
    participant P as PPTExporter

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
    Note over M,P: Step 3 — Transform to PPT-friendly JSON
    M->>P: PPTExporter(template_path, output_dir)
    M->>P: ppt_exporter.transform_for_ppt(tickets)
    P->>P: Build attachment lookup per ticket
    P->>P: Resolve media_alts to attachment URLs
    P->>F: json.dump() → jira_export_pptx.json
    F-->>P: Success
    P-->>M: Output Path
    end

    rect rgb(255, 235, 235)
    Note over M,F: Step 4 — Fill PPT template via COM
    M->>F: open(ppt_json) → json.load() → ppt_data
    M->>P: ppt_exporter.fill(ppt_data, exporter.client)
    P->>F: shutil.copy2(template → output.pptx)
    P->>P: COM Dispatch("PowerPoint.Application")
    P->>P: Open presentation
    P->>P: _fill_overview_table: group by module, Rows.Add()
    P->>P: _duplicate_slide_pair × (N-1) before filling
    P->>P: _replace_text_in_slide: [key] [summary], [Module]
    P->>J: client.download(Solution URLs)
    J-->>P: temp image paths
    P->>P: _insert_images_on_slide (grid, aspect ratio)
    P->>J: client.download(Self Test Report URLs)
    J-->>P: temp image paths
    P->>P: _insert_images_on_slide (grid, aspect ratio)
    P->>F: Save → Close → Quit
    F-->>P: jira_export_pptx.pptx
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
        C[dict<br/>via asdict]
        D[jira_export.json<br/>file on disk]
    end

    subgraph "Step 3 — Transform"
        E[JiraTicket<br/>domain object]
        F[jira_export_pptx.json<br/>PPT-friendly JSON]
    end

    subgraph "Step 4 — Fill PPT"
        G[jira_export_pptx.json<br/>loaded into memory]
        H[jira_export_pptx.pptx<br/>COM-filled presentation]
    end

    A -->|"JiraTicket.from_dict()"| B
    B -->|"asdict(ticket)"| C
    C -->|"json.dump()"| D
    D -->|"tickets list passed directly"| E
    E -->|"PPTExporter.transform_for_ppt()"| F
    F -->|"json.load()"| G
    G -->|"PPTExporter.fill()"| H
```

