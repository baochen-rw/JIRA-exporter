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
        +to_dict() dict
        -_extract_text(raw) Any
        -_adf_to_text(adf) tuple
        -_extract_image_urls(desc) list~str~
    }

    class JiraClient {
        -str url
        +Session session
        -str _rest_base
        -_get(path, params) dict
        -_post(path, json_body) dict
        +load_field_names() void
        +search(jql, fields, max_results) list~JiraTicket~
        +get_ticket(key, fields) JiraTicket
    }

    class JiraExporter {
        +JiraConfig config
        +JiraClient client
        -list _last_tickets
        +run(jql) list~dict~
        +export_tickets(keys) list~dict~
        -_write_output(results) void
        -_print_summary(results) void
    }

    class PPTTemplateFiller {
        +Path template_path
        -dict _download_cache
        -Any _session
        +fill(tickets, output_path, attachments_map, session) Path
        -_fill_summary_slide(slide, tickets) void
        -_fill_ticket_slide(slide, ticket, role, attachments_map) void
        -_replace_picture(slide, shape, ticket, role, atts) void
        -_place_images_grid(slide, shape, paths) void
        -_download(att, url) str|None
    }

    class MainModule {
        <<module>>
        +main() int
        -_ppt_only(args, config) int
    }
```


## 2. Class Relationships

```mermaid
graph TB
    subgraph "definition.py"
        CFG[JiraConfig]
        TK[JiraTicket]
        ATT[JiraAttachment]
        CL[JiraClient]
        EX[JiraExporter]
    end

    subgraph "Entry Point"
        MC[main.py]
    end

    subgraph "Presentation"
        PPT[ppt_exporter.py<br/>PPTTemplateFiller]
    end

    MC -->|imports| CFG
    MC -->|imports| EX
    MC -->|creates| PPT

    EX -->|owns| CL
    EX -->|produces| TK
    CL -->|creates| TK
    TK -->|contains *| ATT
    PPT -->|consumes| ATT
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
    B -->|"ticket.to_dict()"| C
    C -->|"json.dump()"| D
    D -->|"json.load()"| E
    E -->|"PPTTemplateFiller.fill()"| F
```

