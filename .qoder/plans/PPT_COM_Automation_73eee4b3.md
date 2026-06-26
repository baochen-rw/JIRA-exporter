# PPT COM Automation

## Template Structure Analysis

The template (`res/PPTTemplate/Template.pptx`) has 3 slides per ticket:

| Slide | Layout | Shapes | Purpose |
|-------|--------|--------|---------|
| 1 | OUTLINE 1_Text&Image_Positive | TextBox "Summary" + Table (1 row x 2 cols: MODULE, ID) | Overview table |
| 2 | Title slide | "[key] [summary]", "[Module]", Picture, "Solution" | Solution slide |
| 3 | Title slide | "[key] [summary]", "[Module]", "Self Test Report" | Self Test Report slide |

## JSON Data Mapping (`export/jira_export_pptx.json`)

- `key` -> Slide 1 table "ID" column, Slides 2+3 title text
- `summary` -> Slides 2+3 title text
- `module` -> Slide 1 table "MODULE" column, Slides 2+3 module text
- `Solution` (customfield_10370) -> Download images, insert into Slide 2 picture area
- `Self Test Report` (customfield_10432) -> Download images, insert into Slide 3 picture area

## Task 1: Fix field name in `transform_for_ppt`

In `src/ppt_exporter.py`, rename the `"Diff"` key to `"Self Test Report"` in the output dict (line 74).

## Task 2: Implement `fill()` method in `PPTExporter`

The method signature is already stubbed in `src/ppt_exporter.py:83-91`. Implementation:

1. **Copy template** to `output_dir / jira_export_pptx.pptx` using `shutil.copy2`
2. **Open the copy via COM** (`win32com.client.Dispatch("PowerPoint.Application")`)
3. **Fill Slide 1 table (grouped by module)**:
   - Group all tickets by their `module` field
   - One row per module: MODULE name in col 0, comma-separated ticket keys in col 1
   - Example: `TBT | VEL-16399, VEL-16400`
   - Use `table.Rows.Add()` for each module group (keep the header row)
4. **For each ticket, fill Slides 2+3**:
   - Replace `[key] [summary]` with actual key + summary
   - Replace `[Module]` with actual module
   - Download Solution images via `client.download()`, arrange in a **grid layout** on Slide 2 (replace the existing picture placeholder)
   - Download Self Test Report images via `client.download()`, arrange in a **grid layout** on Slide 3
   - For tickets after the first, duplicate Slides 2+3 before filling
5. **Delete extra template slides** if needed (template has slides 2-3 for one ticket; duplicate for N tickets)
6. **Save and close**

Key COM API notes (from memory):
- Use `table.Rows.Add()` for new rows (Row.Duplicate() is NOT supported)
- Use `slide.Shapes.Paste()` after `shape.Copy()` for duplication
- Use `slide.Shapes.AddPicture()` for inserting images

Multiple images grid layout:
- When a field (Solution or Self Test Report) has multiple images, arrange them in a grid on the same slide
- Calculate grid dimensions (e.g., 2 images = 1x2, 3-4 images = 2x2) based on the placeholder's position and size
- Each image is sized to fit its grid cell, maintaining aspect ratio

## Task 3: Wire up `fill()` in `main.py`

After `transform_for_ppt()`, call `fill()` with the ticket data, passing the JiraClient for image downloads:

```python
ppt_json = ppt_exporter.transform_for_ppt(tickets_obj)
# Load the PPT JSON
with open(ppt_json) as f:
    ppt_data = json.load(f)
ppt_exporter.fill(ppt_data, exporter.client)
```

## Task 4: Update `fill()` signature

Simplify the signature to `fill(self, tickets: list[dict], client: JiraClient) -> Path` since the attachments_map can be derived from the ticket data itself.

## Task 5: Test and verify

Run the exporter with a ticket key to verify the PPT is generated correctly.
