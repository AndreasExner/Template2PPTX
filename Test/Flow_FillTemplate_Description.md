# Flow: FillTemplate

## Purpose

Replaces text placeholders in a PowerPoint template (.pptx) by shape name. Multiple shapes with multiple placeholders can be filled in a single call. The modified presentation is stored in Azure Blob Storage and the blob ID is returned.

## When to Use

Call this flow when the user wants to fill or update text content in the PowerPoint template. This covers titles, subtitles, KPIs, table cells, and any other text shapes that contain `{{placeholder}}` markers.

## Input Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `placeholders` | string | Yes | JSON string containing the placeholder data. Structure: `{ "<shapeName>": { "{{placeholder}}": "replacement value" } }`. Multiple shapes and multiple placeholders per shape are supported. Shape names must match the template schema (case-insensitive). See example below. |
| `conversation.id` | string | Yes | The unique conversation ID. Used to locate the template and store the result. |

### Placeholders JSON Structure

```json
{
  "Slide1_Titel": {
    "{{Title}}": "Vehicle Assessment Report"
  },
  "Slide1_Subtitle": {
    "{{Subtitle}}": "Tesla Model 3",
    "{{Date}}": "26.02.2026"
  },
  "Slide2_KPI1": {
    "{{KPI-Overall-Score-Label}}": "Overall Score",
    "{{KPI-Overall-Score-Value}}": "85/100"
  }
}
```

Each top-level key is a **shape name** (from the PowerPoint Selection Pane). Each nested key is a **placeholder string** exactly as it appears in the template (including `{{}}`). Each nested value is the **replacement text**.

## Output

| Parameter | Type | Description |
|---|---|---|
| `blob.id` | string | The blob ID (filename) of the modified PowerPoint file in Azure Blob Storage container `temp` (e.g., `result_abc123.pptx`). |

## Flow Logic (internal)

1. **Get template** — Retrieve the .pptx template from blob storage.
2. **Call API** — POST to `/api/FillTemplate` with multipart body:
   - `template`: the .pptx binary
   - `placeholders`: the JSON string from input
3. **Store result** — Save the modified .pptx to container `temp` with a new blob ID.
4. **Return** — Return the `blob.id` of the stored result.

## Error Handling

- If a shape name does not match any shape in the template, those placeholders are silently skipped (the API returns a `notFound` list in the response log).
- If a placeholder string is not found in the shape's text, it is silently skipped.
- If the `placeholders` JSON is malformed, the API returns HTTP 400.

## Example

**Agent action:**
1. Agent gathers the required data (e.g., vehicle name, scores, dates).
2. Agent calls this flow with:
   - `placeholders`: `{"Slide1_Titel":{"{{Title}}":"Tesla Model 3 Assessment"},"Slide1_Subtitle":{"{{Subtitle}}":"Compact EV Segment","{{Date}}":"26.02.2026"},"Slide2_KPI1":{"{{KPI-Overall-Score-Label}}":"Overall Score","{{KPI-Overall-Score-Value}}":"85/100"}}`
   - `conversation.id`: `conv-9f3a`
3. Flow returns: `blob.id`: `result_9f3a.pptx`

## Constraints

- All placeholder replacements are sent in a **single call** — no need to call the flow multiple times for different shapes.
- Placeholder strings must match **exactly** (including `{{` and `}}`).
- Shape name matching is **case-insensitive**.
- Only text content is replaced — formatting (font, size, color) is preserved from the template.
