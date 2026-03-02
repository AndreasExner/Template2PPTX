# Flow: FillChart

## Purpose

Fills a chart in a PowerPoint template (.pptx) with new category labels and series data. The chart is identified by its shape name. The modified presentation is stored in Azure Blob Storage and the blob ID is returned.

## When to Use

Call this flow when the user wants to populate or update a chart in the PowerPoint template with data (e.g., bar chart, pie chart, line chart). Each call fills one chart.

## Input Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `placeholders` | string | Yes | JSON string containing the chart data. Structure: `{ "shapeName": "<name>", "categories": [...], "series": [...] }`. See example below. |
| `conversation.id` | string | Yes | The unique conversation ID. Used to locate the template and store the result. |

### Placeholders JSON Structure

```json
{
  "shapeName": "Slide3_Chart1",
  "categories": ["Overall Score", "Reliability", "Safety", "Technology", "Value for Money"],
  "series": [
    {
      "name": "Our Model",
      "values": [85, 88, 92, 78, 81]
    },
    {
      "name": "Competitor A",
      "values": [79, 82, 88, 85, 74]
    }
  ]
}
```

| Field | Type | Description |
|---|---|---|
| `shapeName` | string | The name of the chart shape in the template (from the PowerPoint Selection Pane). Case-insensitive. |
| `categories` | string[] | Category labels for the chart (X-axis labels or pie slice names). |
| `series` | array | One or more data series. Each series has a `name` (legend label) and `values` (numeric data points, same length as categories). |

## Output

| Parameter | Type | Description |
|---|---|---|
| `blob.id` | string | The blob ID (filename) of the modified PowerPoint file in Azure Blob Storage container `temp` (e.g., `result_abc123.pptx`). |

## Flow Logic (internal)

1. **Get template** — Retrieve the .pptx template from blob storage.
2. **Call API** — POST to `/api/FillChart` with multipart body:
   - `template`: the .pptx binary
   - `chartData`: the JSON string from input
3. **Store result** — Save the modified .pptx to container `temp` with a new blob ID.
4. **Return** — Return the `blob.id` of the stored result.

## Error Handling

- If `shapeName` does not match any chart shape in the template, the API returns HTTP 400 with a list of all available chart shape names.
- If `categories` or `series` are missing or empty, the API returns HTTP 400.
- If `values` array length does not match `categories` length, the chart may render incorrectly.
- If the JSON is malformed, the API returns HTTP 400.

## Example

**Agent action:**
1. Agent gathers chart data (e.g., competitor scores from research).
2. Agent calls this flow with:
   - `placeholders`: `{"shapeName":"Slide3_Chart1","categories":["Overall Score","Reliability","Safety","Technology","Value for Money"],"series":[{"name":"Our Model","values":[85,88,92,78,81]},{"name":"Competitor A","values":[79,82,88,85,74]}]}`
   - `conversation.id`: `conv-9f3a`
3. Flow returns: `blob.id`: `result_9f3a.pptx`

## Constraints

- Only **one chart** can be filled per flow call. To fill multiple charts, call this flow multiple times with different `shapeName` values.
- The chart type (bar, pie, line, etc.) is defined in the template and cannot be changed via this API — only the data is updated.
- Both the chart XML cache and the embedded Excel workbook are updated, so the chart remains **fully editable** in PowerPoint.
