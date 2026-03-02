# Flow: ReplaceImage

## Purpose

Replaces a single image in a PowerPoint template (.pptx) by shape name. The modified presentation is stored in Azure Blob Storage and the blob path is returned.

## Prerequisites

Before calling this flow, the agent must:

1. **Upload the image file** to Azure Blob Storage container `temp` using a temporary filename (e.g., `img_abc123.jpg`). No folder structure required.
2. **Know the shape name** of the image placeholder in the PowerPoint template (e.g., `Slide1_Picture1`). Shape names are defined in the template schema.

## Input Parameters

| Parameter | Type | Required | Description |
|---|---|---|---|
| `blob.id` | string | Yes | The blob ID (filename) of the uploaded image file in Azure Blob Storage container `temp` (e.g., `img_abc123.jpg`). |
| `placeholder.name` | string | Yes | The name of the picture shape in the PowerPoint template to replace (e.g., `Slide1_Picture1`). Must match a shape name from the template schema. Case-insensitive. |
| `conversation.id` | string | Yes | The unique conversation ID. Used to locate the template and store the result in the correct folder. |

## Output

| Parameter | Type | Description |
|---|---|---|
| `blob.id` | string | The blob ID (filename) of the modified PowerPoint file in Azure Blob Storage container `temp` (e.g., `result_abc123.pptx`). |

## Flow Logic (internal)

1. **Get template** — Retrieve the .pptx template from blob storage.
2. **Get image** — Retrieve the image from container `temp` using `blob.id`.
3. **Get image metadata** — Read content type from blob metadata (e.g., `image/jpeg`).
4. **Call API** — POST to `/api/ReplaceImages` with multipart body:
   - `template`: the .pptx binary
   - `{placeholder.name}`: the image binary (form-field name = shape name)
5. **Store result** — Save the modified .pptx to container `temp` with a new blob ID.
6. **Return** — Return the `blob.id` of the stored result.

## Error Handling

- If `placeholder.name` does not match any picture shape in the template, the API returns HTTP 400 with a list of available picture shape names.
- If `blob.id` does not exist in container `temp`, the flow fails at the blob retrieval step.

## Example

**Agent action:**
1. User provides an image (e.g., uploads `car.jpg`).
2. Agent uploads `car.jpg` to container `temp` as `img_9f3a.jpg`.
3. Agent calls this flow with:
   - `blob.id`: `img_9f3a.jpg`
   - `placeholder.name`: `Slide1_Picture1`
   - `conversation.id`: `conv-9f3a`
4. Flow returns: `blob.id`: `result_9f3a.pptx`

## Supported Image Formats

PNG, JPEG, GIF, BMP, TIFF, SVG

## Constraints

- Only **one image** can be replaced per flow call. To replace multiple images, call this flow multiple times with different `placeholder.name` values.
- The image replaces the existing picture while **preserving the original position and dimensions** on the slide.
