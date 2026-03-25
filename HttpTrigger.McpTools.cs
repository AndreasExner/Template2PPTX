using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Extensions.Mcp;
using Microsoft.Extensions.Logging;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Template2PPTX.Function;

/// <summary>
/// MCP tool endpoints for Template2PPTX.
/// Same partial class as HttpTrigger so all private processing methods are shared.
/// </summary>
public partial class HttpTrigger
{
    // =====================================================================
    // MCP Tool: list_shapes
    // =====================================================================

    [Function("ListShapesMcp")]
    public string ListShapesMcp(
        [McpToolTrigger("list_shapes",
            "Lists all named shapes across all slides in a PPTX template. " +
            "Returns shape names, types, and slide numbers. Use this to discover " +
            "available shapes before calling fill_template, fill_chart, fill_table, or replace_images.")]
        ToolInvocationContext context,
        [McpToolProperty("templateBase64",
            "Base64-encoded PPTX template file.", true)]
        string templateBase64)
    {
        _logger.LogInformation("MCP list_shapes triggered.");
        try
        {
            var bytes = Convert.FromBase64String(templateBase64);
            using var ms = new MemoryStream(bytes);
            using var pptx = PresentationDocument.Open(ms, false);

            var shapes = CollectAllShapes(pptx);
            return JsonSerializer.Serialize(new { success = true, shapes },
                new JsonSerializerOptions { WriteIndented = true });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, error = ex.Message });
        }
    }

    private static List<ShapeInfo> CollectAllShapes(PresentationDocument pptx)
    {
        var result = new List<ShapeInfo>();
        var presentationPart = pptx.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList is null) return result;

        int slideNumber = 0;
        foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
        {
            slideNumber++;
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            if (slidePart.Slide is null) continue;

            // Regular shapes
            foreach (var shape in slidePart.Slide.Descendants<Shape>())
            {
                var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
                if (name is not null)
                    result.Add(new ShapeInfo { Name = name, Type = "TextShape", Slide = slideNumber });
            }

            // Picture shapes
            foreach (var pic in slidePart.Slide.Descendants<Picture>())
            {
                var name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                if (name is not null)
                    result.Add(new ShapeInfo { Name = name, Type = "Picture", Slide = slideNumber });
            }

            // GraphicFrame shapes (tables, charts)
            foreach (var gf in slidePart.Slide.Descendants<GraphicFrame>())
            {
                var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;
                if (name is null) continue;

                string type = "GraphicFrame";
                if (gf.Descendants<Drawing.Table>().Any())
                    type = "Table";
                else if (gf.Descendants<Drawing.GraphicData>()
                    .SelectMany(gd => gd.Elements())
                    .Any(el => el.LocalName == "chart"))
                    type = "Chart";

                result.Add(new ShapeInfo { Name = name, Type = type, Slide = slideNumber });
            }
        }
        return result;
    }

    private class ShapeInfo
    {
        public string Name { get; set; } = "";
        public string Type { get; set; } = "";
        public int Slide { get; set; }
    }

    // =====================================================================
    // MCP Tool: fill_template
    // =====================================================================

    [Function("FillTemplateMcp")]
    public string FillTemplateMcp(
        [McpToolTrigger("fill_template",
            "Replaces text placeholders in named shapes of a PPTX template. " +
            "Supports regular shapes, grouped shapes, and table cells. " +
            "Use list_shapes first to discover available shape names.")]
        ToolInvocationContext context,
        [McpToolProperty("templateBase64",
            "Base64-encoded PPTX template file.", true)]
        string templateBase64,
        [McpToolProperty("placeholders",
            "JSON object keyed by shape name. Each value is an object of placeholder-to-replacement pairs. " +
            "Example: {\"Title 1\":{\"{{Title}}\":\"Hello\"},\"Subtitle 2\":{\"{{Sub}}\":\"World\"}}", true)]
        string placeholders)
    {
        _logger.LogInformation("MCP fill_template triggered.");
        try
        {
            var shapeReplacements = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, string>>>(placeholders);
            if (shapeReplacements is null || shapeReplacements.Count == 0)
                return JsonSerializer.Serialize(new { success = false, error = "The 'placeholders' dictionary must not be empty." });

            var bytes = Convert.FromBase64String(templateBase64);
            using var ms = new MemoryStream(bytes);

            using (var pptx = PresentationDocument.Open(ms, isEditable: true))
            {
                var result = ReplacePlaceholders(pptx, shapeReplacements);
                if (result.NotFound.Count > 0)
                    _logger.LogWarning("MCP fill_template: shapes not found: {Shapes}", string.Join(", ", result.NotFound));
            }

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Template filled successfully.",
                fileName = "filled.pptx",
                contentBase64 = Convert.ToBase64String(ms.ToArray())
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, error = ex.Message });
        }
    }

    // =====================================================================
    // MCP Tool: fill_chart
    // =====================================================================

    [Function("FillChartMcp")]
    public string FillChartMcp(
        [McpToolTrigger("fill_chart",
            "Fills a chart in a PPTX template with new categories and series data. " +
            "Updates both the chart XML cache and the embedded Excel workbook. " +
            "Use list_shapes first to find chart shape names.")]
        ToolInvocationContext context,
        [McpToolProperty("templateBase64",
            "Base64-encoded PPTX template file.", true)]
        string templateBase64,
        [McpToolProperty("chartData",
            "JSON object with chart data. Must contain 'shapeName' (string), " +
            "'categories' (string array), and 'series' (array of {name, values}). " +
            "Example: {\"shapeName\":\"Chart 1\",\"categories\":[\"Q1\",\"Q2\"]," +
            "\"series\":[{\"name\":\"Sales\",\"values\":[10.5,20.3]}]}", true)]
        string chartData)
    {
        _logger.LogInformation("MCP fill_chart triggered.");
        try
        {
            var data = JsonSerializer.Deserialize<ChartDataInput>(chartData,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            if (data is null || data.Series.Count == 0 || data.Categories.Count == 0)
                return JsonSerializer.Serialize(new { success = false, error = "chartData must contain 'categories' and at least one 'series'." });
            if (string.IsNullOrWhiteSpace(data.ShapeName))
                return JsonSerializer.Serialize(new { success = false, error = "chartData must contain 'shapeName'." });

            var bytes = Convert.FromBase64String(templateBase64);
            using var ms = new MemoryStream(bytes);

            using (var pptx = PresentationDocument.Open(ms, isEditable: true))
            {
                var chartMap = GetAllChartPartsByShapeName(pptx);
                if (chartMap.Count == 0)
                    return JsonSerializer.Serialize(new { success = false, error = "No charts found in the presentation." });

                if (!chartMap.TryGetValue(data.ShapeName, out var chartPart))
                {
                    var available = string.Join(", ", chartMap.Keys.Select(k => $"'{k}'"));
                    return JsonSerializer.Serialize(new { success = false, error = $"No chart with shapeName '{data.ShapeName}' found. Available: {available}" });
                }

                UpdateChartCache(chartPart, data);
                UpdateEmbeddedExcel(chartPart, data);
            }

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Chart '{data.ShapeName}' filled with {data.Categories.Count} categories and {data.Series.Count} series.",
                fileName = "chart-filled.pptx",
                contentBase64 = Convert.ToBase64String(ms.ToArray())
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, error = ex.Message });
        }
    }

    // =====================================================================
    // MCP Tool: fill_table
    // =====================================================================

    [Function("FillTableMcp")]
    public string FillTableMcp(
        [McpToolTrigger("fill_table",
            "Fills a table in a PPTX template with dynamic data rows. " +
            "The template defines the column count. Header placeholders can be replaced. " +
            "If the template table has a second row, its formatting is used for all data rows. " +
            "Use list_shapes first to find table shape names.")]
        ToolInvocationContext context,
        [McpToolProperty("templateBase64",
            "Base64-encoded PPTX template file.", true)]
        string templateBase64,
        [McpToolProperty("tableData",
            "JSON object with table data. Must contain 'shapeName' (string) and 'rows' (array of string arrays). " +
            "Optionally contains 'headers' (object of placeholder-to-value mappings). " +
            "Example: {\"shapeName\":\"Table 1\",\"headers\":{\"{{Col1}}\":\"Name\",\"{{Col2}}\":\"Score\"}," +
            "\"rows\":[[\"Alice\",\"42\"],[\"Bob\",\"17\"]]}", true)]
        string tableData)
    {
        _logger.LogInformation("MCP fill_table triggered.");
        try
        {
            var data = JsonSerializer.Deserialize<TableDataInput>(tableData,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            if (data is null || data.Rows.Count == 0)
                return JsonSerializer.Serialize(new { success = false, error = "tableData must contain at least one row in 'rows'." });
            if (string.IsNullOrWhiteSpace(data.ShapeName))
                return JsonSerializer.Serialize(new { success = false, error = "tableData must contain 'shapeName'." });

            var bytes = Convert.FromBase64String(templateBase64);
            using var ms = new MemoryStream(bytes);

            using (var pptx = PresentationDocument.Open(ms, isEditable: true))
            {
                var tableResult = FillTableInPresentation(pptx, data);
                if (!tableResult.Success)
                    return JsonSerializer.Serialize(new { success = false, error = tableResult.ErrorMessage });
            }

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = $"Table '{data.ShapeName}' filled with {data.Rows.Count} data rows.",
                fileName = "table-filled.pptx",
                contentBase64 = Convert.ToBase64String(ms.ToArray())
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, error = ex.Message });
        }
    }

    // =====================================================================
    // MCP Tool: replace_images
    // =====================================================================

    [Function("ReplaceImagesMcp")]
    public string ReplaceImagesMcp(
        [McpToolTrigger("replace_images",
            "Replaces images in picture shapes of a PPTX template. " +
            "Each image is identified by its shape name and provided as Base64-encoded data. " +
            "Position and dimensions of the original shape are preserved. " +
            "Use list_shapes first to find picture shape names.")]
        ToolInvocationContext context,
        [McpToolProperty("templateBase64",
            "Base64-encoded PPTX template file.", true)]
        string templateBase64,
        [McpToolProperty("images",
            "JSON object mapping shape names to image data. " +
            "Each value has 'contentType' (e.g. 'image/png') and 'dataBase64' (Base64-encoded image). " +
            "Example: {\"Picture 3\":{\"contentType\":\"image/png\",\"dataBase64\":\"iVBORw0KGgo...\"}}", true)]
        string images)
    {
        _logger.LogInformation("MCP replace_images triggered.");
        try
        {
            var imageMap = JsonSerializer.Deserialize<Dictionary<string, ImageInput>>(images,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            if (imageMap is null || imageMap.Count == 0)
                return JsonSerializer.Serialize(new { success = false, error = "The 'images' dictionary must not be empty." });

            var bytes = Convert.FromBase64String(templateBase64);
            using var ms = new MemoryStream(bytes);

            using (var pptx = PresentationDocument.Open(ms, isEditable: true))
            {
                var presentationPart = pptx.PresentationPart;
                if (presentationPart?.Presentation?.SlideIdList is null)
                    return JsonSerializer.Serialize(new { success = false, error = "The presentation contains no slides." });

                var matched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
                {
                    var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
                    if (slidePart.Slide is null) continue;

                    foreach (var pic in slidePart.Slide.Descendants<Picture>())
                    {
                        var shapeName = pic.NonVisualPictureProperties?
                            .NonVisualDrawingProperties?.Name?.Value;
                        if (shapeName is null || !imageMap.TryGetValue(shapeName, out var imageInput))
                            continue;

                        matched.Add(shapeName);

                        var blipFill = pic.BlipFill;
                        var blip = blipFill?.Blip;
                        if (blip?.Embed?.Value is null) continue;

                        string contentType = imageInput.ContentType?.ToLowerInvariant() ?? "";
                        var partType = contentType switch
                        {
                            "image/png" => ImagePartType.Png,
                            "image/gif" => ImagePartType.Gif,
                            "image/bmp" => ImagePartType.Bmp,
                            "image/tiff" => ImagePartType.Tiff,
                            "image/svg+xml" => ImagePartType.Svg,
                            _ => ImagePartType.Jpeg
                        };

                        var newImagePart = slidePart.AddImagePart(partType);
                        var imageBytes = Convert.FromBase64String(imageInput.DataBase64 ?? "");
                        using (var imgStream = new MemoryStream(imageBytes))
                        {
                            newImagePart.FeedData(imgStream);
                        }

                        blip.Embed = slidePart.GetIdOfPart(newImagePart);

                        var srcRect = blipFill.GetFirstChild<Drawing.SourceRectangle>();
                        srcRect?.Remove();

                        _logger.LogInformation("MCP replace_images: shape '{Shape}' replaced ({ContentType}).",
                            shapeName, imageInput.ContentType);
                    }
                }

                var notFound = imageMap.Keys.Where(k => !matched.Contains(k)).ToList();
                if (notFound.Count > 0)
                    _logger.LogWarning("MCP replace_images: shapes not found: {NotFound}", string.Join(", ", notFound));
            }

            return JsonSerializer.Serialize(new
            {
                success = true,
                message = "Images replaced successfully.",
                fileName = "images-replaced.pptx",
                contentBase64 = Convert.ToBase64String(ms.ToArray())
            });
        }
        catch (Exception ex)
        {
            return JsonSerializer.Serialize(new { success = false, error = ex.Message });
        }
    }

    private class ImageInput
    {
        public string? ContentType { get; set; }
        public string? DataBase64 { get; set; }
    }
}
