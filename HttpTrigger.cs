using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Logging;
using Charts = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace Template2PPTX.Function;

public class HttpTrigger
{
    private readonly ILogger<HttpTrigger> _logger;

    public HttpTrigger(ILogger<HttpTrigger> logger)
    {
        _logger = logger;
    }

    /// <summary>
    /// Accepts a multipart/form-data request with:
    ///   - "template" : the .pptx template file
    ///   - "placeholders" : JSON object keyed by shape name, e.g.
    ///     { "Title 1": {"{{Title}}":"Hello"}, "Subtitle 2": {"{{Sub}}":"World"} }
    /// Returns the modified .pptx file.
    /// </summary>
    [Function("FillTemplate")]
    public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
    {
        _logger.LogInformation("FillTemplate function triggered.");

        // --- 1. Read the uploaded template file ---
        var form = await req.ReadFormAsync();
        var templateFile = form.Files.GetFile("template");
        if (templateFile is null || templateFile.Length == 0)
        {
            return new BadRequestObjectResult("Please upload a .pptx file as form field 'template'.");
        }

        // --- 2. Read placeholder dictionary from form field ---
        string? placeholdersJson = form["placeholders"];
        if (string.IsNullOrWhiteSpace(placeholdersJson))
        {
            return new BadRequestObjectResult(
                "Please provide a 'placeholders' form field with JSON, e.g. " +
                "{\"Title 1\":{\"{{Title}}\":\"Hello\"}}");
        }

        Dictionary<string, Dictionary<string, string>>? shapeReplacements;
        try
        {
            shapeReplacements = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, string>>>(placeholdersJson);
        }
        catch (JsonException ex)
        {
            return new BadRequestObjectResult($"Invalid JSON in 'placeholders': {ex.Message}");
        }

        if (shapeReplacements is null || shapeReplacements.Count == 0)
        {
            return new BadRequestObjectResult("The 'placeholders' dictionary must not be empty.");
        }

        // --- 3. Copy template into a writable MemoryStream ---
        var ms = new MemoryStream();
        await templateFile.CopyToAsync(ms);
        _logger.LogInformation("Template received: {Length} bytes, ContentType: {ContentType}",
            ms.Length, templateFile.ContentType);
        ms.Position = 0;

        // --- 4. Open PPTX and replace text ---
        try
        {
            using (var pptx = PresentationDocument.Open(ms, isEditable: true))
            {
                var result = ReplacePlaceholders(pptx, shapeReplacements);
                if (result.NotFound.Count > 0)
                {
                    _logger.LogWarning("Shapes not found: {Shapes}", string.Join(", ", result.NotFound));
                }
            }
        }
        catch (System.IO.FileFormatException ex)
        {
            _logger.LogError(ex, "Failed to open PPTX. First 4 bytes: {Header}",
                BitConverter.ToString(ms.Length >= 4 ? ms.GetBuffer()[..4] : ms.GetBuffer()));
            return new BadRequestObjectResult(
                $"The uploaded file is not a valid PPTX. Ensure you're uploading a real .pptx file. Error: {ex.Message}");
        }

        // --- 5. Return the modified file ---
        ms.Position = 0;
        return new FileStreamResult(ms,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        {
            FileDownloadName = "filled.pptx"
        };
    }

    /// <summary>
    /// Result of placeholder replacement, tracking which shapes were not found.
    /// </summary>
    private class ReplaceResult
    {
        public List<string> NotFound { get; set; } = new();
    }

    /// <summary>
    /// Iterates over every slide and replaces placeholder strings in shapes
    /// identified by their shape name (visible in PowerPoint's Selection Pane).
    /// </summary>
    private ReplaceResult ReplacePlaceholders(
        PresentationDocument pptx, Dictionary<string, Dictionary<string, string>> shapeReplacements)
    {
        var result = new ReplaceResult();
        var presentationPart = pptx.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList is null) return result;

        // Build a case-insensitive lookup: shapeName → replacements
        var lookup = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (shapeName, replacements) in shapeReplacements)
            lookup[shapeName] = replacements;

        var matched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            if (slidePart.Slide is null) continue;

            // Process regular shapes (sp)
            foreach (var shape in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Shape>())
            {
                var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
                if (name is null || !lookup.TryGetValue(name, out var replacements)) continue;

                matched.Add(name);
                ReplaceTextInElement(shape, replacements);
            }

            // Process grouped shapes
            foreach (var groupShape in slidePart.Slide.Descendants<GroupShape>())
            {
                foreach (var shape in groupShape.Descendants<DocumentFormat.OpenXml.Presentation.Shape>())
                {
                    var name = shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value;
                    if (name is null || !lookup.TryGetValue(name, out var replacements)) continue;

                    matched.Add(name);
                    ReplaceTextInElement(shape, replacements);
                }
            }

            // Process table cells inside graphic frames
            foreach (var gf in slidePart.Slide.Descendants<GraphicFrame>())
            {
                var name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;
                if (name is null || !lookup.TryGetValue(name, out var replacements)) continue;

                matched.Add(name);
                ReplaceTextInElement(gf, replacements);
            }

            slidePart.Slide.Save();
        }

        // Report shapes that were requested but not found
        result.NotFound = lookup.Keys.Where(k => !matched.Contains(k)).ToList();
        return result;
    }

    /// <summary>
    /// Replaces placeholder text in all Drawing.Text descendants of a given element.
    /// </summary>
    private void ReplaceTextInElement(OpenXmlElement element, Dictionary<string, string> replacements)
    {
        foreach (var textElement in element.Descendants<Drawing.Text>())
        {
            foreach (var (placeholder, replacement) in replacements)
            {
                if (textElement.Text?.Contains(placeholder) == true)
                {
                    _logger.LogInformation("Replacing '{Placeholder}' → '{Replacement}'",
                        placeholder, replacement);
                    textElement.Text = textElement.Text.Replace(placeholder, replacement);
                }
            }
        }
    }

    // =====================================================================
    // ReplaceImages endpoint
    // =====================================================================

    /// <summary>
    /// Accepts a multipart/form-data request with:
    ///   - "template" : the .pptx template file
    ///   - one or more image files whose form-field name matches the shape name
    ///     in the presentation (visible in PowerPoint's Selection Pane).
    /// The image inside each matched shape is replaced while keeping the
    /// exact same position and dimensions on the slide.
    /// Returns the modified .pptx file.
    /// 
    /// Example (curl):
    ///   curl -X POST .../api/ReplaceImages \
    ///     -F "template=@deck.pptx" \
    ///     -F "Picture 3=@logo.png" \
    ///     -F "Picture 5=@photo.jpg"
    /// </summary>
    [Function("ReplaceImages")]
    public async Task<IActionResult> RunReplaceImages(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
    {
        _logger.LogInformation("ReplaceImages function triggered.");

        var form = await req.ReadFormAsync();
        var templateFile = form.Files.GetFile("template");
        if (templateFile is null || templateFile.Length == 0)
            return new BadRequestObjectResult("Please upload a .pptx file as form field 'template'.");

        // Collect all uploaded image files (every field except "template")
        var imageFiles = form.Files.Where(f => !f.Name.Equals("template", StringComparison.OrdinalIgnoreCase)).ToList();
        if (imageFiles.Count == 0)
            return new BadRequestObjectResult(
                "Please provide at least one image file whose form-field name matches a shape name in the presentation.");

        var ms = new MemoryStream();
        await templateFile.CopyToAsync(ms);
        ms.Position = 0;

        try
        {
            using var pptx = PresentationDocument.Open(ms, isEditable: true);
            var presentationPart = pptx.PresentationPart;
            if (presentationPart?.Presentation?.SlideIdList is null)
                return new BadRequestObjectResult("The presentation contains no slides.");

            // Build a case-insensitive set of requested shape names → image data
            var imageMap = new Dictionary<string, IFormFile>(StringComparer.OrdinalIgnoreCase);
            foreach (var file in imageFiles)
                imageMap[file.Name] = file;

            var matched = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
            {
                var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
                if (slidePart.Slide is null) continue;

                // Find all Picture shapes (p:pic)
                foreach (var pic in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>())
                {
                    var shapeName = pic.NonVisualPictureProperties?
                        .NonVisualDrawingProperties?.Name?.Value;
                    if (shapeName is null || !imageMap.TryGetValue(shapeName, out var newImageFile))
                        continue;

                    matched.Add(shapeName);

                    // Get the blipFill → blip → r:embed relationship id
                    var blipFill = pic.BlipFill;
                    var blip = blipFill?.Blip;
                    if (blip?.Embed?.Value is null)
                    {
                        _logger.LogWarning("Shape '{Shape}': no embedded image reference found — skipping.", shapeName);
                        continue;
                    }

                    // Determine image content type
                    string contentType = newImageFile.ContentType?.ToLowerInvariant() ?? "";
                    var partType = contentType switch
                    {
                        "image/png" => ImagePartType.Png,
                        "image/gif" => ImagePartType.Gif,
                        "image/bmp" => ImagePartType.Bmp,
                        "image/tiff" => ImagePartType.Tiff,
                        "image/svg+xml" => ImagePartType.Svg,
                        _ => ImagePartType.Jpeg // default to JPEG
                    };

                    // Create a new image part and feed the uploaded image data
                    var newImagePart = slidePart.AddImagePart(partType);
                    using (var imgStream = newImageFile.OpenReadStream())
                    {
                        newImagePart.FeedData(imgStream);
                    }

                    // Point the blip to the new image part
                    blip.Embed = slidePart.GetIdOfPart(newImagePart);

                    // Remove any stretch/crop transforms on the blipFill so the
                    // image fills the entire shape bounds without distortion artifacts.
                    // The shape's spPr (extent/offset) stays untouched → same size & position.
                    var srcRect = blipFill.GetFirstChild<Drawing.SourceRectangle>();
                    srcRect?.Remove();

                    _logger.LogInformation(
                        "Shape '{Shape}': image replaced ({ContentType}, {Size} bytes).",
                        shapeName, newImageFile.ContentType, newImageFile.Length);
                }
            }

            // Report unmatched shape names
            var notFound = imageMap.Keys.Where(k => !matched.Contains(k)).ToList();
            if (notFound.Count > 0)
            {
                // Collect all picture shape names for the error message
                var allPicNames = new List<string>();
                foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
                {
                    var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
                    if (slidePart.Slide is null) continue;
                    foreach (var pic in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Presentation.Picture>())
                    {
                        var n = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value;
                        if (n is not null) allPicNames.Add(n);
                    }
                }
                var available = allPicNames.Count > 0
                    ? string.Join(", ", allPicNames.Select(n => $"'{n}'"))
                    : "(none)";
                _logger.LogWarning(
                    "Image shapes not found: {NotFound}. Available picture shapes: {Available}",
                    string.Join(", ", notFound), available);
            }
        }
        catch (System.IO.FileFormatException ex)
        {
            return new BadRequestObjectResult($"Invalid PPTX file: {ex.Message}");
        }

        ms.Position = 0;
        return new FileStreamResult(ms,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        {
            FileDownloadName = "images-replaced.pptx"
        };
    }

    // =====================================================================
    // FillChart endpoint
    // =====================================================================

    /// <summary>
    /// Input model for chart data.
    /// Example JSON:
    /// {
    ///   "shapeName": "Diagramm 1",
    ///   "categories": ["Q1", "Q2", "Q3", "Q4"],
    ///   "series": [
    ///     { "name": "Sales", "values": [10.5, 20.3, 15.0, 25.7] }
    ///   ]
    /// }
    /// </summary>
    public class ChartDataInput
    {
        public string ShapeName { get; set; } = "";
        public List<string> Categories { get; set; } = new();
        public List<SeriesInput> Series { get; set; } = new();
    }

    public class SeriesInput
    {
        public string Name { get; set; } = "";
        public List<double> Values { get; set; } = new();
    }

    /// <summary>
    /// Accepts a multipart/form-data request with:
    ///   - "template" : the .pptx template file
    ///   - "chartData" : JSON with categories and series values
    /// Updates both the chart XML cache and the embedded Excel workbook.
    /// Returns the modified .pptx file.
    /// </summary>
    [Function("FillChart")]
    public async Task<IActionResult> RunFillChart(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
    {
        _logger.LogInformation("FillChart function triggered.");

        var form = await req.ReadFormAsync();
        var templateFile = form.Files.GetFile("template");
        if (templateFile is null || templateFile.Length == 0)
            return new BadRequestObjectResult("Please upload a .pptx file as form field 'template'.");

        string? chartDataJson = form["chartData"];
        if (string.IsNullOrWhiteSpace(chartDataJson))
            return new BadRequestObjectResult("Please provide a 'chartData' form field with JSON.");

        ChartDataInput? chartData;
        try
        {
            chartData = JsonSerializer.Deserialize<ChartDataInput>(chartDataJson,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch (JsonException ex)
        {
            return new BadRequestObjectResult($"Invalid JSON in 'chartData': {ex.Message}");
        }

        if (chartData is null || chartData.Series.Count == 0 || chartData.Categories.Count == 0)
            return new BadRequestObjectResult("chartData must contain 'categories' and at least one 'series'.");

        if (string.IsNullOrWhiteSpace(chartData.ShapeName))
            return new BadRequestObjectResult("chartData must contain 'shapeName' (the name of the chart shape in the slide).");

        var ms = new MemoryStream();
        await templateFile.CopyToAsync(ms);
        ms.Position = 0;

        try
        {
            using var pptx = PresentationDocument.Open(ms, isEditable: true);
            var chartMap = GetAllChartPartsByShapeName(pptx);

            if (chartMap.Count == 0)
                return new BadRequestObjectResult("No charts found in the presentation.");

            if (!chartMap.TryGetValue(chartData.ShapeName, out var chartPart))
            {
                var available = string.Join(", ", chartMap.Keys.Select(k => $"'{k}'"));
                return new BadRequestObjectResult(
                    $"No chart with shapeName '{chartData.ShapeName}' found. Available: {available}");
            }

            // 1. Update the chart XML cache
            UpdateChartCache(chartPart, chartData);

            // 2. Update the embedded Excel workbook
            UpdateEmbeddedExcel(chartPart, chartData);

            _logger.LogInformation("Chart '{ShapeName}' updated with {CatCount} categories and {SerCount} series.",
                chartData.ShapeName, chartData.Categories.Count, chartData.Series.Count);
        }
        catch (System.IO.FileFormatException ex)
        {
            return new BadRequestObjectResult($"Invalid PPTX file: {ex.Message}");
        }

        ms.Position = 0;
        return new FileStreamResult(ms,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        {
            FileDownloadName = "chart-filled.pptx"
        };
    }

    /// <summary>
    /// Collects all ChartParts across all slides, keyed by their shape name
    /// (the name visible in PowerPoint's Selection Pane).
    /// </summary>
    private static Dictionary<string, ChartPart> GetAllChartPartsByShapeName(PresentationDocument pptx)
    {
        var result = new Dictionary<string, ChartPart>(StringComparer.OrdinalIgnoreCase);
        var presentationPart = pptx.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList is null) return result;

        foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            if (slidePart.Slide is null) continue;

            // Find all graphic frames that contain a chart relationship
            var graphicFrames = slidePart.Slide.Descendants<GraphicFrame>().ToList();
            foreach (var gf in graphicFrames)
            {
                // The shape name is in the NonVisualGraphicFrameProperties
                var nvProps = gf.NonVisualGraphicFrameProperties;
                var shapeName = nvProps?.NonVisualDrawingProperties?.Name?.Value ?? "";

                // Find the chart relationship ID inside the graphic data
                var chartRef = gf.Descendants<Drawing.GraphicData>()
                    .SelectMany(gd => gd.Elements())
                    .Where(el => el.LocalName == "chart")
                    .Select(el => el.GetAttribute("id",
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value)
                    .FirstOrDefault();

                if (!string.IsNullOrEmpty(chartRef))
                {
                    try
                    {
                        var part = slidePart.GetPartById(chartRef);
                        if (part is ChartPart chartPart && !string.IsNullOrEmpty(shapeName))
                        {
                            result[shapeName] = chartPart;
                        }
                    }
                    catch { /* relationship not found — skip */ }
                }
            }
        }
        return result;
    }

    /// <summary>
    /// Updates the cached category labels, series values, AND formula references
    /// inside the chart XML.  Works for Bar, Line, Pie, Pie3D, Area, etc.
    /// </summary>
    private void UpdateChartCache(ChartPart chartPart, ChartDataInput chartData)
    {
        var chartSpace = chartPart.ChartSpace;
        var plotArea = chartSpace.Descendants<Charts.PlotArea>().FirstOrDefault()
            ?? throw new InvalidOperationException("No PlotArea found in chart.");

        // Find all series elements (works for any chart type)
        var seriesElements = plotArea.Descendants<Charts.Index>()
            .Select(idx => idx.Parent)
            .Where(p => p is not null)
            .OfType<OpenXmlCompositeElement>()
            .Where(el => el.LocalName.EndsWith("Ser") || el.LocalName == "ser")
            .OrderBy(el => el.Descendants<Charts.Order>().FirstOrDefault()?.Val?.Value ?? 0)
            .ToList();

        uint lastDataRow = (uint)(chartData.Categories.Count + 1); // +1 for header

        for (int s = 0; s < Math.Min(seriesElements.Count, chartData.Series.Count); s++)
        {
            var serElement = seriesElements[s];
            var seriesInput = chartData.Series[s];
            string serCol = GetColumnLetter(s + 1); // B, C, D …

            // --- Series name cache + formula ---
            var txStrCache = serElement.Descendants<Charts.SeriesText>()
                .SelectMany(st => st.Descendants<Charts.StringCache>())
                .FirstOrDefault();
            if (txStrCache is not null)
                UpdateStringCache(txStrCache, new[] { seriesInput.Name });

            var txFormula = serElement.Descendants<Charts.SeriesText>()
                .SelectMany(st => st.Descendants<Charts.Formula>())
                .FirstOrDefault();
            if (txFormula is not null)
            {
                var sn = ExtractSheetPrefix(txFormula.Text);
                txFormula.Text = $"{sn}!${serCol}$1";
            }

            // --- Category cache + formula ---
            var catStrCache = serElement.Descendants<Charts.CategoryAxisData>()
                .SelectMany(c => c.Descendants<Charts.StringCache>())
                .FirstOrDefault();
            if (catStrCache is not null)
                UpdateStringCache(catStrCache, chartData.Categories.ToArray());

            // Categories might also live in a numRef (numeric categories)
            var catNumCache = serElement.Descendants<Charts.CategoryAxisData>()
                .SelectMany(c => c.Descendants<Charts.NumberingCache>())
                .FirstOrDefault();
            if (catNumCache is not null)
                UpdateNumberCache(catNumCache, chartData.Categories.Select(c => double.TryParse(c, out var v) ? v : 0).ToArray());

            var catFormula = serElement.Descendants<Charts.CategoryAxisData>()
                .SelectMany(c => c.Descendants<Charts.Formula>())
                .FirstOrDefault();
            if (catFormula is not null)
            {
                var sn = ExtractSheetPrefix(catFormula.Text);
                catFormula.Text = $"{sn}!$A$2:$A${lastDataRow}";
            }

            // --- Values cache + formula ---
            var numCache = serElement.Descendants<Charts.Values>()
                .SelectMany(v => v.Descendants<Charts.NumberingCache>())
                .FirstOrDefault();
            if (numCache is not null)
                UpdateNumberCache(numCache, seriesInput.Values.ToArray());

            var valFormula = serElement.Descendants<Charts.Values>()
                .SelectMany(v => v.Descendants<Charts.Formula>())
                .FirstOrDefault();
            if (valFormula is not null)
            {
                var sn = ExtractSheetPrefix(valFormula.Text);
                valFormula.Text = $"{sn}!${serCol}$2:${serCol}${lastDataRow}";
            }
        }

        chartSpace.Save();
    }

    /// <summary>
    /// Extracts the sheet-name prefix from a chart formula, e.g.
    /// "Sheet1!$A$2:$A$5" → "Sheet1",  "'Tabelle 1'!$B$1" → "'Tabelle 1'".
    /// Returns the prefix as-is (with quotes if present) so it can be reused.
    /// </summary>
    private static string ExtractSheetPrefix(string? formula)
    {
        if (string.IsNullOrEmpty(formula) || !formula.Contains('!'))
            return "Sheet1";
        return formula.Split('!')[0];
    }

    /// <summary>
    /// Returns the clean sheet name (without surrounding single-quotes).
    /// </summary>
    private static string CleanSheetName(string prefix)
    {
        if (prefix.StartsWith('\'') && prefix.EndsWith('\''))
            return prefix[1..^1];
        return prefix;
    }

    private static void UpdateStringCache(Charts.StringCache cache, string[] values)
    {
        // Remove existing points
        cache.RemoveAllChildren<Charts.PointCount>();
        cache.RemoveAllChildren<Charts.StringPoint>();

        cache.AppendChild(new Charts.PointCount { Val = (uint)values.Length });
        for (int i = 0; i < values.Length; i++)
        {
            cache.AppendChild(new Charts.StringPoint(new Charts.NumericValue(values[i]))
            {
                Index = (uint)i
            });
        }
    }

    private static void UpdateNumberCache(Charts.NumberingCache cache, double[] values)
    {
        var formatCode = cache.GetFirstChild<Charts.FormatCode>()?.Text ?? "General";

        cache.RemoveAllChildren<Charts.FormatCode>();
        cache.RemoveAllChildren<Charts.PointCount>();
        cache.RemoveAllChildren<Charts.NumericPoint>();

        cache.AppendChild(new Charts.FormatCode(formatCode));
        cache.AppendChild(new Charts.PointCount { Val = (uint)values.Length });
        for (int i = 0; i < values.Length; i++)
        {
            cache.AppendChild(new Charts.NumericPoint(
                new Charts.NumericValue(values[i].ToString(System.Globalization.CultureInfo.InvariantCulture)))
            {
                Index = (uint)i
            });
        }
    }

    /// <summary>
    /// Creates a brand-new embedded Excel workbook from scratch and feeds it
    /// back into the chart's EmbeddedPackagePart.  This avoids any corruption
    /// from modifying the existing file (SharedStrings, Table definitions, etc.).
    /// Column A = categories, Column B+ = series values.
    /// </summary>
    private void UpdateEmbeddedExcel(ChartPart chartPart, ChartDataInput chartData)
    {
        var embeddedPart = chartPart.EmbeddedPackagePart;
        if (embeddedPart is null)
        {
            _logger.LogWarning("No embedded Excel found in chart — skipping Excel update.");
            return;
        }

        // Determine the sheet name the chart formulas reference
        string sheetPrefix = "Sheet1";
        var anyFormula = chartPart.ChartSpace.Descendants<Charts.Formula>().FirstOrDefault();
        if (anyFormula?.Text is not null && anyFormula.Text.Contains('!'))
            sheetPrefix = anyFormula.Text.Split('!')[0];
        string sheetName = CleanSheetName(sheetPrefix);

        using var excelStream = new MemoryStream();

        // ---- Build a minimal, clean XLSX from scratch ----
        using (var excel = SpreadsheetDocument.Create(excelStream, SpreadsheetDocumentType.Workbook))
        {
            // -- WorkbookPart --
            var workbookPart = excel.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(
                new Sheets(
                    new Sheet { Id = "rId1", SheetId = 1, Name = sheetName }
                )
            );

            // -- Minimal StylesPart (required for a valid xlsx) --
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font()) { Count = 1 },
                new Fills(
                    new Fill(new PatternFill { PatternType = PatternValues.None }),
                    new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
                ) { Count = 2 },
                new Borders(new Border()) { Count = 1 },
                new CellFormats(new CellFormat()) { Count = 1 }
            );
            stylesPart.Stylesheet.Save();

            // -- WorksheetPart --
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>("rId1");
            var sheetData = new SheetData();

            // Row 1: Header — empty cell + series names
            var headerRow = new Row { RowIndex = 1 };
            headerRow.AppendChild(CreateTextCell("A", 1, ""));
            for (int s = 0; s < chartData.Series.Count; s++)
            {
                string col = GetColumnLetter(s + 1);
                headerRow.AppendChild(CreateTextCell(col, 1, chartData.Series[s].Name));
            }
            sheetData.AppendChild(headerRow);

            // Data rows
            for (int r = 0; r < chartData.Categories.Count; r++)
            {
                uint rowIdx = (uint)(r + 2);
                var dataRow = new Row { RowIndex = rowIdx };
                dataRow.AppendChild(CreateTextCell("A", rowIdx, chartData.Categories[r]));
                for (int s = 0; s < chartData.Series.Count; s++)
                {
                    string col = GetColumnLetter(s + 1);
                    double val = r < chartData.Series[s].Values.Count ? chartData.Series[s].Values[r] : 0;
                    dataRow.AppendChild(CreateNumberCell(col, rowIdx, val));
                }
                sheetData.AppendChild(dataRow);
            }

            worksheetPart.Worksheet = new Worksheet(sheetData);
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }

        _logger.LogInformation("Created new embedded Excel (sheet={SheetName}): {Size} bytes",
            sheetName, excelStream.Length);

        excelStream.Position = 0;
        embeddedPart.FeedData(excelStream);
    }

    private static Cell CreateTextCell(string column, uint row, string text)
    {
        return new Cell
        {
            CellReference = $"{column}{row}",
            DataType = CellValues.String,
            CellValue = new CellValue(text)
        };
    }

    private static Cell CreateNumberCell(string column, uint row, double value)
    {
        return new Cell
        {
            CellReference = $"{column}{row}",
            DataType = CellValues.Number,
            CellValue = new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture))
        };
    }

    private static string GetColumnLetter(int index)
    {
        // 0=A, 1=B, 2=C, ...
        string result = "";
        index++; // 1-based
        while (index > 0)
        {
            index--;
            result = (char)('A' + index % 26) + result;
            index /= 26;
        }
        return result;
    }

    // =====================================================================
    // FillTable endpoint
    // =====================================================================

    /// <summary>
    /// Input model for table data.
    /// Example JSON:
    /// {
    ///   "shapeName": "Table 1",
    ///   "headers": { "{{Col1}}": "Name", "{{Col2}}": "Score" },
    ///   "rows": [
    ///     ["Alice", "42"],
    ///     ["Bob",   "17"]
    ///   ]
    /// }
    /// </summary>
    public class TableDataInput
    {
        public string ShapeName { get; set; } = "";
        public Dictionary<string, string>? Headers { get; set; }
        public List<List<string>> Rows { get; set; } = new();
    }

    /// <summary>
    /// Accepts a multipart/form-data request with:
    ///   - "template" : the .pptx template file
    ///   - "tableData" : JSON with shapeName, optional header replacements, and row data
    /// The template must contain a table (inside a GraphicFrame) whose shape name
    /// matches "shapeName".  The table should have a header row (row 0) and
    /// optionally a second "style-template" row (row 1) whose formatting will be
    /// cloned for every data row.  The column count is defined by the template.
    /// Returns the modified .pptx file.
    /// </summary>
    [Function("FillTable")]
    public async Task<IActionResult> RunFillTable(
        [HttpTrigger(AuthorizationLevel.Function, "post")] HttpRequest req)
    {
        _logger.LogInformation("FillTable function triggered.");

        var form = await req.ReadFormAsync();
        var templateFile = form.Files.GetFile("template");
        if (templateFile is null || templateFile.Length == 0)
            return new BadRequestObjectResult("Please upload a .pptx file as form field 'template'.");

        string? tableDataJson = form["tableData"];
        if (string.IsNullOrWhiteSpace(tableDataJson))
            return new BadRequestObjectResult("Please provide a 'tableData' form field with JSON.");

        TableDataInput? tableData;
        try
        {
            tableData = JsonSerializer.Deserialize<TableDataInput>(tableDataJson,
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
        }
        catch (JsonException ex)
        {
            return new BadRequestObjectResult($"Invalid JSON in 'tableData': {ex.Message}");
        }

        if (tableData is null || tableData.Rows.Count == 0)
            return new BadRequestObjectResult("tableData must contain at least one row in 'rows'.");

        if (string.IsNullOrWhiteSpace(tableData.ShapeName))
            return new BadRequestObjectResult("tableData must contain 'shapeName' (the name of the table shape in the slide).");

        var ms = new MemoryStream();
        await templateFile.CopyToAsync(ms);
        ms.Position = 0;

        try
        {
            using var pptx = PresentationDocument.Open(ms, isEditable: true);
            var tableResult = FillTableInPresentation(pptx, tableData);

            if (!tableResult.Success)
                return new BadRequestObjectResult(tableResult.ErrorMessage);

            _logger.LogInformation("Table '{ShapeName}' filled with {RowCount} data rows.",
                tableData.ShapeName, tableData.Rows.Count);
        }
        catch (System.IO.FileFormatException ex)
        {
            return new BadRequestObjectResult($"Invalid PPTX file: {ex.Message}");
        }

        ms.Position = 0;
        return new FileStreamResult(ms,
            "application/vnd.openxmlformats-officedocument.presentationml.presentation")
        {
            FileDownloadName = "table-filled.pptx"
        };
    }

    private class FillTableResult
    {
        public bool Success { get; set; }
        public string? ErrorMessage { get; set; }
    }

    /// <summary>
    /// Finds a table by shape name across all slides and fills it with dynamic rows.
    /// The header row (row 0) is kept; its placeholders are replaced if "headers" is provided.
    /// If the template table has a second row (row 1), it is used as a style template
    /// for all new data rows and then removed.  Otherwise rows are created with
    /// minimal formatting cloned from the first cell of the header row.
    /// </summary>
    private FillTableResult FillTableInPresentation(PresentationDocument pptx, TableDataInput tableData)
    {
        var presentationPart = pptx.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList is null)
            return new FillTableResult { ErrorMessage = "The presentation contains no slides." };

        Drawing.Table? targetTable = null;
        SlidePart? targetSlidePart = null;

        // Collect all table shape names for error reporting
        var allTableNames = new List<string>();

        foreach (var slideId in presentationPart.Presentation.SlideIdList.Elements<SlideId>())
        {
            var slidePart = (SlidePart)presentationPart.GetPartById(slideId.RelationshipId!);
            if (slidePart.Slide is null) continue;

            foreach (var gf in slidePart.Slide.Descendants<GraphicFrame>())
            {
                var shapeName = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value;

                // Check if this graphic frame contains a table (a:tbl)
                var table = gf.Descendants<Drawing.Table>().FirstOrDefault();
                if (table is null) continue;

                if (shapeName is not null) allTableNames.Add(shapeName);

                if (string.Equals(shapeName, tableData.ShapeName, StringComparison.OrdinalIgnoreCase))
                {
                    targetTable = table;
                    targetSlidePart = slidePart;
                    break;
                }
            }
            if (targetTable is not null) break;
        }

        if (targetTable is null)
        {
            var available = allTableNames.Count > 0
                ? string.Join(", ", allTableNames.Select(n => $"'{n}'"))
                : "(none)";
            return new FillTableResult
            {
                ErrorMessage = $"No table with shapeName '{tableData.ShapeName}' found. Available table shapes: {available}"
            };
        }

        var existingRows = targetTable.Elements<Drawing.TableRow>().ToList();
        if (existingRows.Count == 0)
            return new FillTableResult { ErrorMessage = "The table has no rows (not even a header row)." };

        int columnCount = targetTable.TableGrid?.Elements<Drawing.GridColumn>().Count() ?? 0;
        if (columnCount == 0)
            return new FillTableResult { ErrorMessage = "The table has no columns defined in its grid." };

        // --- 1. Replace header placeholders (row 0) ---
        var headerRow = existingRows[0];
        if (tableData.Headers is not null && tableData.Headers.Count > 0)
        {
            foreach (var textElement in headerRow.Descendants<Drawing.Text>())
            {
                foreach (var (placeholder, replacement) in tableData.Headers)
                {
                    if (textElement.Text?.Contains(placeholder) == true)
                    {
                        _logger.LogInformation("Table header: replacing '{Placeholder}' → '{Replacement}'",
                            placeholder, replacement);
                        textElement.Text = textElement.Text.Replace(placeholder, replacement);
                    }
                }
            }
        }

        // --- 2. Determine style template row ---
        // If row 1 exists, use it as style template; otherwise clone the header style
        Drawing.TableRow? styleTemplateRow = existingRows.Count > 1 ? existingRows[1] : null;

        // Remove all existing rows except the header
        for (int i = existingRows.Count - 1; i >= 1; i--)
        {
            existingRows[i].Remove();
        }

        // --- 3. Add data rows ---
        foreach (var rowData in tableData.Rows)
        {
            var newRow = CreateTableRow(
                styleTemplateRow ?? headerRow,
                rowData,
                columnCount,
                styleTemplateRow is not null);
            targetTable.AppendChild(newRow);
        }

        // --- 4. Update table height (optional: distribute row heights evenly) ---
        // The total table height is defined by the graphic frame's extent.
        // We distribute it evenly across header + data rows.
        // Row heights are in EMUs (English Metric Units).
        var totalRows = 1 + tableData.Rows.Count; // header + data rows
        long headerHeight = headerRow.Height?.Value ?? 370840L; // default ~1cm
        long templateRowHeight = styleTemplateRow?.Height?.Value ?? headerHeight;

        foreach (var row in targetTable.Elements<Drawing.TableRow>())
        {
            if (row == headerRow) continue;
            row.Height = templateRowHeight;
        }

        targetSlidePart!.Slide.Save();
        return new FillTableResult { Success = true };
    }

    /// <summary>
    /// Creates a new table row by cloning either the style-template row or
    /// the header row, then setting the cell text values.
    /// </summary>
    private static Drawing.TableRow CreateTableRow(
        Drawing.TableRow templateRow, List<string> cellValues, int columnCount, bool isStyleTemplate)
    {
        // Deep-clone the template row to preserve all formatting
        var newRow = (Drawing.TableRow)templateRow.CloneNode(deep: true);

        var cells = newRow.Elements<Drawing.TableCell>().ToList();

        for (int c = 0; c < columnCount; c++)
        {
            string value = c < cellValues.Count ? cellValues[c] : "";

            if (c < cells.Count)
            {
                // Replace all text in the cell with the new value
                SetCellText(cells[c], value);
            }
            else
            {
                // More columns in grid than cells in template row — add a new cell
                var newCell = c > 0 && cells.Count > 0
                    ? (Drawing.TableCell)cells[^1].CloneNode(deep: true)
                    : CreateMinimalTableCell();
                SetCellText(newCell, value);
                newRow.AppendChild(newCell);
            }
        }

        return newRow;
    }

    /// <summary>
    /// Sets the text of a table cell while preserving run-level formatting.
    /// Keeps the first paragraph and first run, removes all others,
    /// then sets the text of the remaining run.
    /// </summary>
    private static void SetCellText(Drawing.TableCell cell, string text)
    {
        var textBody = cell.GetFirstChild<Drawing.TextBody>();
        if (textBody is null)
        {
            textBody = new Drawing.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text(text))));
            cell.AppendChild(textBody);
            return;
        }

        var paragraphs = textBody.Elements<Drawing.Paragraph>().ToList();

        // Keep only the first paragraph
        for (int i = paragraphs.Count - 1; i >= 1; i--)
            paragraphs[i].Remove();

        var paragraph = paragraphs.FirstOrDefault();
        if (paragraph is null)
        {
            paragraph = new Drawing.Paragraph(new Drawing.Run(new Drawing.Text(text)));
            textBody.AppendChild(paragraph);
            return;
        }

        // Find or create a run
        var runs = paragraph.Elements<Drawing.Run>().ToList();

        // Remove all runs except the first
        for (int i = runs.Count - 1; i >= 1; i--)
            runs[i].Remove();

        var run = runs.FirstOrDefault();
        if (run is null)
        {
            run = new Drawing.Run(new Drawing.Text(text));
            paragraph.AppendChild(run);
        }
        else
        {
            var textEl = run.GetFirstChild<Drawing.Text>();
            if (textEl is not null)
                textEl.Text = text;
            else
                run.AppendChild(new Drawing.Text(text));
        }

        // Remove any explicit line breaks or fields that were in the template
        foreach (var br in paragraph.Elements<Drawing.Break>().ToList()) br.Remove();
        foreach (var fld in paragraph.Elements<Drawing.Field>().ToList()) fld.Remove();
    }

    /// <summary>
    /// Creates a minimal table cell with default formatting.
    /// </summary>
    private static Drawing.TableCell CreateMinimalTableCell()
    {
        return new Drawing.TableCell(
            new Drawing.TextBody(
                new Drawing.BodyProperties(),
                new Drawing.ListStyle(),
                new Drawing.Paragraph(new Drawing.Run(new Drawing.Text("")))),
            new Drawing.TableCellProperties());
    }
}