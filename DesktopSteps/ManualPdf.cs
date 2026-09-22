using System.Text;
using System.Text.RegularExpressions;
using System.Drawing.Imaging;

namespace DesktopSteps;

internal static class ManualPdf
{
    public static async Task CreateAsync(ExecutionPlan plan, string sessionDirectory, string outputPath, CancellationToken token)
    {
        var manualSteps = ManualSteps(plan.Steps);
        var eventsPath = Path.Combine(sessionDirectory, "recorded_events.json");
        var recordedEvents = File.Exists(eventsPath)
            ? await JsonFile.LoadAsync<List<RecordedEvent>>(eventsPath, token) ?? [] : [];
        // A compact PDF writer keeps the desktop app independent of external PDF tools.
        await using var output = File.Create(outputPath);
        var offsets = new List<long> { 0 };
        async Task Write(string value) => await output.WriteAsync(Encoding.ASCII.GetBytes(value), token);
        async Task Object(int id, string body)
        {
            offsets.Add(output.Position);
            await Write($"{id} 0 obj\n{body}\nendobj\n");
        }
        await Write("%PDF-1.4\n");
        var count = manualSteps.Count;
        var pageIds = Enumerable.Range(0, count).Select(i => 4 + 3 * i).ToArray();
        await Object(1, "<< /Type /Catalog /Pages 2 0 R >>");
        await Object(2, $"<< /Type /Pages /Count {count} /Kids [{string.Join(" ", pageIds.Select(x => $"{x} 0 R"))}] >>");
        await Object(3, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");
        for (var i = 0; i < count; i++)
        {
            var step = manualSteps[i];
            var pageId = pageIds[i];
            var imagePath = step.Screenshot is null ? null : Path.Combine(sessionDirectory, step.Screenshot);
            byte[]? image = imagePath is not null && File.Exists(imagePath)
                ? await File.ReadAllBytesAsync(imagePath, token) : null;
            if (image is not null && step.Action is "manual-menu" or "manual-menu-path" &&
                FindMenuPoint(step, recordedEvents) is { } point)
                image = CropMenuImage(image, point.X, point.Y);
            int width = 1, height = 1;
            if (image is not null)
            {
                using var bitmap = new Bitmap(new MemoryStream(image));
                width = bitmap.Width; height = bitmap.Height;
            }
            var fieldName = step.Target?.Label ?? (step.Target?.Window == "Find and Replace"
                ? step.Target.AutomationId == "18" ? "Find what" :
                  step.Target.AutomationId == "21" ? "Replace with" : null : null);
            var text = step.Action == "type"
                ? $"Step {step.Number}: Type text in {fieldName ?? step.Target?.Name ?? "the selected control"}"
                : step.Action == "manual-click"
                ? $"Step {step.Number}: Perform the recorded click manually"
                : step.Action == "manual-menu"
                ? $"Step {step.Number}: Choose {step.Value} from the open context menu"
                : step.Action == "manual-menu-path"
                ? $"Step {step.Number}: Open {step.Value}"
                : step.Action == "manual-paste"
                ? $"Step {step.Number}: Select {step.Target?.Name ?? "the destination"} and paste"
                : step.Action == "ensure-state" && step.ExpectedState?.StartsWith("sort:") == true
                ? $"Step {step.Number}: Sort {step.Target?.Name ?? "the list"} {step.ExpectedState[5..]}"
                : step.Action == "rename-sheet"
                ? $"Step {step.Number}: Rename sheet using {(step.RelativeWeekday is null ? step.Value : "next " + step.RelativeWeekday)}"
                : step.Action == "scroll"
                ? $"Step {step.Number}: Scroll {step.Key?.ToLowerInvariant()}"
                : step.Action == "resize-column"
                ? $"Step {step.Number}: Resize column {step.Target?.Name} by dragging its right edge"
                : step.Action == "select-first-row"
                ? $"Step {step.Number}: Select the first row in the sorted list"
                : step.Action == "select-all-items"
                ? $"Step {step.Number}: Select all items in the list"
                : step.Action == "context-selection"
                ? $"Step {step.Number}: Open the context menu for the selected items"
                : step.Action == "click" && step.Target?.ControlType == "ControlType.MenuItem"
                ? $"Step {step.Number}: Click {step.Target.Name} in the menu"
                : step.Action == "click"
                ? $"Step {step.Number}: Click {step.Target?.Name ?? "the recorded control"}"
                : $"Step {step.Number}: {step.Action} - {step.Target?.Name ?? "unknown control"}";
            var detail = step.Action == "type" ? step.RelativeWeekday is null
                    ? $"Enter: {step.Value}" : $"Enter text using the date of next {step.RelativeWeekday}." :
                step.Action == "manual-click" ? "The recorder could not verify the clicked control. Use the screenshot to identify the intended target." :
                step.Action == "manual-menu" ? $"Right-click {(step.ExpectedState?.StartsWith("selection-intent:", StringComparison.Ordinal) == true ? "the current selected rows" : step.Target?.Name ?? "the current selection")}, then click {step.Value}. The screenshot shows the menu before the command was chosen." :
                step.Action == "manual-menu-path" ? $"Follow the menu path {step.Value}. The screenshot shows the first menu; the submenu was not captured." :
                step.Action == "manual-paste" ? $"Select {step.Target?.Name ?? "the destination"}, then paste the clipboard contents." :
                step.Action == "ensure-state" && step.ExpectedState?.StartsWith("sort:") == true
                ? $"Make the list {step.ExpectedState[5..]} by {step.Target?.Name ?? "the selected column"}; verify the order before continuing." :
                step.Action == "rename-sheet" ? step.RelativeWeekday is null
                    ? $"Select {step.Target?.Name}, then set its final name to {step.Value}."
                    : $"Select {step.Target?.Name}, then replace the recorded date in its name with the date of next {step.RelativeWeekday}." :
                step.Action == "scroll" ? $"Move the {step.Key?.ToLowerInvariant()} view to the position shown in the screenshot." :
                step.Action == "select-first-row" ? "Select the first row, regardless of the data currently shown in it." :
                step.Action == "select-all-items" ? "Select every item in the list, regardless of the data currently shown in each row." :
                step.Action == "context-selection" ? "Right-click the current selection, regardless of the text in any item." :
                step.Explanation ?? step.ExpectedState ?? step.Key ?? "Follow the named action; the screenshot may show the result after the click.";
            if (step.Action == "click" && step.Target?.ControlType == "ControlType.MenuItem" && image is null)
                detail += "\nNo screenshot of this open submenu was captured. Use the named menu command.";
            else if (image is null)
                detail += "\nNo screenshot was captured for this step; follow the named action.";
            if (i == 0)
                detail += "\nRecording shortcut: Ctrl+Alt+Shift+I adds an intent note; this shortcut is not replayed.";
            if (!string.IsNullOrWhiteSpace(step.Intent))
                detail += "\nRecorded intent: " + Regex.Replace(step.Intent,
                    @"\b(password|pin|passcode)\s*(?:is|=|:)?\s*\S+", "$1 [redacted]", RegexOptions.IgnoreCase);
            var titleLines = Wrap(text, 62).Take(3).ToArray();
            var detailLines = Wrap(detail, 88).Take(8).ToArray();
            var detailStart = 755 - titleLines.Length * 21 - 4;
            var imageTop = detailStart - detailLines.Length * 15 - 20;
            var scale = Math.Min(540.0 / width, (imageTop - 40.0) / height);
            var w = width * scale; var h = height * scale;
            var stream = "";
            for (var line = 0; line < titleLines.Length; line++)
                stream += $"BT /F1 17 Tf 36 {755 - line * 21} Td ({Escape(titleLines[line])}) Tj ET\n";
            for (var line = 0; line < detailLines.Length; line++)
                stream += $"BT /F1 11 Tf 36 {detailStart - line * 15} Td ({Escape(detailLines[line])}) Tj ET\n";
            if (image is not null)
                stream += $"BT /F1 9 Tf 36 {imageTop + 5} Td ({Escape(step.Action is "manual-menu" or "manual-menu-path" ? "Menu before selection" : "Recorded visual reference")}) Tj ET\n";
            if (image is not null) stream += $"q {w:F2} 0 0 {h:F2} 36 {imageTop - h:F2} cm /Im Do Q\n";
            var data = Encoding.ASCII.GetBytes(stream);
            await Object(pageId, $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R >> {(image is null ? "" : $"/XObject << /Im {pageId + 2} 0 R >>")} >> /Contents {pageId + 1} 0 R >>");
            offsets.Add(output.Position);
            await Write($"{pageId + 1} 0 obj\n<< /Length {data.Length} >>\nstream\n");
            await output.WriteAsync(data, token);
            await Write("endstream\nendobj\n");
            offsets.Add(output.Position);
            if (image is null) await Write($"{pageId + 2} 0 obj\nnull\nendobj\n");
            else
            {
                await Write($"{pageId + 2} 0 obj\n<< /Type /XObject /Subtype /Image /Width {width} /Height {height} /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length {image.Length} >>\nstream\n");
                await output.WriteAsync(image, token);
                await Write("\nendstream\nendobj\n");
            }
        }
        var xref = output.Position;
        await Write($"xref\n0 {offsets.Count}\n0000000000 65535 f \n");
        foreach (var offset in offsets.Skip(1)) await Write($"{offset:0000000000} 00000 n \n");
        await Write($"trailer\n<< /Size {offsets.Count} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF");
    }
    private static List<PlanStep> ManualSteps(IReadOnlyList<PlanStep> steps)
    {
        var ordered = steps.ToList();
        // Menu providers can report a submenu command before its parent, and
        // completion timers can insert another application's steps before the
        // menu choice. Put each choice beside the menu that exposed it in the guide.
        for (var i = 0; i + 1 < ordered.Count; i++)
            if (ordered[i] is { Action: "click", Target: { ControlType: "ControlType.MenuItem" } child } &&
                ordered[i + 1] is { Action: "click", Target: { ControlType: "ControlType.MenuItem" } parent } &&
                child.Process == parent.Process && child.Window == parent.Window &&
                child.ParentName == parent.Name)
                (ordered[i], ordered[i + 1]) = (ordered[i + 1], ordered[i]);
        for (var i = 0; i < ordered.Count; i++)
        {
            var opener = ordered[i];
            if (opener.Action is not ("context-click" or "context-selection") || opener.Target is null) continue;
            for (var j = i + 1; j < ordered.Count; j++)
            {
                var candidate = ordered[j];
                if (candidate.Target is { } candidateTarget &&
                    candidateTarget.Process == opener.Target.Process &&
                    candidateTarget.Window == opener.Target.Window)
                {
                    if (candidate is { Action: "click", Target: { ControlType: "ControlType.MenuItem" } })
                    {
                        ordered.RemoveAt(j);
                        ordered.Insert(i + 1, candidate);
                    }
                    break;
                }
            }
        }
        var result = new List<PlanStep>();
        for (var i = 0; i < ordered.Count; i++)
        {
            var current = ordered[i];
            var next = i + 1 < ordered.Count ? ordered[i + 1] : null;
            var third = i + 2 < ordered.Count ? ordered[i + 2] : null;
            if (current is { Action: "click", Target: { ControlType: "ControlType.Button" } root } &&
                next is { Action: "click", Target: { ControlType: "ControlType.MenuItem" } parent } &&
                third is { Action: "click", Target: { ControlType: "ControlType.MenuItem" } child } &&
                root.Process == parent.Process && parent.Process == child.Process &&
                parent.Window == child.Window && child.ParentName == parent.Name)
            {
                result.Add(current with { Action = "manual-menu-path",
                    Value = $"{root.Name} > {parent.Name} > {child.Name}",
                    Screenshot = current.Screenshot });
                i += 2;
                continue;
            }
            if (current.Action is "context-click" or "context-selection" &&
                next?.Target?.ControlType == "ControlType.MenuItem" && next.Action == "click")
            {
                result.Add(current with { Action = "manual-menu", Value = next.Target.Name,
                    Screenshot = current.Screenshot });
                i++;
                continue;
            }
            if (current.Action == "click" && current.Target?.ControlType == "ControlType.DataItem" &&
                next?.Action == "key" && next.Key is "Control+V" or "Ctrl+V")
            {
                result.Add(current with { Action = "manual-paste", Screenshot = next.Screenshot ?? current.Screenshot });
                i++;
                continue;
            }
            if (current is { Action: "click", Target: { ControlType: "ControlType.MenuItem" } })
                current = current with { Screenshot = null };
            result.Add(current);
        }
        return result.Select((step, index) => step with { Number = index + 1 }).ToList();
    }
    private static (int X, int Y)? FindMenuPoint(PlanStep step, IReadOnlyList<RecordedEvent> events)
    {
        var name = step.Action == "manual-menu-path"
            ? step.Value?.Split('>').LastOrDefault()?.Trim() : step.Value;
        var click = events.LastOrDefault(item => item.Kind == "click" &&
            item.Target?.Process == step.Target?.Process && item.Target?.Name == name &&
            item.ClickX is not null && item.ClickY is not null);
        return click is null ? null : (click.ClickX!.Value, click.ClickY!.Value);
    }
    private static byte[] CropMenuImage(byte[] jpeg, int screenX, int screenY)
    {
        using var source = new Bitmap(new MemoryStream(jpeg));
        var screen = Screen.FromPoint(new System.Drawing.Point(screenX, screenY)).Bounds;
        if (screen.Width != source.Width || screen.Height != source.Height) return jpeg;
        var x = screenX - screen.Left;
        var y = screenY - screen.Top;
        if (x < 0 || x >= source.Width || y < 0 || y >= source.Height) return jpeg;
        var width = Math.Min(1000, source.Width);
        var height = Math.Min(650, source.Height);
        var left = Math.Clamp(x - 350, 0, source.Width - width);
        var top = Math.Clamp(y - 450, 0, source.Height - height);
        using var cropped = source.Clone(new Rectangle(left, top, width, height), PixelFormat.Format24bppRgb);
        using var stream = new MemoryStream();
        cropped.Save(stream, ImageFormat.Jpeg);
        return stream.ToArray();
    }
    private static IEnumerable<string> Wrap(string value, int width)
    {
        foreach (var paragraph in value.Replace("\r", "").Split('\n'))
        {
            var line = new StringBuilder();
            foreach (var word in paragraph.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                if (line.Length > 0 && line.Length + word.Length + 1 > width)
                {
                    yield return line.ToString();
                    line.Clear();
                }
                if (line.Length > 0) line.Append(' ');
                line.Append(word);
            }
            if (line.Length > 0) yield return line.ToString();
        }
    }
    private static string Escape(string text) => new string(text.Where(c => c >= 32 && c < 127).ToArray())
        .Replace("\\", "\\\\").Replace("(", "\\(").Replace(")", "\\)");
}
