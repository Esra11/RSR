using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace DesktopSteps;

// Provides verified workbook operations through Excel's native object model when UIA
// cannot express dynamic ranges, formula extension or chart state reliably. Callers
// still supply recorded workbook, worksheet and target evidence before using it.
internal static class ExcelNativeSheet
{
    internal sealed record LegendLayout(double Left, double Top, double Width, double Height);
    internal sealed record ChartSource(string Sheet, string Range);
    private static int ColumnNumber(string column) =>
        column.Aggregate(0, (number, character) => checked(number * 26 + character - 'A' + 1));

    internal static void ValidateChartSource(ChartSource source)
    {
        var match = Regex.Match(source.Range ?? "",
            @"^(?<first>[A-Z]{1,3})(?<row>[1-9]\d{0,6}):(?<last>[A-Z]{1,3})(?<end>[1-9]\d{0,6})$");
        if (string.IsNullOrWhiteSpace(source.Sheet) || !match.Success ||
            ColumnNumber(match.Groups["last"].Value) > 16384 ||
            ColumnNumber(match.Groups["last"].Value) != ColumnNumber(match.Groups["first"].Value) + 1 ||
            int.Parse(match.Groups["end"].Value) > 1048576 ||
            int.Parse(match.Groups["row"].Value) > int.Parse(match.Groups["end"].Value))
            throw new InvalidDataException("Captured chart source must be an ordered, adjacent two-column worksheet range.");
    }

    internal static ChartSource ReadChartSource(dynamic sheet, string chartName)
    {
        object charts = sheet.ChartObjects();
        object chartObject = ((dynamic)charts).Item(chartName);
        object chart = ((dynamic)chartObject).Chart;
        object series = ((dynamic)chart).SeriesCollection();
        object? item = null;
        try
        {
            if ((int)((dynamic)series).Count != 1)
                throw new NotSupportedException("Chart source capture currently requires one category/value series.");
            item = ((dynamic)series).Item(1);
            var formula = (string)((dynamic)item).Formula;
            var match = Regex.Match(formula,
                @"^=SERIES\(.*?,(?<sheet>'(?:[^']|'')+'|[^!,]+)!\$(?<first>[A-Z]+)\$(?<row>[1-9]\d*):\$\k<first>\$(?<end>[1-9]\d*),\k<sheet>!\$(?<last>[A-Z]+)\$\k<row>:\$\k<last>\$\k<end>,\d+\)$");
            var sheetReference = match.Groups["sheet"].Value;
            var sheetName = sheetReference.StartsWith('\'')
                ? sheetReference[1..^1].Replace("''", "'") : sheetReference;
            if (!match.Success || sheetName != (string)sheet.Name || sheetName.Contains('[') ||
                ColumnNumber(match.Groups["last"].Value) != ColumnNumber(match.Groups["first"].Value) + 1)
                throw new NotSupportedException("Chart source capture requires adjacent category/value ranges on the active worksheet.");
            return new(sheetName, match.Groups["first"].Value + match.Groups["row"].Value + ":" +
                match.Groups["last"].Value + match.Groups["end"].Value);
        }
        finally
        {
            foreach (var value in new[] { item, series, chart, chartObject, charts })
                if (value is not null && Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
    }

    internal static ChartSource ReadChartSourceFromSheet(dynamic activeSheet, string sheetName, string chartName)
    {
        object? workbook = null, worksheets = null, sheet = null;
        try
        {
            workbook = activeSheet.Parent;
            worksheets = ((dynamic)workbook).Worksheets;
            sheet = ((dynamic)worksheets).Item(sheetName);
            if (!string.Equals((string)((dynamic)sheet).Name, sheetName, StringComparison.Ordinal))
                throw new InvalidOperationException("The recorded chart worksheet could not be resolved uniquely.");
            return ReadChartSource((dynamic)sheet, chartName);
        }
        finally
        {
            foreach (var item in new[] { sheet, worksheets, workbook })
                if (item is not null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
        }
    }

    internal static void ApplyChartSource(dynamic sheet, string chartName, ChartSource source, TablePaste? paste)
    {
        ValidateChartSource(source);
        if ((string)sheet.Name != source.Sheet)
            throw new InvalidOperationException("The captured chart worksheet is not active.");
        var start = source.Range.Split(':')[0];
        if (paste is not null && paste.Sheet == source.Sheet && paste.Destination == start &&
            paste.Values.GetLength(1) == 2)
        {
            SetChartSource(sheet, chartName, paste);
            return;
        }
        object range = sheet.Range[source.Range];
        try
        {
            object values = ((dynamic)range).Value2;
            if (values is not Array array || array.Rank != 2 || array.GetLength(1) != 2)
                throw new InvalidOperationException("Captured chart source is not a populated two-column rectangle.");
            var snapshot = new object?[array.GetLength(0), 2];
            for (var row = 0; row < snapshot.GetLength(0); row++)
                for (var column = 0; column < 2; column++)
                    snapshot[row, column] = array.GetValue(row + 1, column + 1);
            SetChartSource(sheet, chartName, new TablePaste(source.Sheet, start, snapshot));
        }
        finally { Marshal.ReleaseComObject(range); }
    }

    internal sealed record ChartCapture(string Name, string Sheet, LegendLayout? Legend);

    internal static string? ChartNameAtPoint(dynamic sheet, int x, int y) =>
        CaptureChartAtPoint(sheet, x, y, false)?.Name;

    internal static ChartCapture? CaptureChartAtPoint(dynamic sheet, int x, int y, bool captureLegend,
        IReadOnlyList<string>? ancestorNames = null)
    {
        object application = sheet.Application;
        object window = ((dynamic)application).ActiveWindow;
        object? hit = null, charts = null, chart = null;
        try
        {
            hit = ((dynamic)window).RangeFromPoint(x, y);
            if (hit is null && ancestorNames is null) return null;
            object? hitName = hit is null ? null : ((dynamic)hit).Name;
            var name = hitName as string;
            if (hitName is not null && Marshal.IsComObject(hitName)) Marshal.ReleaseComObject(hitName);
            charts = sheet.ChartObjects();
            var nativeNames = new List<string>();
            var candidates = (name is not null ? new[] { name } : Array.Empty<string>())
                .Concat(ancestorNames ?? Array.Empty<string>()).Distinct();
            foreach (var candidate in candidates)
            {
                try
                {
                    chart = ((dynamic)charts).Item(candidate);
                    if ((string)((dynamic)chart).Name == candidate) nativeNames.Add(candidate);
                }
                catch (COMException ex)
                {
                    System.Diagnostics.Trace.WriteLine("Chart ancestor is not a native chart: " + ex.Message);
                }
                finally
                {
                    if (chart is not null) Marshal.ReleaseComObject(chart);
                    chart = null;
                }
            }
            var matchedNames = name is not null && nativeNames.Contains(name)
                ? new List<string> { name }
                : nativeNames.Where(candidate => ancestorNames?.Contains(candidate) == true).ToList();
            if (matchedNames.Count > 1)
                throw new InvalidOperationException("Chart ancestors matched multiple native charts.");
            if (matchedNames.Count == 0) return null;
            ChartCapture? captured = null;
            foreach (var chartName in matchedNames)
            {
                chart = ((dynamic)charts).Item(chartName);
                {
                    if (captured is not null)
                        throw new InvalidOperationException("Chart ancestors matched multiple native charts.");
                    LegendLayout? layout = null;
                    if (captureLegend)
                    {
                        object contents = ((dynamic)chart).Chart;
                        try
                        {
                            if ((bool)((dynamic)contents).HasLegend)
                                layout = ReadLegendLayoutFromChartObject(chart, contents);
                        }
                        finally { Marshal.ReleaseComObject(contents); }
                    }
                    captured = new ChartCapture(chartName, (string)sheet.Name, layout);
                }
                Marshal.ReleaseComObject(chart);
                chart = null;
            }
            return captured;
        }
        finally
        {
            foreach (var item in new[] { chart, charts, hit, window, application })
                if (item is not null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
        }
    }

    internal static LegendLayout ReadLegendLayout(dynamic sheet, string chartName)
    {
        object charts = sheet.ChartObjects();
        object chartObject = ((dynamic)charts).Item(chartName);
        object chart = ((dynamic)chartObject).Chart;
        try { return ReadLegendLayoutFromChartObject(chartObject, chart); }
        finally
        {
            Marshal.ReleaseComObject(chart);
            Marshal.ReleaseComObject(chartObject); Marshal.ReleaseComObject(charts);
        }
    }

    private static LegendLayout ReadLegendLayoutFromChartObject(object chartObject, object chart)
    {
        object legend = ((dynamic)chart).Legend;
        try
        {
            var width = (double)((dynamic)chartObject).Width;
            var height = (double)((dynamic)chartObject).Height;
            return new((double)((dynamic)legend).Left / width, (double)((dynamic)legend).Top / height,
                (double)((dynamic)legend).Width / width, (double)((dynamic)legend).Height / height);
        }
        finally
        {
            Marshal.ReleaseComObject(legend);
        }
    }

    internal static void SetLegendLayout(dynamic sheet, string chartName, LegendLayout layout)
    {
        if (new[] { layout.Left, layout.Top, layout.Width, layout.Height }.Any(value => !double.IsFinite(value)) ||
            layout.Left < 0 || layout.Top < 0 || layout.Width <= 0 || layout.Height <= 0 ||
            layout.Left + layout.Width > 1.02 || layout.Top + layout.Height > 1.02)
            throw new InvalidDataException("Recorded legend layout is outside its chart.");
        object charts = sheet.ChartObjects();
        object chartObject = ((dynamic)charts).Item(chartName);
        object chart = ((dynamic)chartObject).Chart;
        object legend = ((dynamic)chart).Legend;
        try
        {
            var width = (double)((dynamic)chartObject).Width;
            var height = (double)((dynamic)chartObject).Height;
            ((dynamic)legend).Width = layout.Width * width;
            ((dynamic)legend).Height = layout.Height * height;
            ((dynamic)legend).Left = layout.Left * width;
            ((dynamic)legend).Top = layout.Top * height;
        }
        finally
        {
            Marshal.ReleaseComObject(legend); Marshal.ReleaseComObject(chart);
            Marshal.ReleaseComObject(chartObject); Marshal.ReleaseComObject(charts);
        }
        var actual = ReadLegendLayout(sheet, chartName);
        if (Math.Abs(actual.Left - layout.Left) > .005 || Math.Abs(actual.Top - layout.Top) > .005 ||
            Math.Abs(actual.Width - layout.Width) > .005 || Math.Abs(actual.Height - layout.Height) > .005)
            throw new InvalidOperationException("Legend resize did not match its recorded chart-relative layout.");
    }

    internal static void SetChartSource(dynamic sheet, string chartName, TablePaste paste)
    {
        if ((string)sheet.Name != paste.Sheet || paste.Values.GetLength(1) != 2)
            throw new InvalidOperationException("Chart source requires the verified pasted two-column range on the active destination sheet.");
        object charts = sheet.ChartObjects();
        object? chartObject = null, chart = null, first = null, range = null, seriesCollection = null, series = null;
        try
        {
            chartObject = ((dynamic)charts).Item(chartName);
            if ((string)((dynamic)chartObject).Name != chartName)
                throw new InvalidOperationException("The recorded chart identity could not be verified.");
            chart = ((dynamic)chartObject).Chart;
            first = sheet.Range[paste.Destination];
            range = ((dynamic)first).Resize[paste.Values.GetLength(0), 2];
            ((dynamic)chart).SetSourceData(range, 2); // xlColumns: categories followed by values.
            seriesCollection = ((dynamic)chart).SeriesCollection();
            if ((int)((dynamic)seriesCollection).Count != 1)
                throw new InvalidOperationException("Chart source did not produce one category/value series.");
            series = ((dynamic)seriesCollection).Item(1);
            // Explicit assignment avoids Excel treating the first data row as a header.
            object categories = ((dynamic)range).Columns[1];
            object values = ((dynamic)range).Columns[2];
            try
            {
                ((dynamic)series).XValues = categories;
                ((dynamic)series).Values = values;
            }
            finally { Marshal.ReleaseComObject(categories); Marshal.ReleaseComObject(values); }
            object actualCategories = ((dynamic)series).XValues;
            object actualValues = ((dynamic)series).Values;
            if (actualCategories is not Array x || actualValues is not Array y ||
                x.Length != paste.Values.GetLength(0) || y.Length != paste.Values.GetLength(0))
                throw new InvalidOperationException("Chart data length does not match the pasted range.");
            for (var row = 0; row < x.Length; row++)
                if (!Equals(x.GetValue(row + x.GetLowerBound(0)), paste.Values[row, 0]) ||
                    !Equals(y.GetValue(row + y.GetLowerBound(0)), paste.Values[row, 1]))
                    throw new InvalidOperationException($"Chart source mismatch at pasted row {row + 1}.");
        }
        finally
        {
            foreach (var item in new[] { series, seriesCollection, range, first, chart, chartObject, charts })
                if (item is not null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
        }
    }

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(nint parent, EnumWindow callback, nint parameter);
    private delegate bool EnumWindow(nint window, nint parameter);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(nint window, System.Text.StringBuilder name, int capacity);
    [DllImport("user32.dll")]
    private static extern nint GetAncestor(nint window, uint flags);
    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(nint window, uint objectId,
        ref Guid iid, [MarshalAs(UnmanagedType.Interface)] out object nativeObject);

    internal static T Read<T>(nint workbookWindow, Func<dynamic, T> read)
    {
        nint document = 0;
        EnumChildWindows(workbookWindow, (window, _) =>
        {
            var name = new System.Text.StringBuilder(256);
            GetClassName(window, name, name.Capacity);
            if (name.ToString() != "EXCEL7") return true;
            document = window;
            return false;
        }, 0);
        if (document == 0) throw new InvalidOperationException("The verified workbook exposes no native worksheet object.");
        var iid = new Guid("00020400-0000-0000-C000-000000000046");
        Marshal.ThrowExceptionForHR(AccessibleObjectFromWindow(document, 0xFFFFFFF0, ref iid, out var nativeObject));
        object? sheet = null;
        try
        {
            dynamic window = nativeObject;
            if (GetAncestor((nint)(int)window.Hwnd, 2) != GetAncestor(workbookWindow, 2))
                throw new InvalidOperationException("Native Excel object belongs to a different workbook window.");
            sheet = window.ActiveSheet;
            return read(sheet);
        }
        finally
        {
            if (sheet is not null && Marshal.IsComObject(sheet)) Marshal.ReleaseComObject(sheet);
            if (Marshal.IsComObject(nativeObject)) Marshal.ReleaseComObject(nativeObject);
        }
    }

    internal static List<int> VisibleDataRows(dynamic sheet, string column, int firstRow,
        CancellationToken token)
    {
        if (!Regex.IsMatch(column, @"\A[A-Z]{1,3}\z") || firstRow < 2 || firstRow > 1048576)
            throw new InvalidDataException("Invalid visible-fill data address.");
        var lastRow = Math.Min(1048576, firstRow + 10000);
        object range = sheet.Range[$"{column}{firstRow}:{column}{lastRow}"];
        object allRows = sheet.Rows;
        try
        {
            object? values = ((dynamic)range).Value2;
            var rows = new List<int>();
            for (var row = firstRow; row <= lastRow; row++)
            {
                token.ThrowIfCancellationRequested();
                object rowRange = ((dynamic)allRows)[row];
                bool hidden;
                try { hidden = (bool)((dynamic)rowRange).Hidden; }
                finally { Marshal.ReleaseComObject(rowRange); }
                if (hidden) continue;
                var value = values is Array array ? array.GetValue(row - firstRow + 1, 1) : values;
                if (value is null || string.IsNullOrWhiteSpace(Convert.ToString(value,
                    System.Globalization.CultureInfo.InvariantCulture))) return rows;
                rows.Add(row);
            }
            throw new InvalidOperationException("Visible fill found no empty adjacent cell within 10000 rows; no fill was performed.");
        }
        finally
        {
            Marshal.ReleaseComObject(allRows);
            Marshal.ReleaseComObject(range);
        }
    }

    internal static void VerifyValues(dynamic sheet, string column, IReadOnlyList<int> rows, string value)
    {
        object range = sheet.Range[$"{column}{rows[0]}:{column}{rows[^1]}"];
        try
        {
            object values = ((dynamic)range).Value2;
            foreach (var row in rows)
            {
                var actual = values is Array array ? array.GetValue(row - rows[0] + 1, 1) : values;
                if (Convert.ToString(actual, System.Globalization.CultureInfo.InvariantCulture) != value)
                    throw new InvalidOperationException($"Filled value was not verified in {column}{row}.");
            }
        }
        finally { Marshal.ReleaseComObject(range); }
    }

    internal static void VerifyEmptyColumn(dynamic sheet, string column)
    {
        object range = sheet.Range[column + ":" + column];
        object application = sheet.Application;
        object functions = ((dynamic)application).WorksheetFunction;
        try
        {
            if ((double)((dynamic)functions).CountA(range) != 0)
                throw new InvalidOperationException($"Column {column} still contains data after Delete.");
        }
        finally
        {
            Marshal.ReleaseComObject(functions);
            Marshal.ReleaseComObject(application);
            Marshal.ReleaseComObject(range);
        }
    }

    internal static void VerifyColumnSelection(dynamic sheet, string column)
    {
        object application = sheet.Application;
        object selection = ((dynamic)application).Selection;
        object rows = ((dynamic)selection).Rows;
        object columns = ((dynamic)selection).Columns;
        object expected = sheet.Range[column + "1"];
        try
        {
            if ((int)((dynamic)rows).Count != 1048576 || (int)((dynamic)columns).Count != 1 ||
                (int)((dynamic)selection).Column != (int)((dynamic)expected).Column)
                throw new InvalidOperationException($"Excel did not select the entire {column} column; Delete was not sent.");
        }
        finally
        {
            Marshal.ReleaseComObject(expected);
            Marshal.ReleaseComObject(columns);
            Marshal.ReleaseComObject(rows);
            Marshal.ReleaseComObject(selection);
            Marshal.ReleaseComObject(application);
        }
    }

    internal static void SetValues(dynamic sheet, string column, IReadOnlyList<int> rows, string value)
    {
        if (!Regex.IsMatch(column, @"\A[A-Z]{1,3}\z") || rows.Count == 0 ||
            rows.Any(row => row < 2 || row > 1048576) || rows.Distinct().Count() != rows.Count)
            throw new InvalidDataException("Invalid visible-fill destination rows.");
        foreach (var row in rows)
        {
            object cell = sheet.Range[$"{column}{row}"];
            try { ((dynamic)cell).Value2 = value; }
            finally { Marshal.ReleaseComObject(cell); }
        }
        VerifyValues(sheet, column, rows, value);
    }

    internal static string ScrollViewport(dynamic sheet, string axis, int direction, int amount)
    {
        if (axis is not ("Vertical" or "Horizontal") || direction is < -1 or > 1 || amount < 1)
            throw new InvalidDataException("Invalid native Excel viewport movement.");
        object application = sheet.Application;
        object window = ((dynamic)application).ActiveWindow;
        try
        {
            var property = axis == "Vertical" ? "ScrollRow" : "ScrollColumn";
            var maximum = axis == "Vertical" ? 1048576 : 16384;
            var before = Convert.ToInt32(window.GetType().InvokeMember(property,
                System.Reflection.BindingFlags.GetProperty, null, window, null),
                System.Globalization.CultureInfo.InvariantCulture);
            var after = direction == 0 ? before : Math.Clamp(before + direction * amount, 1, maximum);
            window.GetType().InvokeMember(property, System.Reflection.BindingFlags.SetProperty,
                null, window, [after]);
            var verified = Convert.ToInt32(window.GetType().InvokeMember(property,
                System.Reflection.BindingFlags.GetProperty, null, window, null),
                System.Globalization.CultureInfo.InvariantCulture);
            if (after != before && (direction > 0 && verified <= before || direction < 0 && verified >= before))
                throw new InvalidOperationException($"Excel {axis.ToLowerInvariant()} viewport did not move in the requested direction.");
            return $"{property} {before} to {verified}";
        }
        finally
        {
            Marshal.ReleaseComObject(window);
            Marshal.ReleaseComObject(application);
        }
    }

    internal static void ExpandLegendOppositePlot(dynamic sheet, string chartName)
    {
        object charts = sheet.ChartObjects();
        object chartObject = ((dynamic)charts).Item(chartName);
        object chart = ((dynamic)chartObject).Chart;
        object legend = ((dynamic)chart).Legend;
        try
        {
            var width = (double)((dynamic)chartObject).Width;
            var height = (double)((dynamic)chartObject).Height;
            // Use chart-relative geometry so zoom, resolution and worksheet
            // position cannot change the result. The plot occupies the left;
            // give the legend the remaining right side with readable columns.
            ((dynamic)legend).IncludeInLayout = false;
            ((dynamic)legend).Width = width * .44;
            ((dynamic)legend).Height = height * .74;
            ((dynamic)legend).Left = width * .51;
            ((dynamic)legend).Top = height * .10;
        }
        finally
        {
            Marshal.ReleaseComObject(legend); Marshal.ReleaseComObject(chart);
            Marshal.ReleaseComObject(chartObject); Marshal.ReleaseComObject(charts);
        }
        var actual = ReadLegendLayout(sheet, chartName);
        if (actual.Left < .48 || actual.Width < .40 || actual.Height < .65 || actual.Left + actual.Width > 1.02)
            throw new InvalidOperationException("The chart legend did not expand across the side opposite the plot.");
    }

    internal static void SelectEntireColumn(dynamic sheet, string column)
    {
        if (!Regex.IsMatch(column, @"\A[A-Z]{1,3}\z", RegexOptions.IgnoreCase) ||
            ColumnNumber(column.ToUpperInvariant()) > 16384)
            throw new InvalidDataException("Invalid Excel column selection.");
        object range = sheet.Range[column.ToUpperInvariant() + ":" + column.ToUpperInvariant()];
        try
        {
            ((dynamic)range).Select();
            VerifyColumnSelection(sheet, column.ToUpperInvariant());
        }
        finally { Marshal.ReleaseComObject(range); }
    }

    internal static void ClearColumnContents(dynamic sheet, string column)
    {
        if (!Regex.IsMatch(column, @"\A[A-Z]{1,3}\z", RegexOptions.IgnoreCase) ||
            ColumnNumber(column.ToUpperInvariant()) > 16384)
            throw new InvalidDataException("Invalid Excel column clear operation.");
        object range = sheet.Range[column.ToUpperInvariant() + ":" + column.ToUpperInvariant()];
        try { ((dynamic)range).ClearContents(); }
        finally { Marshal.ReleaseComObject(range); }
        VerifyEmptyColumn(sheet, column.ToUpperInvariant());
    }

    internal static void ClearCellContents(dynamic sheet, string address)
    {
        if (!Regex.IsMatch(address, @"\A[A-Z]{1,3}[1-9]\d*\z", RegexOptions.IgnoreCase))
            throw new InvalidDataException("Invalid Excel cell clear operation.");
        object cell = sheet.Range[address.ToUpperInvariant()];
        try
        {
            ((dynamic)cell).ClearContents();
            var actual = Convert.ToString((object?)((dynamic)cell).Value2,
                System.Globalization.CultureInfo.InvariantCulture) ?? "";
            if (actual.Length > 0)
                throw new InvalidOperationException($"Cell {address} still contains data after Delete.");
        }
        finally { Marshal.ReleaseComObject(cell); }
    }

    internal static int ContiguousDataEnd(dynamic sheet, string column, int firstRow,
            CancellationToken token, string[]? stopValues = null)
    {
        if (!Regex.IsMatch(column, @"\A[A-Z]{1,3}\z") || firstRow < 1 || firstRow > 1048576)
            throw new InvalidDataException("Invalid column-copy start address.");
        var lastRow = Math.Min(1048576, firstRow + 10000);
        object range = sheet.Range[$"{column}{firstRow}:{column}{lastRow}"];
        try
        {
            object? values = ((dynamic)range).Value2;
            for (var row = firstRow; row <= lastRow; row++)
            {
                token.ThrowIfCancellationRequested();
                var value = values is Array array ? array.GetValue(row - firstRow + 1, 1) : values;
                var text = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
                if (value is null || text == "" || stopValues?.Contains(text, StringComparer.Ordinal) == true)
                    return row - 1;
                if (stopValues is { Length: > 0 } && value is not string)
                {
                    object cell = sheet.Range[$"{column}{row}"];
                    try
                    {
                        if (stopValues.Contains((string)((dynamic)cell).Text, StringComparer.Ordinal))
                            return row - 1;
                    }
                    finally { Marshal.ReleaseComObject(cell); }
                }
            }
            throw new InvalidOperationException("Column copy found no empty cell or configured stop value within 10000 rows; nothing was copied.");
        }
        finally { Marshal.ReleaseComObject(range); }
    }

    internal static void VerifyRangeSelection(dynamic sheet, string address)
    {
        object application = sheet.Application;
        object selection = ((dynamic)application).Selection;
        object expected = sheet.Range[address];
        try
        {
            if ((string)((dynamic)selection).Address != (string)((dynamic)expected).Address ||
                (string)((dynamic)selection).Parent.Name != (string)sheet.Name)
                throw new InvalidOperationException($"Excel did not select {address}; copy was not sent.");
        }
        finally
        {
            Marshal.ReleaseComObject(expected);
            Marshal.ReleaseComObject(selection);
            Marshal.ReleaseComObject(application);
        }
    }

    internal static void SelectRange(dynamic sheet, string address)
    {
        object range = sheet.Range[address];
        try { ((dynamic)range).Select(); }
        finally { Marshal.ReleaseComObject(range); }
        VerifyRangeSelection(sheet, address);
    }

    internal static void VerifyCopyMode(dynamic sheet)
    {
        object application = sheet.Application;
        try
        {
            if (Convert.ToInt32(((dynamic)application).CutCopyMode) != 1)
                throw new InvalidOperationException("Excel has no active copied cell range; replay stopped before pasting.");
        }
        finally { Marshal.ReleaseComObject(application); }
    }
    internal static object?[,] CaptureSelectionValues(dynamic sheet)
    {
        object application = sheet.Application;
        object selection = ((dynamic)application).Selection;
        object rows = ((dynamic)selection).Rows;
        object columns = ((dynamic)selection).Columns;
        object areas = ((dynamic)selection).Areas;
        try
        {
            var rowCount = (int)((dynamic)rows).Count;
            var columnCount = (int)((dynamic)columns).Count;
            if ((int)((dynamic)areas).Count != 1 || (long)rowCount * columnCount > 100000)
                throw new InvalidOperationException("Copied selection must be one bounded rectangle for paste verification.");
            object? values = ((dynamic)selection).Value2;
            var snapshot = new object?[rowCount, columnCount];
            for (var row = 0; row < rowCount; row++)
                for (var column = 0; column < columnCount; column++)
                    snapshot[row, column] = values is Array array
                        ? array.GetValue(row + 1, column + 1) : values;
            return snapshot;
        }
        finally
        {
            Marshal.ReleaseComObject(areas);
            Marshal.ReleaseComObject(columns);
            Marshal.ReleaseComObject(rows);
            Marshal.ReleaseComObject(selection);
            Marshal.ReleaseComObject(application);
        }
    }

    internal static string SelectionAddress(dynamic sheet)
    {
        object application = sheet.Application;
        object selection = ((dynamic)application).Selection;
        object areas = ((dynamic)selection).Areas;
        try
        {
            if ((int)((dynamic)areas).Count != 1)
                throw new InvalidOperationException("Excel copy selection has multiple areas; record a single rectangular range.");
            var address = ((string)((dynamic)selection).Address).Replace("$", "");
            if (!Regex.IsMatch(address, @"^[A-Z]{1,3}[1-9]\d*(?::[A-Z]{1,3}[1-9]\d*)?$"))
                throw new InvalidOperationException("Excel copy selection is not a bounded cell range.");
            return address;
        }
        finally
        {
            Marshal.ReleaseComObject(areas);
            Marshal.ReleaseComObject(selection);
            Marshal.ReleaseComObject(application);
        }
    }
    internal static void PasteValues(dynamic sheet, string destination, object?[,] expected)
    {
        SelectRange(sheet, destination);
        object application = sheet.Application;
        object firstCell = sheet.Range[destination];
        object range = ((dynamic)firstCell).Resize[expected.GetLength(0), expected.GetLength(1)];
        try
        {
            ((dynamic)range).Value2 = expected;
            object? actual = ((dynamic)range).Value2;
            if (!Equals(((dynamic)range).HasFormula, false))
                throw new InvalidOperationException("Paste Values left formulas in the destination; replay stopped.");
            for (var row = 0; row < expected.GetLength(0); row++)
                for (var column = 0; column < expected.GetLength(1); column++)
                {
                    var value = actual is Array array ? array.GetValue(row + 1, column + 1) : actual;
                    if (!Equals(value, expected[row, column]))
                        throw new InvalidOperationException($"Paste Values was not verified at destination offset {row + 1}, {column + 1}.");
                }
            ((dynamic)application).CutCopyMode = 0;
        }
        finally
        {
            Marshal.ReleaseComObject(range);
            Marshal.ReleaseComObject(firstCell);
            Marshal.ReleaseComObject(application);
        }
    }

    internal static void CopyRange(dynamic sheet, string sourceAddress, string destinationAddress)
    {
        object source = sheet.Range[sourceAddress];
        object destination = sheet.Range[destinationAddress];
        try
        {
            if ((int)((dynamic)source).Rows.Count != (int)((dynamic)destination).Rows.Count ||
                (int)((dynamic)source).Columns.Count != (int)((dynamic)destination).Columns.Count)
                throw new InvalidOperationException("Recorded source and destination ranges have different dimensions.");
            ((dynamic)source).Copy(destination);
            // R1C1 preserves the relative meaning of formulas after Excel adjusts
            // their row references for the destination.
            object sourceFormula = ((dynamic)source).FormulaR1C1;
            object destinationFormula = ((dynamic)destination).FormulaR1C1;
            if (!RangeContentsEqual(sourceFormula, destinationFormula))
                throw new InvalidOperationException("Excel did not reproduce the recorded row in the destination range.");
            SelectRange(sheet, destinationAddress);
        }
        finally
        {
            Marshal.ReleaseComObject(destination);
            Marshal.ReleaseComObject(source);
        }

        static bool RangeContentsEqual(object? left, object? right)
        {
            if (left is not Array leftArray || right is not Array rightArray) return Equals(left, right);
            if (leftArray.Length != rightArray.Length) return false;
            var leftValues = leftArray.Cast<object?>().ToArray();
            var rightValues = rightArray.Cast<object?>().ToArray();
            return leftValues.SequenceEqual(rightValues);
        }
    }

    internal sealed record DuplicatedRow(string SourceRange, string DestinationRange,
        string SourceDateCell, string DestinationDateCell);

    internal static DuplicatedRow CopyLastPopulatedRow(dynamic sheet, string recordedSourceRange)
    {
        var match = Regex.Match(recordedSourceRange,
            @"^(?<first>[A-Z]{1,3})(?<row>[1-9]\d*):(?<last>[A-Z]{1,3})\k<row>$",
            RegexOptions.IgnoreCase);
        if (!match.Success)
            throw new InvalidDataException("The recorded previous-row range is not a single bounded Excel row.");
        var firstColumn = match.Groups["first"].Value.ToUpperInvariant();
        var lastColumn = match.Groups["last"].Value.ToUpperInvariant();
        var recordedRow = int.Parse(match.Groups["row"].Value,
            System.Globalization.CultureInfo.InvariantCulture);
        var firstColumnNumber = ColumnNumber(firstColumn);
        if (ColumnNumber(lastColumn) < firstColumnNumber)
            throw new InvalidDataException("The recorded previous-row range has reversed columns.");

        object rows = sheet.Rows;
        object bottom = sheet.Cells[((dynamic)rows).Count, firstColumnNumber];
        object last = ((dynamic)bottom).End[-4162]; // xlUp
        try
        {
            var sourceRow = (int)((dynamic)last).Row;
            if (sourceRow < recordedRow || sourceRow >= 1048576)
                throw new InvalidOperationException(
                    $"The live last populated row {sourceRow} is outside the safe recorded boundary starting at row {recordedRow}.");
            var destinationRow = sourceRow + 1;
            var sourceRange = $"{firstColumn}{sourceRow}:{lastColumn}{sourceRow}";
            var destinationRange = $"{firstColumn}{destinationRow}:{lastColumn}{destinationRow}";
            object destination = sheet.Range[destinationRange];
            try
            {
                // Never overwrite an earlier history row. The semantic intent is to append
                // after the current last date, so every destination cell in the copied span
                // must still be empty when replay resolves the row.
                object? values = ((dynamic)destination).Value2;
                var occupied = values is Array array
                    ? array.Cast<object?>().Any(value => value is not null &&
                        !string.IsNullOrWhiteSpace(Convert.ToString(value)))
                    : values is not null && !string.IsNullOrWhiteSpace(Convert.ToString(values));
                if (occupied)
                    throw new InvalidOperationException(
                        $"The row after the live last populated date is not empty ({destinationRange}); replay stopped without overwriting it.");
            }
            finally { Marshal.ReleaseComObject(destination); }
            CopyRange(sheet, sourceRange, destinationRange);
            return new(sourceRange, destinationRange, $"{firstColumn}{sourceRow}",
                $"{firstColumn}{destinationRow}");
        }
        finally
        {
            Marshal.ReleaseComObject(last); Marshal.ReleaseComObject(bottom);
            Marshal.ReleaseComObject(rows);
        }
    }

    internal static string UpdateCopiedRowToRelativeWeekday(dynamic sheet, string sourceDateAddress,
        string address, string destinationRange, string weekday, DateTime runDate)
    {
        if (!Enum.TryParse<DayOfWeek>(weekday, true, out var targetWeekday))
            throw new InvalidDataException($"Unknown relative weekday: {weekday}.");
        var days = ((int)targetWeekday - (int)runDate.DayOfWeek + 7) % 7;
        if (days == 0) days = 7;
        var targetDate = runDate.Date.AddDays(days);
        object sourceDateCell = sheet.Range[sourceDateAddress];
        object cell = sheet.Range[address];
        object row = sheet.Range[destinationRange];
        try
        {
            var sourceText = Convert.ToString(((dynamic)sourceDateCell).Text) ?? "";
            var sourceDate = ParseDisplayedDate(sourceText, runDate.Year);
            var text = Convert.ToString(((dynamic)cell).Text) ?? "";
            string updated;
            if (((dynamic)cell).Value2 is double serial && serial > 0)
            {
                ((dynamic)cell).Value2 = targetDate.ToOADate();
                updated = Convert.ToString(((dynamic)cell).Text) ?? targetDate.ToString("d MMM yy");
            }
            else
            {
                updated = Regex.Replace(text, @"\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b",
                    targetDate.ToString("MMM"), RegexOptions.IgnoreCase);
                updated = Regex.Replace(updated, @"\b\d{1,2}(st|nd|rd|th)?\b", OrdinalDay(targetDate.Day),
                    RegexOptions.IgnoreCase);
                if (updated == text)
                    throw new InvalidOperationException($"The copied date in {address} has an unsupported display format: '{text}'.");
                ((dynamic)cell).Value2 = updated;
            }
            object formulas = ((dynamic)row).Formula;
            if (formulas is Array formulaArray)
                for (var formulaRow = 1; formulaRow <= formulaArray.GetLength(0); formulaRow++)
                    for (var formulaColumn = 1; formulaColumn <= formulaArray.GetLength(1); formulaColumn++)
                        if (formulaArray.GetValue(formulaRow, formulaColumn) is string formula &&
                            formula.StartsWith("=", StringComparison.Ordinal))
                            formulaArray.SetValue(RewriteFormulaDate(formula, sourceDate, targetDate),
                                formulaRow, formulaColumn);
            ((dynamic)row).Formula = formulas;
            var actual = Convert.ToString(((dynamic)cell).Text) ?? "";
            if (!actual.Contains(targetDate.Day.ToString(), StringComparison.Ordinal) ||
                !actual.Contains(targetDate.ToString("MMM"), StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"The dynamic date update was not verified in {address}.");
            object verified = ((dynamic)row).Formula;
            if (verified is Array verifiedArray && verifiedArray.Cast<object?>().OfType<string>().Any(formula =>
                    formula.StartsWith("=", StringComparison.Ordinal) &&
                    FormulaContainsDate(formula, sourceDate)))
                throw new InvalidOperationException("One or more copied formulas still reference the prior row's date.");
            SelectRange(sheet, address);
            return actual;
        }
        finally
        {
            Marshal.ReleaseComObject(row); Marshal.ReleaseComObject(cell);
            Marshal.ReleaseComObject(sourceDateCell);
        }

        static DateTime ParseDisplayedDate(string text, int year)
        {
            var match = Regex.Match(text, @"\b(?<month>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(?<day>\d{1,2})(?:st|nd|rd|th)?\b",
                RegexOptions.IgnoreCase);
            if (!match.Success || !DateTime.TryParseExact($"{match.Groups["month"].Value} {match.Groups["day"].Value} {year}",
                    "MMM d yyyy", System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var parsed))
                throw new InvalidOperationException($"The prior row date has an unsupported display format: '{text}'.");
            return parsed;
        }

        static bool FormulaContainsDate(string formula, DateTime date) =>
            Regex.IsMatch(formula, $@"\b{date.Day:00}\s+{Regex.Escape(date.ToString("MMM", System.Globalization.CultureInfo.InvariantCulture))}\b",
                RegexOptions.IgnoreCase);

        static string RewriteFormulaDate(string formula, DateTime from, DateTime to)
        {
            var pattern = $@"\b{from.Day:00}\s+{Regex.Escape(from.ToString("MMM", System.Globalization.CultureInfo.InvariantCulture))}\b";
            return Regex.Replace(formula, pattern,
                $"{to.Day:00} {to.ToString("MMM", System.Globalization.CultureInfo.InvariantCulture)}",
                RegexOptions.IgnoreCase);
        }

    }

    internal static string OrdinalDay(int day)
    {
        if (day is < 1 or > 31) throw new ArgumentOutOfRangeException(nameof(day));
        var twoDigitDay = day.ToString("00", System.Globalization.CultureInfo.InvariantCulture);
        if (day % 100 is 11 or 12 or 13) return twoDigitDay + "th";
        return twoDigitDay + ((day % 10) switch { 1 => "st", 2 => "nd", 3 => "rd", _ => "th" });
    }

    internal sealed record FormulaFill(int FirstRow, int LastRow, string Formula);

    internal static FormulaFill PlanFormulaExtension(dynamic sheet, string column, string adjacent,
        CancellationToken token, int? sourceRow = null)
    {
        if (sourceRow is < 1) throw new InvalidDataException("Formula source row must be positive.");
        var lastRow = ContiguousDataEnd(sheet, adjacent, sourceRow ?? 1, token);
        if (lastRow < 1) throw new InvalidOperationException("Adjacent column is empty; no formula fill was performed.");
        if (sourceRow is { } firstRow)
        {
            if (lastRow < firstRow) throw new InvalidOperationException("Formula source has no adjacent data; no fill was performed.");
            object cell = sheet.Range[$"{column}{firstRow}"];
            try
            {
                if (((dynamic)cell).FormulaR1C1 is not string formula ||
                    !formula.StartsWith("=", StringComparison.Ordinal))
                    throw new InvalidOperationException("The explicit source cell is not a formula; no fill was performed.");
                return new FormulaFill(firstRow, lastRow, formula);
            }
            finally { Marshal.ReleaseComObject(cell); }
        }
        object range = sheet.Range[$"{column}1:{column}{lastRow}"];
        try
        {
            object? formulas = ((dynamic)range).FormulaR1C1;
            for (var row = lastRow; row >= 1; row--)
            {
                token.ThrowIfCancellationRequested();
                var value = formulas is Array array ? array.GetValue(row, 1) : formulas;
                if (value is string formula && formula.StartsWith("=", StringComparison.Ordinal))
                    return new FormulaFill(row, lastRow, formula);
            }
            throw new InvalidOperationException($"Column {column} contains no source formula; fill was not performed.");
        }
        finally { Marshal.ReleaseComObject(range); }
    }

    internal static void VerifyFormulaExtension(dynamic sheet, string column, FormulaFill fill)
    {
        object range = sheet.Range[$"{column}{fill.FirstRow}:{column}{fill.LastRow}"];
        try
        {
            object? formulas = ((dynamic)range).FormulaR1C1;
            for (var row = 0; row <= fill.LastRow - fill.FirstRow; row++)
            {
                var actual = formulas is Array array ? array.GetValue(row + 1, 1) : formulas;
                if (!Equals(actual, fill.Formula))
                    throw new InvalidOperationException($"Relative formula fill was not verified in {column}{fill.FirstRow + row}.");
            }
        }
        finally { Marshal.ReleaseComObject(range); }
    }

    internal static bool HasSheet(dynamic sheet, string name)
    {
        object workbook = sheet.Parent;
        object sheets = ((dynamic)workbook).Sheets;
        try
        {
            var count = (int)((dynamic)sheets).Count;
            for (var index = 1; index <= count; index++)
            {
                object candidate = ((dynamic)sheets)[index];
                try
                {
                    if (string.Equals((string)((dynamic)candidate).Name, name,
                        StringComparison.OrdinalIgnoreCase)) return true;
                }
                finally { Marshal.ReleaseComObject(candidate); }
            }
            return false;
        }
        finally
        {
            Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(workbook);
        }
    }
}
