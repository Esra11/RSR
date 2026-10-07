using System.Text.RegularExpressions;
using System.Windows.Automation;

namespace DesktopSteps;

[Flags]
internal enum TableCapability
{
    CopyRange = 1,
    FormulaExtension = 2,
    PasteValues = 4,
    ChartSource = 8,
    ChartLayout = 16
}

internal sealed record TableCopy(string Range, object?[,] Values);
internal sealed record TablePaste(string Sheet, string Destination, object?[,] Values);

// Advanced table operations are capability based rather than tied to workflow names.
// Applications opt in through an adapter; unsupported capabilities stop safely.
internal interface ITableReplayAdapter
{
    bool Supports(nint window, TableCapability capability);
    Task<TableCopy> CopyUntilEmptyAsync(nint window, ControlRef target, string start, string[] stopValues, CancellationToken token);
    Task<TableCopy> CopyPopulatedColumnsAsync(nint window, ControlRef target, string start, CancellationToken token);
    Task<string> ExtendFormulaAsync(nint window, ControlRef target, string column, string adjacent, CancellationToken token, int? sourceRow = null);
    void PasteValues(nint window, string destination, object?[,] expected);
    void SetChartSource(nint window, string chartName, TablePaste paste);
    void SetLegendLayout(nint window, string chartName, string sheetName, string layout);
    void ExpandLegendOppositePlot(nint window, string chartName);
}

internal static class TableReplayAdapters
{
    private static readonly ITableReplayAdapter[] Adapters = [new ExcelTableReplayAdapter()];

    internal static ITableReplayAdapter Require(nint window, TableCapability capability)
    {
        var matches = Adapters.Where(adapter => window != 0 && adapter.Supports(window, capability)).ToList();
        return matches.Count switch
        {
            1 => matches[0],
            0 => throw new NotSupportedException($"The target program exposes no registered adapter for {capability}. No table input was sent; perform this PDF action manually."),
            _ => throw new InvalidOperationException($"Multiple table adapters support {capability}; no table input was sent.")
        };
    }
}

internal sealed class ExcelTableReplayAdapter : ITableReplayAdapter
{
    public bool Supports(nint window, TableCapability capability) =>
        AutomationElement.FromHandle(window).Current.ClassName == "XLMAIN" &&
        (capability & ~(TableCapability.CopyRange | TableCapability.FormulaExtension | TableCapability.PasteValues | TableCapability.ChartSource | TableCapability.ChartLayout)) == 0;

    public async Task<TableCopy> CopyUntilEmptyAsync(nint window, ControlRef target, string start, string[] stopValues,
        CancellationToken token)
    {
        var match = Regex.Match(start, @"^([A-Z]{1,3})([1-9]\d*)$");
        if (!match.Success || !int.TryParse(match.Groups[2].Value, out var row))
            throw new InvalidDataException("Invalid table-copy start address.");
        var column = match.Groups[1].Value;
        Executor.ActivateWindow(AutomationElement.FromHandle(window));
        await Executor.GoToExcelAddressAsync(start, token);
        _ = await Executor.WaitForFocusedExcelCellAsync(target, column, row, token);
        var end = ExcelNativeSheet.Read<int>(window, sheet => ExcelNativeSheet.ContiguousDataEnd(sheet, column, row, token, stopValues));
        if (end < row) throw new InvalidOperationException($"{start} is a copy boundary; nothing was copied.");
        var range = $"{start}:{column}{end}";
        return await CopyVerifiedRangeAsync(window, range, token);
    }

    public async Task<TableCopy> CopyPopulatedColumnsAsync(nint window, ControlRef target, string start,
        CancellationToken token)
    {
        var match = Regex.Match(start, @"^([A-Z]{1,3})([1-9]\d*):([A-Z]{1,3})\2$");
        if (!match.Success || !int.TryParse(match.Groups[2].Value, out var row))
            throw new InvalidDataException("Invalid populated-column starting range.");
        static int Number(string column) => column.Aggregate(0, (value, letter) => value * 26 + letter - 'A' + 1);
        static string Column(int number)
        {
            var name = "";
            while (number > 0) { number--; name = (char)('A' + number % 26) + name; number /= 26; }
            return name;
        }
        var first = Number(match.Groups[1].Value);
        var last = Number(match.Groups[3].Value);
        if (first >= last || last > 16384) throw new InvalidDataException("Invalid populated-column scope.");
        Executor.ActivateWindow(AutomationElement.FromHandle(window));
        var address = match.Groups[1].Value + row;
        await Executor.GoToExcelAddressAsync(address, token);
        _ = await Executor.WaitForFocusedExcelCellAsync(target, match.Groups[1].Value, row, token);
        var end = ExcelNativeSheet.Read<int>(window, sheet =>
            Enumerable.Range(first, last - first + 1).Min(column =>
                ExcelNativeSheet.ContiguousDataEnd(sheet, Column(column), row, token)));
        if (end < row) throw new InvalidOperationException("The first row contains an empty cell; nothing was copied.");
        return await CopyVerifiedRangeAsync(window, $"{address}:{match.Groups[3].Value}{end}", token);
    }

    private static async Task<TableCopy> CopyVerifiedRangeAsync(nint window, string range, CancellationToken token)
    {
        await Executor.GoToExcelAddressAsync(range, token);
        var values = ExcelNativeSheet.Read<object?[,]>(window, sheet =>
        {
            ExcelNativeSheet.VerifyRangeSelection(sheet, range);
            return ExcelNativeSheet.CaptureSelectionValues(sheet);
        });
        SendKeys.SendWait("^c");
        await Task.Delay(150, token);
        ExcelNativeSheet.Read<bool>(window, sheet =>
        {
            ExcelNativeSheet.VerifyCopyMode(sheet);
            return true;
        });
        return new TableCopy(range, values);
    }

    public async Task<string> ExtendFormulaAsync(nint window, ControlRef target, string column,
        string adjacent, CancellationToken token, int? sourceRow = null)
    {
        if (!Regex.IsMatch(column, @"^[A-Z]{1,3}$") || !Regex.IsMatch(adjacent, @"^[A-Z]{1,3}$"))
            throw new InvalidDataException("Invalid formula-extension columns.");
        Executor.ActivateWindow(AutomationElement.FromHandle(window));
        var fill = ExcelNativeSheet.Read<ExcelNativeSheet.FormulaFill>(window,
            sheet => ExcelNativeSheet.PlanFormulaExtension(sheet, column, adjacent, token, sourceRow));
        var start = $"{column}{fill.FirstRow}";
        ExcelNativeSheet.Read<bool>(window, sheet =>
        {
            ExcelNativeSheet.SelectRange(sheet, start);
            return true;
        });
        await Task.Delay(100, token);
        var range = $"{start}:{column}{fill.LastRow}";
        ExcelNativeSheet.Read<bool>(window, sheet =>
        {
            ExcelNativeSheet.SelectRange(sheet, range);
            return true;
        });
        if (fill.FirstRow < fill.LastRow) SendKeys.SendWait("^d");
        await Task.Delay(150, token);
        ExcelNativeSheet.Read<bool>(window, sheet =>
        {
            ExcelNativeSheet.VerifyFormulaExtension(sheet, column, fill);
            return true;
        });
        return range;
    }

    public void PasteValues(nint window, string destination, object?[,] expected) =>
        ExcelNativeSheet.Read<bool>(window, sheet =>
        {
            ExcelNativeSheet.PasteValues(sheet, destination, expected);
            return true;
        });

    public void SetChartSource(nint window, string chartName, TablePaste paste) =>
        ExcelNativeSheet.Read<bool>(window, sheet =>
        {
            ExcelNativeSheet.SetChartSource(sheet, chartName, paste);
            return true;
        });

    public void SetLegendLayout(nint window, string chartName, string sheetName, string layout) =>
        ExcelNativeSheet.Read<bool>(window, sheet =>
        {
            if ((string)sheet.Name != sheetName)
                throw new InvalidOperationException("The legend's recorded worksheet is not active; no resize was applied.");
            var recorded = System.Text.Json.JsonSerializer.Deserialize<ExcelNativeSheet.LegendLayout>(layout)
                ?? throw new InvalidDataException("Recorded legend layout is missing.");
            ExcelNativeSheet.SetLegendLayout(sheet, chartName, recorded);
            return true;
        });

    public void ExpandLegendOppositePlot(nint window, string chartName) =>
        ExcelNativeSheet.Read<bool>(window, sheet =>
        {
            ExcelNativeSheet.ExpandLegendOppositePlot(sheet, chartName);
            return true;
        });
}
