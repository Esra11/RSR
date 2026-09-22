using System.Diagnostics;
using System.Windows.Automation;

namespace DesktopSteps;

internal static class Automation
{
    private static readonly object TabStripLock = new();
    private static readonly Dictionary<(int Window, string Id), AutomationElement> TabStripCache = [];

    private static AutomationElement? TabStrip(AutomationElement window, string automationId)
    {
        var handle = window.Current.NativeWindowHandle;
        var key = (handle, automationId);
        if (handle != 0)
        {
            lock (TabStripLock)
                if (TabStripCache.TryGetValue(key, out var cached))
                {
                    try
                    {
                        if (cached.FindFirst(TreeScope.Children,
                            new PropertyCondition(AutomationElement.AutomationIdProperty, automationId)) is not null)
                            return cached;
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                        System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
                    TabStripCache.Remove(key);
                }
        }
        var seed = window.FindFirst(TreeScope.Descendants, new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
            new PropertyCondition(AutomationElement.AutomationIdProperty, automationId)));
        if (seed is null) return null;
        var strip = TreeWalker.ControlViewWalker.GetParent(seed);
        if (strip is not null && handle != 0)
            lock (TabStripLock) TabStripCache[key] = strip;
        return strip;
    }

    public static AutomationElement? SelectedTab(AutomationElement window, bool sheetOnly)
    {
        try
        {
            var root = sheetOnly ? TabStrip(window, "SheetTab") : window;
            if (root is null) return null;
            var selected = root.FindAll(sheetOnly ? TreeScope.Children : TreeScope.Descendants,
                new AndCondition(
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new PropertyCondition(SelectionItemPattern.IsSelectedProperty, true)));
            foreach (AutomationElement tab in selected)
                if (!sheetOnly || tab.Current.AutomationId == "SheetTab") return tab;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
        return null;
    }
    public static string? FieldLabel(AutomationElement element)
    {
        try
        {
            if (element.Current.LabeledBy is { } labeled &&
                !string.IsNullOrWhiteSpace(labeled.Current.Name)) return labeled.Current.Name;
            var bounds = element.Current.BoundingRectangle;
            if (bounds.IsEmpty) return null;
            var parent = TreeWalker.ControlViewWalker.GetParent(element);
            if (parent is null) return null;
            var parentBounds = parent.Current.BoundingRectangle;
            var screen = Screen.FromPoint(new System.Drawing.Point((int)bounds.Left, (int)bounds.Top)).Bounds;
            if (parentBounds.Width * parentBounds.Height > screen.Width * screen.Height * 0.8)
                return null;
            var labels = parent.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Text));
            return labels.Cast<AutomationElement>()
                .Select(label => (Name: label.Current.Name, Bounds: label.Current.BoundingRectangle))
                .Where(label => !string.IsNullOrWhiteSpace(label.Name) && !label.Bounds.IsEmpty &&
                    label.Bounds.Right <= bounds.Left + 8 &&
                    Math.Abs((label.Bounds.Top + label.Bounds.Bottom) / 2 -
                        (bounds.Top + bounds.Bottom) / 2) <= Math.Max(18, bounds.Height))
                .OrderBy(label => bounds.Left - label.Bounds.Right).FirstOrDefault().Name;
        }
        catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
            System.Runtime.InteropServices.COMException) { Trace.WriteLine(ex); return null; }
    }

    public static AutomationElement? VisibleGridCell(ControlRef target, int visibleRow)
    {
        if (visibleRow < 1 || string.IsNullOrWhiteSpace(target.AutomationId)) return null;
        var match = System.Text.RegularExpressions.Regex.Match(target.AutomationId, @"^(?<column>[A-Za-z]+)\d+$");
        if (!match.Success) return null;
        var column = match.Groups["column"].Value;
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(target.Process, StringComparison.OrdinalIgnoreCase) ||
                    !window.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase)) continue;
                var items = window.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem));
                var visible = new List<(AutomationElement Element, double Top)>();
                foreach (AutomationElement item in items)
                {
                    var current = item.Current;
                    if (current.IsOffscreen ||
                        !System.Text.RegularExpressions.Regex.IsMatch(current.AutomationId,
                            "^" + System.Text.RegularExpressions.Regex.Escape(column) + @"\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                        continue;
                    var bounds = current.BoundingRectangle;
                    if (bounds.IsEmpty || bounds.Width < 4 || bounds.Height < 4) continue;
                    visible.Add((item, bounds.Top));
                }
                var ordered = visible.OrderBy(item => item.Top).ToList();
                return visibleRow <= ordered.Count ? ordered[visibleRow - 1].Element : null;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                System.Runtime.InteropServices.COMException) { Trace.WriteLine(ex); }
        }
        return null;
    }

    public static int? VisibleGridOrdinal(ControlRef target)
    {
        if (string.IsNullOrWhiteSpace(target.AutomationId)) return null;
        var match = System.Text.RegularExpressions.Regex.Match(target.AutomationId, @"^(?<column>[A-Za-z]+)\d+$");
        if (!match.Success) return null;
        var column = match.Groups["column"].Value;
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(target.Process, StringComparison.OrdinalIgnoreCase) ||
                    !window.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase)) continue;
                var items = window.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem));
                var visible = new List<(string Id, double Top)>();
                foreach (AutomationElement item in items)
                {
                    var current = item.Current;
                    if (current.IsOffscreen || !System.Text.RegularExpressions.Regex.IsMatch(current.AutomationId,
                        "^" + System.Text.RegularExpressions.Regex.Escape(column) + @"\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                        continue;
                    var bounds = current.BoundingRectangle;
                    if (bounds.IsEmpty || bounds.Width < 4 || bounds.Height < 4) continue;
                    visible.Add((current.AutomationId, bounds.Top));
                }
                var ordered = visible.OrderBy(item => item.Top).ToList();
                var position = ordered.FindIndex(item => item.Id.Equals(target.AutomationId, StringComparison.OrdinalIgnoreCase));
                return position < 0 ? null : position + 1;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                System.Runtime.InteropServices.COMException) { Trace.WriteLine(ex); }
        }
        return null;
    }

    public static AutomationElement ScrollableAncestor(AutomationElement element)
    {
        var current = element;
        for (var depth = 0; depth < 8; depth++)
        {
            if (current.TryGetCurrentPattern(ScrollPattern.Pattern, out _)) return current;
            var parent = TreeWalker.ControlViewWalker.GetParent(current);
            if (parent is null || parent == AutomationElement.RootElement) break;
            current = parent;
        }
        return element;
    }

    public static AutomationElement? ListItemAncestor(AutomationElement element)
    {
        var current = element;
        for (var depth = 0; depth < 6; depth++)
        {
            if (current.Current.ControlType == ControlType.ListItem) return current;
            var parent = TreeWalker.ControlViewWalker.GetParent(current);
            if (parent is null || parent == AutomationElement.RootElement) break;
            current = parent;
        }
        return null;
    }

    public static AutomationElement? FirstListItemUnderHeader(AutomationElement header)
    {
        var parent = TreeWalker.ControlViewWalker.GetParent(header);
        for (var level = 0; level < 4 && parent is not null; level++)
        {
            var row = parent.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
            if (row is not null) return row;
            parent = TreeWalker.ControlViewWalker.GetParent(parent);
        }
        return null;
    }

    public static AutomationElement? FirstListItemInWindow(ControlRef target)
    {
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!System.Diagnostics.Process.GetProcessById(window.Current.ProcessId).ProcessName
                    .Equals(target.Process, StringComparison.OrdinalIgnoreCase) ||
                    !window.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase)) continue;
                var rows = window.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem));
                foreach (AutomationElement row in rows)
                {
                    var parent = TreeWalker.ControlViewWalker.GetParent(row);
                    if (target.ParentName is null || parent?.Current.Name == target.ParentName) return row;
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                System.Runtime.InteropServices.COMException) { }
        }
        return null;
    }

    public static string ScrollBarAxis(AutomationElement element)
    {
        var bounds = element.Current.BoundingRectangle;
        return bounds.Width > bounds.Height ? "Horizontal" : "Vertical";
    }

    public static AutomationElement? FindScrollBar(ControlRef target, string axis)
    {
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(target.Process) &&
                    !Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(target.Process, StringComparison.OrdinalIgnoreCase)) continue;
                if (!string.IsNullOrWhiteSpace(target.Window) &&
                    !window.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase)) continue;
                var bars = window.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ScrollBar));
                foreach (AutomationElement bar in bars)
                    if (ScrollBarAxis(bar) == axis && bar.TryGetCurrentPattern(RangeValuePattern.Pattern, out _)) return bar;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or NullReferenceException) { }
        }
        return null;
    }
    public static AutomationElement? ActionableAncestor(AutomationElement element)
    {
        var current = element;
        for (var depth = 0; depth < 5 && current is not null; depth++)
        {
            if (current.TryGetCurrentPattern(InvokePattern.Pattern, out _) ||
                current.TryGetCurrentPattern(TogglePattern.Pattern, out _) ||
                current.TryGetCurrentPattern(SelectionItemPattern.Pattern, out _) ||
                current.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out _)) return current;
            current = TreeWalker.ControlViewWalker.GetParent(current);
        }
        return null;
    }

    public static HashSet<string> SiblingTabNames(AutomationElement tab)
    {
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var parent = TreeWalker.ControlViewWalker.GetParent(tab);
        if (parent is null) return names;
        var child = TreeWalker.ControlViewWalker.GetFirstChild(parent);
        while (child is not null)
        {
            try
            {
                if (child.Current.ControlType == ControlType.TabItem &&
                    !string.IsNullOrWhiteSpace(child.Current.Name)) names.Add(child.Current.Name);
            }
            catch (ElementNotAvailableException) { }
            child = TreeWalker.ControlViewWalker.GetNextSibling(child);
        }
        return names;
    }
    public static ControlRef? Describe(AutomationElement? element)
    {
        if (element is null) return null;
        try
        {
            var c = element.Current;
            var root = element;
            AutomationElement? parent = null;
            while (true)
            {
                var next = TreeWalker.ControlViewWalker.GetParent(root);
                if (next is null || next == AutomationElement.RootElement) break;
                parent = next;
                root = next;
            }
            string? process = null;
            try { process = Process.GetProcessById(c.ProcessId).ProcessName; } catch { }
            string? parentName = null;
            try { parentName = TreeWalker.ControlViewWalker.GetParent(element)?.Current.Name; } catch { }
            var orientation = c.ControlType == ControlType.ScrollBar ? ScrollBarAxis(element) : c.Orientation.ToString();
            return new(process, root.Current.Name, c.AutomationId, c.Name,
                c.ControlType.ProgrammaticName, c.ClassName, parentName, orientation,
                ProcessId: root.Current.ProcessId);
        }
        catch (ElementNotAvailableException) { return null; }
    }

    public static string? State(AutomationElement? element)
    {
        if (element is null) return null;
        try
        {
            var c = element.Current;
            var value = element.TryGetCurrentPattern(ValuePattern.Pattern, out var p)
                ? ((ValuePattern)p).Current.Value : "";
            var toggle = element.TryGetCurrentPattern(TogglePattern.Pattern, out var t)
                ? ((TogglePattern)t).Current.ToggleState.ToString() : "";
            return $"enabled={c.IsEnabled};name={c.Name};help={c.HelpText};value={value};toggle={toggle}";
        }
        catch (ElementNotAvailableException) { return null; }
    }

    public static string? ScrollState(AutomationElement? element)
    {
        var current = element;
        for (var depth = 0; depth < 6 && current is not null; depth++)
        {
            try
            {
                if (current.TryGetCurrentPattern(RangeValuePattern.Pattern, out var range))
                    return "range:" + ((RangeValuePattern)range).Current.Value.ToString(System.Globalization.CultureInfo.InvariantCulture);
                if (current.TryGetCurrentPattern(ScrollPattern.Pattern, out var scroll))
                {
                    var state = ((ScrollPattern)scroll).Current;
                    return $"scroll:{state.HorizontalScrollPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)},{state.VerticalScrollPercent.ToString(System.Globalization.CultureInfo.InvariantCulture)}";
                }
                current = TreeWalker.ControlViewWalker.GetParent(current);
            }
            catch (ElementNotAvailableException) { return null; }
        }
        return null;
    }

    public static AutomationElement? Resolve(ControlRef target, int maxNodes = 2500)
    {
        if (string.IsNullOrWhiteSpace(target.AutomationId) && string.IsNullOrWhiteSpace(target.Name) &&
            !(target.ControlType == "ControlType.ScrollBar" && !string.IsNullOrWhiteSpace(target.ClassName))) return null;
        if (target.ControlType == "ControlType.TabItem" &&
            !string.IsNullOrWhiteSpace(target.AutomationId) && !string.IsNullOrWhiteSpace(target.Name))
            return ResolveNamedTab(target);
        var desktop = AutomationElement.RootElement;
        var windows = desktop.FindAll(TreeScope.Children, Condition.TrueCondition);
        var candidates = windows.Cast<AutomationElement>().OrderByDescending(window =>
        {
            try { return window.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase); }
            catch (Exception ex) when (ex is ElementNotAvailableException or System.Runtime.InteropServices.COMException)
            { return false; }
        });
        foreach (AutomationElement window in candidates)
        {
            try
            {
                var w = window.Current;
                if (!string.IsNullOrWhiteSpace(target.Process))
                {
                    var name = Process.GetProcessById(w.ProcessId).ProcessName;
                    if (!name.Equals(target.Process, StringComparison.OrdinalIgnoreCase)) continue;
                }
                // Large native dialogs can put an Edit beyond the bounded Control View
                // traversal. Its recorded AutomationId remains the strongest locator.
                if (target.ControlType == "ControlType.Edit" &&
                    !string.IsNullOrWhiteSpace(target.AutomationId))
                {
                    var direct = window.FindFirst(TreeScope.Descendants, new AndCondition(
                        new PropertyCondition(AutomationElement.AutomationIdProperty, target.AutomationId),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit)));
                    if (direct is not null && Matches(direct, target) &&
                        direct.Current.IsEnabled && !direct.Current.IsOffscreen)
                        return direct;
                }
                // Document titles often change after each character; prefer the old title but
                // allow the process and control identity to locate the same editor later.
                var queue = new Queue<AutomationElement>();
                queue.Enqueue(window);
                var examined = 0;
                while (queue.Count > 0 && examined++ < maxNodes)
                {
                    var node = queue.Dequeue();
                    if (Matches(node, target)) return node;
                    var child = TreeWalker.ControlViewWalker.GetFirstChild(node);
                    while (child is not null)
                    {
                        queue.Enqueue(child);
                        child = TreeWalker.ControlViewWalker.GetNextSibling(child);
                    }
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                NullReferenceException or System.Runtime.InteropServices.COMException)
            { System.Diagnostics.Trace.WriteLine($"UIA window traversal failed: {ex}"); }
        }
        return null;
    }

    private static AutomationElement? ResolveNamedTab(ControlRef target)
    {
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        var exactWindowExists = windows.Cast<AutomationElement>().Any(window =>
        {
            try { return window.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase) &&
                Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(target.Process,
                    StringComparison.OrdinalIgnoreCase); }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                System.Runtime.InteropServices.COMException) { return false; }
        });
        var condition = new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
            new PropertyCondition(AutomationElement.AutomationIdProperty, target.AutomationId),
            new PropertyCondition(AutomationElement.NameProperty, target.Name));
        AutomationElement? unique = null;
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(target.Process,
                    StringComparison.OrdinalIgnoreCase) ||
                    exactWindowExists && !window.Current.Name.Equals(target.Window,
                        StringComparison.OrdinalIgnoreCase)) continue;
                var strip = TabStrip(window, target.AutomationId!);
                if (strip is null) continue;
                var matches = strip.FindAll(TreeScope.Children, condition);
                foreach (AutomationElement match in matches)
                {
                    if (unique is not null) return null; // Never select an ambiguous tab.
                    unique = match;
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                InvalidOperationException or System.Runtime.InteropServices.COMException)
            { System.Diagnostics.Trace.WriteLine(ex); }
        }
        return unique;
    }

    public static string TabLookupDiagnostics(ControlRef target)
    {
        var found = new List<string>();
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        var condition = new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
            new PropertyCondition(AutomationElement.AutomationIdProperty, target.AutomationId ?? ""));
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(target.Process,
                    StringComparison.OrdinalIgnoreCase)) continue;
                var strip = TabStrip(window, target.AutomationId ?? "");
                if (strip is null)
                {
                    found.Add($"{window.Current.Name}: tab strip not exposed (node cap not used)");
                    continue;
                }
                var tabs = strip.FindAll(TreeScope.Children, condition);
                var names = tabs.Cast<AutomationElement>().Select(tab => tab.Current.Name)
                    .Where(name => !string.IsNullOrWhiteSpace(name)).Distinct().Take(30);
                found.Add($"{window.Current.Name}: checked {tabs.Count} tab-strip children, " +
                    $"node cap not used; [{string.Join(", ", names)}]");
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                InvalidOperationException or System.Runtime.InteropServices.COMException)
            { System.Diagnostics.Trace.WriteLine(ex); }
        }
        return found.Count == 0 ? "No matching application window exposed tabs through UI Automation."
            : string.Join("; ", found);
    }

    public static AutomationElement? ResolveMenuItem(ControlRef target)
    {
        if (string.IsNullOrWhiteSpace(target.Name)) return null;
        var condition = new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem),
            new PropertyCondition(AutomationElement.NameProperty, target.Name));
        if (!string.IsNullOrWhiteSpace(target.AutomationId) && !string.IsNullOrWhiteSpace(target.ParentName))
        {
            // Repeated controls such as worksheet filter buttons have identical names and IDs.
            // Resolve within the recorded parent instead of taking the first global match.
            var roots = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
            foreach (AutomationElement root in roots)
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(target.Process) &&
                        !Process.GetProcessById(root.Current.ProcessId).ProcessName.Equals(target.Process, StringComparison.OrdinalIgnoreCase)) continue;
                    if (!string.IsNullOrWhiteSpace(target.Window) &&
                        !root.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase)) continue;
                    var parent = root.FindFirst(TreeScope.Descendants,
                        new PropertyCondition(AutomationElement.NameProperty, target.ParentName));
                    if (parent is null) continue;
                    var candidates = parent.FindAll(TreeScope.Descendants, condition);
                    foreach (AutomationElement candidate in candidates)
                        if (candidate.Current.AutomationId == target.AutomationId &&
                            TreeWalker.ControlViewWalker.GetParent(candidate)?.Current.Name == target.ParentName)
                            return candidate;
                }
                catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or InvalidOperationException or
                    System.Runtime.InteropServices.COMException) { System.Diagnostics.Trace.WriteLine(ex); }
            }
            return null; // A different column's identical button is never a safe fallback.
        }
        try
        {
            var focused = AutomationElement.FocusedElement;
            for (var depth = 0; focused is not null && depth < 6; depth++)
            {
                if (focused.Current.ControlType == ControlType.MenuItem && focused.Current.Name == target.Name)
                    return focused;
                var nearby = focused.FindFirst(TreeScope.Descendants, condition);
                if (nearby is not null) return nearby;
                focused = TreeWalker.ControlViewWalker.GetParent(focused);
            }
        }
        catch (Exception e) when (e is ElementNotAvailableException or InvalidOperationException) { }
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(target.Process) &&
                    !Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(target.Process, StringComparison.OrdinalIgnoreCase))
                    continue;
                var item = window.FindFirst(TreeScope.Descendants, condition);
                if (item is not null) return item;
            }
            catch (Exception e) when (e is ElementNotAvailableException or ArgumentException or InvalidOperationException) { }
        }
        return null;
    }

    public static AutomationElement? FindSelectedListRow(ControlRef target)
    {
        var condition = new AndCondition(
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem),
            new PropertyCondition(SelectionItemPattern.IsSelectedProperty, true));
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(target.Process) &&
                    !Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(target.Process, StringComparison.OrdinalIgnoreCase))
                    continue;
                if (!string.IsNullOrWhiteSpace(target.Window) &&
                    !window.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase)) continue;
                var row = window.FindFirst(TreeScope.Descendants, condition);
                if (row is not null) return row;
            }
            catch (Exception e) when (e is ElementNotAvailableException or ArgumentException or InvalidOperationException) { }
        }
        return null;
    }

    public static AutomationElement? UniqueSelectableByName(string processName, string name)
    {
        AutomationElement? match = null;
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(
                    processName, StringComparison.OrdinalIgnoreCase)) continue;
                var named = window.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.NameProperty, name));
                foreach (AutomationElement candidate in named)
                {
                    try
                    {
                        if (candidate.Current.IsOffscreen || !candidate.Current.IsEnabled ||
                            !candidate.TryGetCurrentPattern(SelectionItemPattern.Pattern, out _)) continue;
                        if (match is not null) return null;
                        match = candidate;
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                        System.Runtime.InteropServices.COMException) { Trace.WriteLine(ex); }
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                InvalidOperationException or System.Runtime.InteropServices.COMException) { Trace.WriteLine(ex); }
        }
        return match;
    }

    public static AutomationElement? UniqueVisibleItemByName(string processName, string name)
    {
        AutomationElement? match = null;
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!Process.GetProcessById(window.Current.ProcessId).ProcessName.Equals(
                    processName, StringComparison.OrdinalIgnoreCase)) continue;
                var named = window.FindAll(TreeScope.Descendants, new AndCondition(
                    new PropertyCondition(AutomationElement.NameProperty, name),
                    new OrCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.DataItem),
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem))));
                foreach (AutomationElement candidate in named)
                {
                    try
                    {
                        if (candidate.Current.IsOffscreen || !candidate.Current.IsEnabled) continue;
                        if (match is not null) return null;
                        match = candidate;
                    }
                    catch (Exception ex) when (ex is ElementNotAvailableException or InvalidOperationException or
                        System.Runtime.InteropServices.COMException) { Trace.WriteLine(ex); }
                }
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or
                InvalidOperationException or System.Runtime.InteropServices.COMException) { Trace.WriteLine(ex); }
        }
        return match;
    }

    public static bool WindowExists(ControlRef target)
    {
        if (string.IsNullOrWhiteSpace(target.Window)) return true;
        var windows = AutomationElement.RootElement.FindAll(TreeScope.Children, Condition.TrueCondition);
        foreach (AutomationElement window in windows)
        {
            try
            {
                if (!window.Current.Name.Equals(target.Window, StringComparison.OrdinalIgnoreCase)) continue;
                if (string.IsNullOrWhiteSpace(target.Process)) return true;
                var process = Process.GetProcessById(window.Current.ProcessId).ProcessName;
                if (process.Equals(target.Process, StringComparison.OrdinalIgnoreCase)) return true;
            }
            catch (Exception ex) when (ex is ElementNotAvailableException or ArgumentException or NullReferenceException) { }
        }
        return false;
    }

    private static bool Matches(AutomationElement node, ControlRef target)
    {
        try
        {
            var c = node.Current;
            if (!string.IsNullOrWhiteSpace(target.AutomationId) && c.AutomationId != target.AutomationId) return false;
            // In some native edit controls UIA Name is the current text value.
            // The recorded name is stale as soon as the user edits the field;
            // prefer its stable AutomationId, type, and class in that case.
            var editableWithStableId = !string.IsNullOrWhiteSpace(target.AutomationId) &&
                (c.ControlType == ControlType.Edit ||
                 node.TryGetCurrentPattern(ValuePattern.Pattern, out _) ||
                 target.ClassName == "EDTBX");
            if (!editableWithStableId && !string.IsNullOrWhiteSpace(target.Name) && c.Name != target.Name) return false;
            if (!string.IsNullOrWhiteSpace(target.ControlType) && c.ControlType.ProgrammaticName != target.ControlType) return false;
            if (!string.IsNullOrWhiteSpace(target.ClassName) && c.ClassName != target.ClassName) return false;
            if (!string.IsNullOrWhiteSpace(target.Orientation) && c.ControlType == ControlType.ScrollBar)
            {
                var orientation = ScrollBarAxis(node);
                if (!orientation.Equals(target.Orientation, StringComparison.OrdinalIgnoreCase)) return false;
            }
            return !string.IsNullOrWhiteSpace(target.AutomationId) || !string.IsNullOrWhiteSpace(target.Name) ||
                target.ControlType == "ControlType.ScrollBar";
        }
        catch (ElementNotAvailableException) { return false; }
    }
}
