using System.Runtime.InteropServices;

namespace DesktopSteps;

// Fallback for legacy controls that expose an MSAA action but no UIA command.
internal static class MsaaActions
{
    internal sealed record Observation(string? Name, int? Role, int? State,
        int X, int Y, int Width, int Height);
    private static readonly Guid AccessibleId = new("618736e0-3c3d-11cf-810c-00aa00389b71");
    private const uint ObjIdClient = 0xFFFFFFFC;
    private const uint ObjIdMenu = 0xFFFFFFFD;

    public static bool TryInvokeUnique(nint window, string name, int role)
    {
        try
        {
            var iid = AccessibleId;
            if (AccessibleObjectFromWindow(window, ObjIdClient, ref iid, out var root) != 0 ||
                root is not Accessibility.IAccessible accessible) return false;
            var matches = new List<(Accessibility.IAccessible Parent, object Child)>();
            var visited = 0;
            Visit(accessible, 0, name, role, matches, ref visited);
            if (matches.Count != 1) return false;
            var (parent, child) = matches[0];
            parent.accDoDefaultAction(child);
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.WriteLine(ex);
            return false;
        }
    }

    public static bool TryInvokeUniqueMenu(nint owner, IReadOnlyList<nint> popups, string name)
    {
        try
        {
            var matches = new List<(Accessibility.IAccessible Parent, object Child)>();
            var roots = popups.Distinct()
                .SelectMany(window => new[] { (Window: window, ObjectId: ObjIdClient),
                    (Window: window, ObjectId: ObjIdMenu) })
                .Append((Window: owner, ObjectId: ObjIdMenu));
            foreach (var (window, objectId) in roots)
            {
                var iid = AccessibleId;
                if (AccessibleObjectFromWindow(window, objectId, ref iid, out var root) != 0 ||
                    root is not Accessibility.IAccessible accessible) continue;
                var visited = 0;
                var found = new List<(Accessibility.IAccessible Parent, object Child)>();
                var complete = true;
                Visit(accessible, 0, name, 0x0C, found, ref visited, () => complete = false);
                // Do not claim uniqueness after an incomplete traversal.
                if (!complete || visited > 500 || found.Count > 1) return false;
                matches.AddRange(found);
                matches = matches.Distinct().ToList();
                if (matches.Count > 1) return false;
            }
            if (matches.Count != 1) return false;
            var (parent, child) = matches[0];
            parent.accDoDefaultAction(child);
            return true;
        }
        catch (Exception ex) when (ex is COMException or InvalidOperationException or ArgumentException)
        {
            System.Diagnostics.Trace.WriteLine(ex);
            return false;
        }
    }

    public static bool TrySelectUnique(nint window, string name)
    {
        try
        {
            var iid = AccessibleId;
            if (AccessibleObjectFromWindow(window, ObjIdClient, ref iid, out var root) != 0 ||
                root is not Accessibility.IAccessible accessible) return false;
            var matches = new List<(Accessibility.IAccessible Parent, object Child)>();
            var visited = 0;
            Visit(accessible, 0, name, 0x22, matches, ref visited);
            if (matches.Count != 1) return false;
            var (parent, child) = matches[0];
            parent.accSelect(0x2, child); // SELFLAG_TAKESELECTION
            return parent.get_accState(child) is int state && (state & 0x2) != 0;
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); return false; }
    }

    public static string? MenuParentAtPoint(int x, int y)
    {
        try
        {
            if (AccessibleObjectFromPoint(new Point { X = x, Y = y }, out var accessible,
                out var child) != 0 || accessible is null ||
                accessible.get_accRole(child) is not int role || role != 0x0C) return null;
            // Simple MSAA children share their menu's IAccessible object.
            var parent = child is int id && id != 0
                ? accessible : accessible.accParent as Accessibility.IAccessible;
            for (var depth = 0; parent is not null && depth < 4; depth++)
            {
                if (parent.get_accRole(0) is int parentRole && parentRole is 0x0B or 0x0C &&
                    parent.get_accName(0) is { Length: > 0 } name)
                    return name.Replace("&", "").Trim();
                parent = parent.accParent as Accessibility.IAccessible;
            }
        }
        catch (Exception ex) when (ex is COMException or InvalidOperationException or ArgumentException)
        { System.Diagnostics.Trace.WriteLine(ex); }
        return null;
    }

    public static bool IsSelectedListItemAtPoint(int x, int y, string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return false;
        try
        {
            if (AccessibleObjectFromPoint(new Point { X = x, Y = y }, out var accessible,
                out var child) != 0 || accessible is null) return false;
            return accessible.get_accRole(child) is int role && role == 0x22 &&
                string.Equals(accessible.get_accName(child)?.Replace("&", "").Trim(), name,
                    StringComparison.OrdinalIgnoreCase) &&
                accessible.get_accState(child) is int state && (state & 0x2) != 0;
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); return false; }
    }

    internal static int? TreeItemStateAtPoint(int x, int y, string name)
    {
        try
        {
            if (AccessibleObjectFromPoint(new Point { X = x, Y = y }, out var accessible,
                out var child) != 0 || accessible is null) return null;
            if (accessible.get_accRole(child) is not int role || role is not (0x24 or 0x2C) ||
                !string.Equals(accessible.get_accName(child), name, StringComparison.Ordinal))
                return null;
            return accessible.get_accState(child) is int state ? state : null;
        }
        catch (Exception ex) when (ex is COMException or InvalidOperationException or ArgumentException)
        { System.Diagnostics.Trace.WriteLine(ex); return null; }
    }

    internal static string? FilterItemNameAtPoint(int x, int y)
    {
        try
        {
            if (AccessibleObjectFromPoint(new Point { X = x, Y = y }, out var accessible,
                    out var child) != 0 || accessible is null ||
                accessible.get_accRole(child) is not int role || role is not (0x24 or 0x2C) ||
                accessible.get_accName(child) is not { Length: > 0 } name ||
                accessible.get_accState(child) is not int state ||
                (state & (0x8000 | 0x10000)) != 0) return null;
            accessible.accLocation(out var left, out var top, out var width, out var height, child);
            if (width <= 0 || height <= 0 || x < left || x >= left + width ||
                y < top || y >= top + height) return null;
            var parent = child is int id && id != 0
                ? accessible : accessible.accParent as Accessibility.IAccessible;
            for (var depth = 0; parent is not null && depth < 5; depth++)
            {
                if (parent.get_accName(0) == "Manual Filter")
                    return name;
                parent = parent.accParent as Accessibility.IAccessible;
            }
        }
        catch (Exception ex) when (ex is COMException or InvalidOperationException or ArgumentException)
        { System.Diagnostics.Trace.WriteLine(ex); }
        return null;
    }

    public static bool IsSubmenuAtPoint(int x, int y, string? name)
    {
        const int StateSystemHasPopup = 0x40000000;
        try
        {
            if (AccessibleObjectFromPoint(new Point { X = x, Y = y }, out var accessible,
                    out var child) != 0 || accessible is null) return false;
            return accessible.get_accRole(child) is int role && role == 0x0C &&
                string.Equals(accessible.get_accName(child)?.Replace("&", "").Trim(), name,
                    StringComparison.OrdinalIgnoreCase) &&
                accessible.get_accState(child) is int state && (state & StateSystemHasPopup) != 0;
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); return false; }
    }

    public static IReadOnlyList<Observation> Probe(nint window) => Probe(window, ObjIdClient);

    public static IReadOnlyList<Observation> ProbeMenu(nint window) => Probe(window, ObjIdMenu);

    private static IReadOnlyList<Observation> Probe(nint window, uint objectId)
    {
        var result = new List<Observation>();
        try
        {
            var iid = AccessibleId;
            if (AccessibleObjectFromWindow(window, objectId, ref iid, out var root) == 0 &&
                root is Accessibility.IAccessible accessible)
            {
                var visited = 0;
                VisitProbe(accessible, 0, result, ref visited);
            }
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
        return result;
    }

    public static (string Name, string Type, int X, int Y, int Width, int Height)? CommandAtPoint(
        nint window, int x, int y)
    {
        return DirectCommandAtPoint(x, y) ?? CommandFromObservations(Probe(window), x, y);
    }

    // Ask MSAA for the object directly beneath the pointer before walking the
    // dialog tree. Some legacy dialogs expose the clicked command here but put
    // it beyond the bounded descendant traversal used by Probe().
    private static (string Name, string Type, int X, int Y, int Width, int Height)? DirectCommandAtPoint(
        int x, int y)
    {
        try
        {
            if (AccessibleObjectFromPoint(new Point { X = x, Y = y }, out var accessible,
                out var child) != 0 || accessible is null) return null;
            var role = accessible.get_accRole(child) is int value ? value : 0;
            if (role is not (0x0C or 0x2B or 0x25 or 0x2C or 0x2D)) return null;
            if (accessible.get_accState(child) is int flags &&
                (flags & (0x1 | 0x8000 | 0x10000)) != 0) return null;
            var name = accessible.get_accName(child)?.Replace("&", "").Trim();
            if (string.IsNullOrWhiteSpace(name)) return null;
            accessible.accLocation(out var left, out var top, out var width, out var height, child);
            if (width <= 0 || height <= 0 || x < left || x >= left + width ||
                y < top || y >= top + height) return null;
            var type = role switch
            {
                0x0C => "ControlType.MenuItem",
                0x25 => "ControlType.TabItem",
                0x2C => "ControlType.CheckBox",
                0x2D => "ControlType.RadioButton",
                _ => "ControlType.Button"
            };
            return (name, type, left, top, width, height);
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); return null; }
    }

    public static (string Name, string Type, int X, int Y, int Width, int Height)? MenuCommandAtPoint(
        nint window, int x, int y)
    {
        var fromMenu = CommandFromObservations(Probe(window, ObjIdMenu), x, y);
        return fromMenu is { Type: "ControlType.MenuItem" }
            ? fromMenu : CommandAtPoint(window, x, y);
    }

    public static (string Name, string Type, int X, int Y, int Width, int Height)? CommandFromObservations(
        IEnumerable<Observation> observations, int x, int y)
    {
        var candidates = observations.Where(item => item.Name is { Length: > 0 } &&
            item.Role is 0x0C or 0x2B or 0x25 or 0x2C or 0x2D && item.Width > 0 && item.Height > 0 &&
            (item.State is not int flags || (flags & (0x1 | 0x8000 | 0x10000)) == 0) &&
            x >= item.X && x < item.X + item.Width && y >= item.Y && y < item.Y + item.Height)
            .OrderBy(item => (long)item.Width * item.Height).ToArray();
        if (candidates.Length == 0) return null;
        var smallestArea = (long)candidates[0].Width * candidates[0].Height;
        var smallest = candidates.Where(item => (long)item.Width * item.Height == smallestArea)
            .Select(item => (Name: item.Name!.Replace("&", "").Trim(), item.Role))
            .Distinct().ToArray();
        if (smallest.Length != 1 || string.IsNullOrWhiteSpace(smallest[0].Name)) return null;
        var selected = candidates[0];
        var type = selected.Role switch
        {
            0x0C => "ControlType.MenuItem",
            0x25 => "ControlType.TabItem",
            0x2C => "ControlType.CheckBox",
            0x2D => "ControlType.RadioButton",
            _ => "ControlType.Button"
        };
        return (smallest[0].Name, type,
            selected.X, selected.Y, selected.Width, selected.Height);
    }

    private static void VisitProbe(Accessibility.IAccessible accessible, int depth,
        List<Observation> result, ref int visited)
    {
        if (depth > 8 || visited++ > 500 || result.Count >= 250) return;
        Observe(accessible, 0, result);
        int count;
        try { count = Math.Min(accessible.accChildCount, 500); }
        catch { return; }
        for (var id = 1; id <= count && visited <= 500 && result.Count < 250; id++)
        {
            object? child = null;
            try { child = accessible.get_accChild(id); }
            catch { }
            if (child is Accessibility.IAccessible nested)
                VisitProbe(nested, depth + 1, result, ref visited);
            else
            {
                visited++;
                Observe(accessible, id, result);
            }
        }
    }

    private static void Observe(Accessibility.IAccessible accessible, object child,
        List<Observation> result)
    {
        try
        {
            var name = accessible.get_accName(child);
            var role = accessible.get_accRole(child);
            var state = accessible.get_accState(child);
            if (string.IsNullOrWhiteSpace(name)) return;
            accessible.accLocation(out var x, out var y, out var width, out var height, child);
            result.Add(new Observation(name, role is int value ? value : null,
                state is int flags ? flags : null, x, y, width, height));
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
    }

    private static void Visit(Accessibility.IAccessible accessible, int depth, string name,
        int role, List<(Accessibility.IAccessible Parent, object Child)> matches, ref int visited,
        Action? truncated = null)
    {
        if (depth > 8 || visited++ > 500)
        {
            truncated?.Invoke();
            return;
        }
        if (matches.Count > 1) return;
        Check(accessible, 0, name, role, matches);
        int count;
        try
        {
            count = accessible.accChildCount;
            if (count > 500) truncated?.Invoke();
            count = Math.Min(count, 500);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Trace.WriteLine(ex);
            truncated?.Invoke();
            return;
        }
        for (var id = 1; id <= count && visited <= 500 && matches.Count <= 1; id++)
        {
            object? child = null;
            try { child = accessible.get_accChild(id); }
            catch { }
            if (child is Accessibility.IAccessible nested)
                Visit(nested, depth + 1, name, role, matches, ref visited, truncated);
            else
            {
                visited++;
                Check(accessible, id, name, role, matches);
            }
        }
    }

    private static void Check(Accessibility.IAccessible accessible, object child, string name,
        int role, List<(Accessibility.IAccessible Parent, object Child)> matches)
    {
        try
        {
            if (accessible.get_accRole(child) is not int foundRole || foundRole != role ||
                !string.Equals(accessible.get_accName(child)?.Replace("&", "").Trim(), name,
                    StringComparison.OrdinalIgnoreCase)) return;
            if (accessible.get_accState(child) is int flags &&
                (flags & (0x1 | 0x8000 | 0x10000)) != 0) return;
            accessible.accLocation(out _, out _, out var width, out var height, child);
            if (width > 0 && height > 0) matches.Add((accessible, child));
        }
        catch (Exception ex) { System.Diagnostics.Trace.WriteLine(ex); }
    }

    [DllImport("oleacc.dll", PreserveSig = true)]
    private static extern int AccessibleObjectFromWindow(nint window, uint objectId,
        ref Guid interfaceId, [MarshalAs(UnmanagedType.Interface)] out object accessible);
    [StructLayout(LayoutKind.Sequential)] private struct Point { public int X, Y; }
    [DllImport("oleacc.dll", PreserveSig = true)]
    private static extern int AccessibleObjectFromPoint(Point point,
        [MarshalAs(UnmanagedType.Interface)] out Accessibility.IAccessible accessible,
        [MarshalAs(UnmanagedType.Struct)] out object child);
}
