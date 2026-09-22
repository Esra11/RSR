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
        return CommandFromObservations(Probe(window), x, y);
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
        int role, List<(Accessibility.IAccessible Parent, object Child)> matches, ref int visited)
    {
        if (depth > 8 || visited++ > 500 || matches.Count > 1) return;
        Check(accessible, 0, name, role, matches);
        int count;
        try { count = Math.Min(accessible.accChildCount, 500); }
        catch { return; }
        for (var id = 1; id <= count && visited <= 500 && matches.Count <= 1; id++)
        {
            object? child = null;
            try { child = accessible.get_accChild(id); }
            catch { }
            if (child is Accessibility.IAccessible nested)
                Visit(nested, depth + 1, name, role, matches, ref visited);
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
