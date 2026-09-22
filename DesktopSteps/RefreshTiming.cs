using System.Text.RegularExpressions;

namespace DesktopSteps;

internal static class RefreshTiming
{
    public static ExecutionPlan AddObservedWait(ExecutionPlan plan, IReadOnlyList<RecordedEvent> events)
    {
        var changed = false;
        var steps = plan.Steps.Select(step =>
        {
            if (step.Action != "click" || step.Target?.ControlType != "ControlType.Button" ||
                !Regex.IsMatch(step.Target.Name ?? "", @"\b(refresh|reload|sync)\b", RegexOptions.IgnoreCase) ||
                step.Value?.StartsWith("refresh-observed-ms:", StringComparison.Ordinal) == true) return step;
            for (var i = 0; i < events.Count - 1; i++)
            {
                var current = events[i];
                if (current.Kind != "click" || current.Target?.Process != step.Target.Process ||
                    current.Target?.Name != step.Target.Name) continue;
                var next = events.Skip(i + 1).FirstOrDefault(e => e.Target?.Process == step.Target.Process &&
                    (e.Kind != "key" || !ActionGrouper.IsStandaloneModifier(e.Key)));
                if (next is null) break;
                var elapsed = (int)(next.At - current.At).TotalMilliseconds;
                if (elapsed < 500) break;
                changed = true;
                return step with { Value = $"refresh-observed-ms:{Math.Clamp(elapsed, 1000, 18000)}" };
            }
            return step;
        }).ToList();
        return changed ? plan with { Steps = steps } : plan;
    }
}
