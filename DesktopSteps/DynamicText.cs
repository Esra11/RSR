using System.Globalization;
using System.Text.RegularExpressions;

namespace DesktopSteps;

internal static class DynamicText
{
    private static readonly Regex Date = new(@"\b(?<day>\d{1,2})(?<space>\s+)(?<month>[A-Za-z]{1,9})\s+(?<year>\d{2}|\d{4})\b");
    private static readonly Regex Token = new(@"\{\{next:(?<weekday>Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday):(?<format>[^}]+)\}\}", RegexOptions.IgnoreCase);
    private static readonly Regex StaleDateBeforeToken = new(@"\(\d{1,2}\s+[A-Za-z]{1,9}\s+\d{2,4}\)\s+(?=\(\{\{next:)", RegexOptions.IgnoreCase);

    public static string ToTemplate(string text, string? relativeWeekday)
    {
        if (string.IsNullOrWhiteSpace(relativeWeekday)) return text;
        if (!Enum.TryParse<DayOfWeek>(relativeWeekday, true, out var weekday))
            throw new InvalidOperationException($"Unknown relative weekday: {relativeWeekday}.");
        if (Token.IsMatch(text)) return StaleDateBeforeToken.Replace(text, "");
        var matches = Date.Matches(text);
        if (matches.Count == 0) return $"{text.TrimEnd()} ({{{{next:{weekday}:d MMM yy}}}})";
        var match = matches[^1];
        var format = Format(match);
        return text[..match.Index] + $"{{{{next:{weekday}:{format}}}}}" + text[(match.Index + match.Length)..];
    }

    public static string Resolve(string text, string? relativeWeekday, DateTime? runDate = null)
    {
        text = StaleDateBeforeToken.Replace(text, "");
        if (Token.IsMatch(text))
            return Token.Replace(text, match => NextDate(match.Groups["weekday"].Value, runDate)
                .ToString(match.Groups["format"].Value, CultureInfo.InvariantCulture));
        if (string.IsNullOrWhiteSpace(relativeWeekday)) return text;
        var date = NextDate(relativeWeekday, runDate);
        var matches = Date.Matches(text);
        if (matches.Count == 0) return $"{text.TrimEnd()} ({date.ToString("d MMM yy", CultureInfo.InvariantCulture)})";
        var match = matches[^1];
        return text[..match.Index] + date.ToString(Format(match), CultureInfo.InvariantCulture) + text[(match.Index + match.Length)..];
    }

    private static DateTime NextDate(string relativeWeekday, DateTime? runDate)
    {
        if (!Enum.TryParse<DayOfWeek>(relativeWeekday, true, out var weekday))
            throw new InvalidOperationException($"Unknown relative weekday: {relativeWeekday}.");
        var today = (runDate ?? DateTime.Today).Date;
        var days = ((int)weekday - (int)today.DayOfWeek + 7) % 7;
        if (days == 0) days = 7; // "next" always means a future occurrence.
        return today.AddDays(days);
    }

    private static string Format(Match match) =>
        (match.Groups["day"].Value.Length == 2 ? "dd" : "d") + " " +
            (match.Groups["month"].Value.Length > 3 ? "MMMM" : "MMM") + " " +
            (match.Groups["year"].Value.Length == 4 ? "yyyy" : "yy");
}
