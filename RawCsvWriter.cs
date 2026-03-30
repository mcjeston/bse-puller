using System.Text;

namespace BsePuller;

internal static class RawCsvWriter
{
    public static void Write(string filePath, IReadOnlyList<Dictionary<string, string?>> rows)
    {
        var headers = CollectHeadersForExport(rows);
        Write(filePath, headers, rows);
    }

    public static void Write(string filePath, IReadOnlyList<string> headers, IReadOnlyList<Dictionary<string, string?>> rows)
    {
        var builder = new StringBuilder();

        builder.AppendLine(string.Join(",", headers.Select(EscapeCsv)));

        foreach (var row in rows)
        {
            var line = headers.Select(header =>
            {
                row.TryGetValue(header, out var value);
                return EscapeCsv(value ?? string.Empty);
            });

            builder.AppendLine(string.Join(",", line));
        }

        File.WriteAllText(filePath, builder.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
    }

    public static List<string> CollectHeadersForExport(IReadOnlyList<Dictionary<string, string?>> rows)
    {
        var headers = new List<string>();
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var row in rows)
        {
            foreach (var key in row.Keys)
            {
                if (seen.Add(key))
                {
                    headers.Add(key);
                }
            }
        }

        if (headers.Count == 0)
        {
            headers.Add("No data returned");
        }

        return headers;
    }

    private static string EscapeCsv(string value)
    {
        var needsQuotes = value.Contains(',') || value.Contains('"') || value.Contains('\r') || value.Contains('\n');
        if (!needsQuotes)
        {
            return value;
        }

        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }
}
