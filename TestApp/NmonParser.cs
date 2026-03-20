using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace TestApp
{
    // ── Data model ────────────────────────────────────────────────────────────

    /// <summary>Metadata from AAA lines.</summary>
    public class NmonMetadata
    {
        public string Host      { get; set; } = "";
        public string Date      { get; set; } = "";
        public string Time      { get; set; } = "";
        public string OS        { get; set; } = "";
        public string Version   { get; set; } = "";
        public int    Snapshots { get; set; }
        public int    Interval  { get; set; }  // seconds
        public string FileName  { get; set; } = "";
        public List<string> BbbLines { get; set; } = new();
    }

    /// <summary>
    /// One parsed section (e.g. CPU_ALL, MEM, DISKBUSY).
    /// Columns[0] is always "Timestamp".
    /// Rows are ordered by interval index.
    /// </summary>
    public class NmonSection
    {
        public string         Tag     { get; set; } = "";
        public List<string>   Columns { get; set; } = new();
        // rows[i][0] = DateTime string, rows[i][j] = value string
        public List<string[]> Rows    { get; set; } = new();
    }

    /// <summary>Full parsed content of one .nmon file.</summary>
    public class NmonFile
    {
        public NmonMetadata                     Meta     { get; set; } = new();
        public Dictionary<string, NmonSection>  Sections { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    }

    // ── Parser ────────────────────────────────────────────────────────────────

    public static class NmonParser
    {
        public static NmonFile Parse(string path)
        {
            var file = new NmonFile();
            file.Meta.FileName = Path.GetFileName(path);

            // timestamp map: "T0001" → DateTime
            var timestamps = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

            // section headers: tag → string[] of column names (excluding tag and hostname)
            var headers = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

            // raw data rows before timestamps are resolved: tag → list of raw fields
            var rawRows = new Dictionary<string, List<string[]>>(StringComparer.OrdinalIgnoreCase);

            var lines = File.ReadAllLines(path);

            foreach (var rawLine in lines)
            {
                if (string.IsNullOrWhiteSpace(rawLine)) continue;

                var fields = SplitCsvLine(rawLine);
                if (fields.Length < 2) continue;

                string tag = fields[0].Trim();

                // ── AAA — metadata ────────────────────────────────────────
                if (tag == "AAA")
                {
                    if (fields.Length >= 3)
                    {
                        string key = fields[1].Trim().ToLowerInvariant();
                        string val = fields[2].Trim();
                        switch (key)
                        {
                            case "host":       file.Meta.Host      = val; break;
                            case "date":       file.Meta.Date      = val; break;
                            case "time":       file.Meta.Time      = val; break;
                            case "os":         file.Meta.OS        = val; break;
                            case "version":    file.Meta.Version   = val; break;
                            case "progname":   file.Meta.Version   = val; break;
                            case "snapshots":  int.TryParse(val, out int snaps); file.Meta.Snapshots = snaps; break;
                            case "interval":   int.TryParse(val, out int intv);  file.Meta.Interval  = intv; break;
                        }
                    }
                    continue;
                }

                // ── BBB — system config ───────────────────────────────────
                if (tag.StartsWith("BBB", StringComparison.OrdinalIgnoreCase))
                {
                    file.Meta.BbbLines.Add(rawLine);
                    continue;
                }

                // ── ZZZZ — timestamps ─────────────────────────────────────
                if (tag == "ZZZZ")
                {
                    // ZZZZ,T0001,10:23:45,14-Mar-2024
                    if (fields.Length >= 4)
                    {
                        string tKey  = fields[1].Trim();
                        string tTime = fields[2].Trim();
                        string tDate = fields[3].Trim();
                        if (TryParseNmonDateTime(tDate, tTime, out DateTime dt))
                            timestamps[tKey] = dt;
                    }
                    continue;
                }

                // ── Section header: second field doesn't start with T + digits ──
                string f1 = fields[1].Trim();
                bool isDataRow = f1.Length > 1
                    && (f1[0] == 'T' || f1[0] == 't')
                    && char.IsDigit(f1[1]);

                if (!isDataRow)
                {
                    // This is a header line: TAG,hostname,col1,col2,...
                    // Some sections (like TOP) have no separate header
                    var cols = new List<string> { "Timestamp" };
                    for (int i = 2; i < fields.Length; i++)
                    {
                        string c = fields[i].Trim();
                        if (!string.IsNullOrEmpty(c)) cols.Add(c);
                    }
                    headers[tag] = cols.ToArray();
                }
                else
                {
                    // Data row: TAG,T0001,val1,val2,...
                    if (!rawRows.ContainsKey(tag))
                        rawRows[tag] = new List<string[]>();
                    rawRows[tag].Add(fields);
                }
            }

            // ── Build sections ────────────────────────────────────────────────
            foreach (var kvp in rawRows)
            {
                string tag = kvp.Key;
                var section = new NmonSection { Tag = tag };

                // Resolve columns
                if (headers.TryGetValue(tag, out string[]? cols))
                    section.Columns.AddRange(cols);
                else
                    section.Columns.Add("Timestamp"); // fallback

                // Sort rows by T-index
                var sorted = kvp.Value
                    .OrderBy(r => r.Length > 1 ? r[1].Trim() : "")
                    .ToList();

                foreach (var fields in sorted)
                {
                    if (fields.Length < 2) continue;
                    string tKey = fields[1].Trim();

                    // Only include intervals within requested range
                    string tsStr = timestamps.TryGetValue(tKey, out DateTime dt)
                        ? dt.ToString("yyyy-MM-dd HH:mm:ss")
                        : tKey;

                    var row = new List<string> { tsStr };
                    // Data values start at index 2
                    for (int i = 2; i < fields.Length; i++)
                        row.Add(fields[i].Trim());

                    // Pad/trim to match column count
                    while (row.Count < section.Columns.Count) row.Add("");
                    section.Rows.Add(row.Take(section.Columns.Count).ToArray());
                }

                file.Sections[tag] = section;
            }

            return file;
        }

        // ── Helpers ───────────────────────────────────────────────────────────

        private static bool TryParseNmonDateTime(string date, string time, out DateTime result)
        {
            // date formats: "14-Mar-2024", "14-MAR-2024", "2024-03-14"
            // time formats: "10:23:45", "10:23:45.123"
            result = default;
            try
            {
                string combined = $"{date} {time}";
                string[] formats =
                {
                    "dd-MMM-yyyy HH:mm:ss",
                    "dd-MMM-yyyy HH:mm:ss.fff",
                    "dd-MMM-yy HH:mm:ss",
                    "yyyy-MM-dd HH:mm:ss",
                    "MM/dd/yyyy HH:mm:ss",
                };
                foreach (var fmt in formats)
                {
                    if (DateTime.TryParseExact(combined, fmt,
                        CultureInfo.InvariantCulture, DateTimeStyles.None, out result))
                        return true;
                }
                return false;
            }
            catch { return false; }
        }

        private static string[] SplitCsvLine(string line)
        {
            var fields = new List<string>();
            var sb = new System.Text.StringBuilder();
            bool inQ = false;
            foreach (char c in line)
            {
                if (inQ)
                {
                    if (c == '"') inQ = false;
                    else sb.Append(c);
                }
                else
                {
                    if (c == '"') inQ = true;
                    else if (c == ',') { fields.Add(sb.ToString()); sb.Clear(); }
                    else sb.Append(c);
                }
            }
            fields.Add(sb.ToString());
            return fields.ToArray();
        }
    }
}
