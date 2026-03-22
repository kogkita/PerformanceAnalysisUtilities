using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;

namespace TestApp
{
    /// <summary>
    /// Represents a chunk of text extracted from a file, with metadata about
    /// where it came from (file name, sheet/page, line range).
    /// </summary>
    public class FileChunk
    {
        public string FileName  { get; set; } = "";
        public string Section   { get; set; } = "";   // e.g. "Sheet1", "Page 3", "Lines 100-200"
        public string Content   { get; set; } = "";
        public int    CharCount => Content.Length;
    }

    /// <summary>
    /// Represents a loaded file with its extracted text chunks ready for
    /// AI context injection.
    /// </summary>
    public class LoadedFile
    {
        public string           FileName   { get; set; } = "";
        public string           FilePath   { get; set; } = "";
        public string           FileType   { get; set; } = "";   // "xlsx", "csv", "txt", etc.
        public long             FileSize   { get; set; }
        public List<FileChunk>  Chunks     { get; set; } = new();
        public int              TotalChars => Chunks.Sum(c => c.CharCount);

        /// <summary>Short display summary, e.g. "report.xlsx (3 sheets, 12,400 chars)".</summary>
        public string Summary => $"{FileName} ({Chunks.Count} chunk(s), {TotalChars:N0} chars)";
    }

    /// <summary>
    /// Reads files of various types and produces text chunks suitable for
    /// AI context windows.  Large files are split into chunks of roughly
    /// <see cref="MaxChunkChars"/> characters each.
    /// </summary>
    public static class FileChunker
    {
        /// <summary>
        /// Approximate max characters per chunk.  ~3,000 chars ≈ ~750 tokens,
        /// leaving room for multiple chunks + conversation in a single request.
        /// </summary>
        public const int MaxChunkChars = 3000;

        /// <summary>Supported file extensions (lowercase, with dot).</summary>
        public static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
        {
            ".xlsx", ".xls", ".csv", ".tsv", ".jtl",
            ".txt", ".log", ".json", ".xml", ".md", ".yaml", ".yml",
            ".pdf", ".docx"
        };

        /// <summary>Returns true if the file extension is supported.</summary>
        public static bool IsSupported(string path)
            => SupportedExtensions.Contains(Path.GetExtension(path));

        /// <summary>
        /// Loads a file and extracts text chunks.  Throws on unsupported types.
        /// </summary>
        public static LoadedFile Load(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("File not found.", path);

            string ext = Path.GetExtension(path).ToLowerInvariant();
            var fi = new FileInfo(path);

            var loaded = new LoadedFile
            {
                FileName = Path.GetFileName(path),
                FilePath = path,
                FileType = ext.TrimStart('.'),
                FileSize = fi.Length,
            };

            loaded.Chunks = ext switch
            {
                ".xlsx" or ".xls" => ReadExcel(path),
                ".csv" or ".tsv" or ".jtl" => ReadDelimited(path, ext),
                ".txt" or ".log" or ".json" or ".xml" or ".md"
                    or ".yaml" or ".yml" => ReadTextFile(path),
                ".pdf"  => ReadPdf(path),
                ".docx" => ReadDocx(path),
                _ => throw new NotSupportedException($"Unsupported file type: {ext}")
            };

            return loaded;
        }

        // ── Excel (.xlsx) ─────────────────────────────────────────────────────

        private static List<FileChunk> ReadExcel(string path)
        {
            var chunks = new List<FileChunk>();
            string fileName = Path.GetFileName(path);

            ExcelPackage.License.SetNonCommercialPersonal("AI File Reader");

            // Copy to temp to avoid file locks (same pattern as TestRunTrends)
            string tempPath = Path.Combine(Path.GetTempPath(),
                $"AiRead_{Guid.NewGuid():N}_{Path.GetFileName(path)}");
            File.Copy(path, tempPath, overwrite: true);

            try
            {
                using var pkg = new ExcelPackage(new FileInfo(tempPath));
                foreach (var ws in pkg.Workbook.Worksheets)
                {
                    if (ws.Dimension == null) continue;

                    int rows = ws.Dimension.Rows;
                    int cols = ws.Dimension.Columns;
                    var sb = new StringBuilder();
                    int chunkStart = 1;

                    for (int r = 1; r <= rows; r++)
                    {
                        var rowParts = new List<string>();
                        for (int c = 1; c <= cols; c++)
                        {
                            string val = ws.Cells[r, c].Text?.Trim() ?? "";
                            if (!string.IsNullOrEmpty(val))
                                rowParts.Add(val);
                        }
                        if (rowParts.Count > 0)
                            sb.AppendLine(string.Join(" | ", rowParts));

                        // Flush chunk if we've exceeded the limit
                        if (sb.Length >= MaxChunkChars)
                        {
                            chunks.Add(new FileChunk
                            {
                                FileName = fileName,
                                Section  = $"{ws.Name} (rows {chunkStart}-{r})",
                                Content  = sb.ToString()
                            });
                            sb.Clear();
                            chunkStart = r + 1;
                        }
                    }

                    // Remaining rows
                    if (sb.Length > 0)
                    {
                        chunks.Add(new FileChunk
                        {
                            FileName = fileName,
                            Section  = chunkStart == 1
                                ? ws.Name
                                : $"{ws.Name} (rows {chunkStart}-{rows})",
                            Content  = sb.ToString()
                        });
                    }
                }
            }
            finally
            {
                try { File.Delete(tempPath); } catch { }
            }

            return chunks;
        }

        // ── CSV / TSV / JTL ───────────────────────────────────────────────────

        private static List<FileChunk> ReadDelimited(string path, string ext)
        {
            string fileName = Path.GetFileName(path);
            var lines = File.ReadAllLines(path);
            return ChunkLines(lines, fileName, "data");
        }

        // ── Plain text / log / json / xml / md / yaml ─────────────────────────

        private static List<FileChunk> ReadTextFile(string path)
        {
            string fileName = Path.GetFileName(path);
            var lines = File.ReadAllLines(path);
            return ChunkLines(lines, fileName, "content");
        }

        // ── PDF (text extraction) ─────────────────────────────────────────────

        private static List<FileChunk> ReadPdf(string path)
        {
            // Simple text extraction using binary scan for text streams.
            // For production, a library like PdfPig would be better, but this
            // avoids adding a NuGet dependency.  Falls back to raw byte scan.
            string fileName = Path.GetFileName(path);
            var chunks = new List<FileChunk>();

            try
            {
                // Attempt to extract text between BT/ET markers in the PDF stream
                byte[] bytes = File.ReadAllBytes(path);
                string rawText = ExtractPdfText(bytes);

                if (string.IsNullOrWhiteSpace(rawText))
                {
                    chunks.Add(new FileChunk
                    {
                        FileName = fileName,
                        Section  = "notice",
                        Content  = "[This PDF appears to be image-based or encrypted. Text extraction was not possible. Consider converting to text first.]"
                    });
                    return chunks;
                }

                var lines = rawText.Split('\n');
                return ChunkLines(lines, fileName, "page content");
            }
            catch
            {
                chunks.Add(new FileChunk
                {
                    FileName = fileName,
                    Section  = "error",
                    Content  = "[Could not extract text from this PDF.]"
                });
                return chunks;
            }
        }

        /// <summary>
        /// Basic PDF text extraction — scans for text between BT/ET operators
        /// and decodes parenthesised string literals.  Handles the majority of
        /// text-based PDFs without any external library.
        /// </summary>
        private static string ExtractPdfText(byte[] pdf)
        {
            string content = Encoding.Latin1.GetString(pdf);
            var sb = new StringBuilder();

            int pos = 0;
            while (pos < content.Length)
            {
                // Find "BT" (begin text) operator
                int bt = content.IndexOf("BT", pos, StringComparison.Ordinal);
                if (bt < 0) break;

                int et = content.IndexOf("ET", bt + 2, StringComparison.Ordinal);
                if (et < 0) break;

                string block = content[bt..et];

                // Extract parenthesised strings: Tj and TJ operators
                int i = 0;
                while (i < block.Length)
                {
                    if (block[i] == '(')
                    {
                        int depth = 1;
                        var literal = new StringBuilder();
                        i++;
                        while (i < block.Length && depth > 0)
                        {
                            if (block[i] == '\\' && i + 1 < block.Length)
                            {
                                i++; // skip escape
                                literal.Append(block[i]);
                            }
                            else if (block[i] == '(') { depth++; literal.Append(block[i]); }
                            else if (block[i] == ')') { depth--; if (depth > 0) literal.Append(block[i]); }
                            else literal.Append(block[i]);
                            i++;
                        }
                        sb.Append(literal);
                    }
                    else
                    {
                        i++;
                    }
                }
                sb.AppendLine();
                pos = et + 2;
            }

            return sb.ToString().Trim();
        }

        // ── DOCX (Word) ──────────────────────────────────────────────────────

        private static List<FileChunk> ReadDocx(string path)
        {
            string fileName = Path.GetFileName(path);
            var chunks = new List<FileChunk>();

            try
            {
                // DOCX is a ZIP — extract word/document.xml and strip XML tags
                using var zip = System.IO.Compression.ZipFile.OpenRead(path);
                var docEntry = zip.Entries.FirstOrDefault(e =>
                    e.FullName.Equals("word/document.xml", StringComparison.OrdinalIgnoreCase));

                if (docEntry == null)
                {
                    chunks.Add(new FileChunk
                    {
                        FileName = fileName,
                        Section  = "error",
                        Content  = "[Could not find document.xml in the DOCX archive.]"
                    });
                    return chunks;
                }

                using var stream = docEntry.Open();
                using var reader = new StreamReader(stream, Encoding.UTF8);
                string xml = reader.ReadToEnd();

                // Simple XML → text: replace paragraph breaks with newlines,
                // strip all tags, decode entities
                string text = xml
                    .Replace("</w:p>", "\n")
                    .Replace("</w:tr>", "\n");

                // Strip XML tags
                var sb = new StringBuilder();
                bool inTag = false;
                foreach (char c in text)
                {
                    if (c == '<') { inTag = true; continue; }
                    if (c == '>') { inTag = false; continue; }
                    if (!inTag) sb.Append(c);
                }

                text = System.Net.WebUtility.HtmlDecode(sb.ToString());
                var lines = text.Split('\n')
                    .Select(l => l.Trim())
                    .Where(l => !string.IsNullOrEmpty(l))
                    .ToArray();

                return ChunkLines(lines, fileName, "content");
            }
            catch
            {
                chunks.Add(new FileChunk
                {
                    FileName = fileName,
                    Section  = "error",
                    Content  = "[Could not extract text from this DOCX file.]"
                });
                return chunks;
            }
        }

        // ── Shared line chunker ───────────────────────────────────────────────

        private static List<FileChunk> ChunkLines(
            string[] lines, string fileName, string sectionPrefix)
        {
            var chunks = new List<FileChunk>();
            var sb = new StringBuilder();
            int chunkStart = 1;

            for (int i = 0; i < lines.Length; i++)
            {
                sb.AppendLine(lines[i]);

                if (sb.Length >= MaxChunkChars)
                {
                    chunks.Add(new FileChunk
                    {
                        FileName = fileName,
                        Section  = $"{sectionPrefix} (lines {chunkStart}-{i + 1})",
                        Content  = sb.ToString()
                    });
                    sb.Clear();
                    chunkStart = i + 2;
                }
            }

            if (sb.Length > 0)
            {
                chunks.Add(new FileChunk
                {
                    FileName = fileName,
                    Section  = chunkStart == 1
                        ? sectionPrefix
                        : $"{sectionPrefix} (lines {chunkStart}-{lines.Length})",
                    Content  = sb.ToString()
                });
            }

            return chunks;
        }
    }
}
