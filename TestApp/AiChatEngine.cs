using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace TestApp
{
    /// <summary>
    /// Manages a multi-file AI chat session.  Handles:
    ///   • Loading and tracking files with their chunks
    ///   • Building the system prompt with file context
    ///   • Maintaining conversation history
    ///   • Dispatching to the selected AI provider with streaming
    ///
    /// Context strategy: includes as many file chunks as possible within
    /// a budget (~100K chars ≈ ~25K tokens), prioritising the most recently
    /// added files and chunks that are most relevant based on simple keyword
    /// overlap with the user's question.
    /// </summary>
    public class AiChatEngine
    {
        // ── Configuration ─────────────────────────────────────────────────────

        /// <summary>Max characters of file context to include in the system prompt.</summary>
        public int MaxContextChars { get; set; } = 100_000;

        /// <summary>Max conversation history messages to include (user + assistant pairs).</summary>
        public int MaxHistoryMessages { get; set; } = 20;

        // ── State ─────────────────────────────────────────────────────────────

        private readonly List<LoadedFile> _files   = new();
        private readonly List<AiMessage>  _history = new();

        public IReadOnlyList<LoadedFile> Files   => _files;
        public IReadOnlyList<AiMessage>  History => _history;

        public int TotalFileChars => _files.Sum(f => f.TotalChars);
        public int TotalChunks   => _files.Sum(f => f.Chunks.Count);

        // ── File management ───────────────────────────────────────────────────

        /// <summary>Loads a file and adds it to the context pool.</summary>
        public LoadedFile AddFile(string path)
        {
            // Don't add the same file twice
            var existing = _files.FirstOrDefault(f =>
                f.FilePath.Equals(path, StringComparison.OrdinalIgnoreCase));
            if (existing != null) return existing;

            var loaded = FileChunker.Load(path);
            _files.Add(loaded);
            return loaded;
        }

        /// <summary>Removes a file from the context pool.</summary>
        public void RemoveFile(string path)
        {
            _files.RemoveAll(f =>
                f.FilePath.Equals(path, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>Removes all files.</summary>
        public void ClearFiles() => _files.Clear();

        /// <summary>Clears conversation history but keeps files.</summary>
        public void ClearHistory() => _history.Clear();

        /// <summary>Clears everything — files and history.</summary>
        public void Reset()
        {
            _files.Clear();
            _history.Clear();
        }

        // ── Chat ──────────────────────────────────────────────────────────────

        /// <summary>
        /// Sends a user message to the AI provider with file context.
        /// Streams tokens via <paramref name="onToken"/>.
        /// Returns the full assistant response.
        /// </summary>
        public async Task<string> SendAsync(
            IAiProvider     provider,
            string          userMessage,
            Action<string>  onToken,
            CancellationToken cancel = default)
        {
            // Add user message to history
            _history.Add(new AiMessage { Role = "user", Content = userMessage });

            // Build system prompt with file context
            string systemPrompt = BuildSystemPrompt(userMessage);

            // Trim history to fit
            var messages = GetTrimmedHistory();

            // Call provider
            string response = await provider.SendAsync(
                systemPrompt, messages, onToken, cancel);

            // Add assistant response to history
            _history.Add(new AiMessage { Role = "assistant", Content = response });

            return response;
        }

        // ── System prompt builder ─────────────────────────────────────────────

        private string BuildSystemPrompt(string userQuery)
        {
            var sb = new StringBuilder();

            sb.AppendLine("You are a helpful AI assistant integrated into a Performance Test Utilities desktop application.");
            sb.AppendLine("The user has loaded one or more files for analysis. Answer their questions based on the file contents provided below.");
            sb.AppendLine("When referencing data, mention the file name and section (e.g. sheet name, line range) so the user can locate it.");
            sb.AppendLine("If the file contents are truncated or partial, say so and suggest what additional context might help.");
            sb.AppendLine();

            if (_files.Count == 0)
            {
                sb.AppendLine("No files are currently loaded. The user may ask general questions.");
                return sb.ToString();
            }

            // Gather all chunks and prioritise by relevance to the query
            var allChunks = _files
                .SelectMany(f => f.Chunks)
                .ToList();

            var selected = SelectChunks(allChunks, userQuery, MaxContextChars);

            sb.AppendLine($"=== LOADED FILES ({_files.Count} file(s), showing {selected.Count} chunk(s)) ===");
            sb.AppendLine();

            // File index
            for (int i = 0; i < _files.Count; i++)
            {
                var f = _files[i];
                sb.AppendLine($"  [{i + 1}] {f.FileName} ({f.FileType}, {f.TotalChars:N0} chars, {f.Chunks.Count} chunk(s))");
            }
            sb.AppendLine();

            // Include selected chunks
            foreach (var chunk in selected)
            {
                sb.AppendLine($"--- {chunk.FileName} / {chunk.Section} ---");
                sb.AppendLine(chunk.Content);
                sb.AppendLine();
            }

            int includedChars = selected.Sum(c => c.CharCount);
            int totalChars = allChunks.Sum(c => c.CharCount);
            if (includedChars < totalChars)
            {
                sb.AppendLine($"[Context truncated: showing {includedChars:N0} of {totalChars:N0} total characters. Ask about specific sections for more detail.]");
            }

            return sb.ToString();
        }

        /// <summary>
        /// Selects chunks to fit within the character budget, prioritising
        /// chunks whose content overlaps with keywords in the user's query.
        /// Falls back to including chunks in file order (most recent files first).
        /// </summary>
        private static List<FileChunk> SelectChunks(
            List<FileChunk> allChunks, string query, int budget)
        {
            if (allChunks.Sum(c => c.CharCount) <= budget)
                return allChunks;  // everything fits

            // Score each chunk by keyword overlap with the query
            var queryWords = query.ToLowerInvariant()
                .Split(new[] { ' ', ',', '.', '?', '!', ':', ';', '\t', '\n' },
                    StringSplitOptions.RemoveEmptyEntries)
                .Where(w => w.Length > 2)
                .ToHashSet();

            var scored = allChunks.Select((chunk, idx) =>
            {
                int score = 0;
                string lower = chunk.Content.ToLowerInvariant();
                foreach (var word in queryWords)
                    if (lower.Contains(word)) score++;

                // Boost more recent files (higher index in _files = added later)
                return (chunk, score, order: idx);
            })
            .OrderByDescending(x => x.score)
            .ThenBy(x => x.order)
            .ToList();

            var selected = new List<FileChunk>();
            int used = 0;
            foreach (var (chunk, _, _) in scored)
            {
                if (used + chunk.CharCount > budget) continue;
                selected.Add(chunk);
                used += chunk.CharCount;
            }

            // Return in original order for readability
            return selected
                .OrderBy(c => allChunks.IndexOf(c))
                .ToList();
        }

        // ── History trimming ──────────────────────────────────────────────────

        private List<AiMessage> GetTrimmedHistory()
        {
            if (_history.Count <= MaxHistoryMessages)
                return new List<AiMessage>(_history);

            // Keep the most recent messages, always including the latest user message
            return _history
                .Skip(_history.Count - MaxHistoryMessages)
                .ToList();
        }
    }
}
