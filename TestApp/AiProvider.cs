using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace TestApp
{
    // ── Provider enum ─────────────────────────────────────────────────────────

    public enum AiProviderType { Claude, ChatGPT, Gemini }

    // ── Message model ─────────────────────────────────────────────────────────

    public class AiMessage
    {
        public string Role    { get; set; } = "user";   // "user" | "assistant" | "system"
        public string Content { get; set; } = "";
    }

    // ── Abstract provider ─────────────────────────────────────────────────────

    /// <summary>
    /// Sends a conversation to an AI provider and streams the response
    /// token-by-token via <paramref name="onToken"/>.
    /// </summary>
    public interface IAiProvider
    {
        AiProviderType Type { get; }
        string DisplayName { get; }

        /// <summary>
        /// Sends messages and calls <paramref name="onToken"/> as each chunk arrives.
        /// Returns the full concatenated response when done.
        /// </summary>
        Task<string> SendAsync(
            string       systemPrompt,
            List<AiMessage> messages,
            Action<string>  onToken,
            CancellationToken cancel = default);
    }

    // ── Claude (Anthropic Messages API) ───────────────────────────────────────

    public class ClaudeProvider : IAiProvider
    {
        public AiProviderType Type => AiProviderType.Claude;
        public string DisplayName => "Claude (Anthropic)";

        private readonly string _apiKey;
        private readonly string _model;
        private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };

        public ClaudeProvider(string apiKey, string model = "claude-sonnet-4-20250514")
        {
            _apiKey = apiKey;
            _model  = model;
        }

        public async Task<string> SendAsync(
            string systemPrompt, List<AiMessage> messages,
            Action<string> onToken, CancellationToken cancel)
        {
            var msgArray = new List<object>();
            foreach (var m in messages)
                msgArray.Add(new { role = m.Role, content = m.Content });

            var body = new
            {
                model      = _model,
                max_tokens = 4096,
                system     = systemPrompt,
                messages   = msgArray,
                stream     = true
            };

            var request = new HttpRequestMessage(HttpMethod.Post,
                "https://api.anthropic.com/v1/messages");
            request.Headers.Add("x-api-key", _apiKey);
            request.Headers.Add("anthropic-version", "2023-06-01");
            request.Content = new StringContent(
                JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

            using var response = await Http.SendAsync(request,
                HttpCompletionOption.ResponseHeadersRead, cancel);

            if (!response.IsSuccessStatusCode)
            {
                string errBody = await response.Content.ReadAsStringAsync(cancel);
                throw new HttpRequestException(
                    $"Claude API error {(int)response.StatusCode}: {errBody}");
            }

            var sb = new StringBuilder();
            using var stream = await response.Content.ReadAsStreamAsync(cancel);
            using var reader = new StreamReader(stream);

            while (!reader.EndOfStream)
            {
                cancel.ThrowIfCancellationRequested();
                string? line = await reader.ReadLineAsync();
                if (line == null) continue;
                if (!line.StartsWith("data: ")) continue;
                string json = line["data: ".Length..];
                if (json == "[DONE]") break;

                try
                {
                    using var doc = JsonDocument.Parse(json);
                    var root = doc.RootElement;

                    // content_block_delta → delta.text
                    if (root.TryGetProperty("type", out var t) &&
                        t.GetString() == "content_block_delta" &&
                        root.TryGetProperty("delta", out var delta) &&
                        delta.TryGetProperty("text", out var text))
                    {
                        string token = text.GetString() ?? "";
                        sb.Append(token);
                        onToken(token);
                    }
                }
                catch { /* skip unparseable SSE lines */ }
            }
            return sb.ToString();
        }
    }

    // ── ChatGPT (OpenAI Chat Completions API) ─────────────────────────────────

    public class ChatGptProvider : IAiProvider
    {
        public AiProviderType Type => AiProviderType.ChatGPT;
        public string DisplayName => "ChatGPT (OpenAI)";

        private readonly string _apiKey;
        private readonly string _model;
        private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };

        public ChatGptProvider(string apiKey, string model = "gpt-4o")
        {
            _apiKey = apiKey;
            _model  = model;
        }

        public async Task<string> SendAsync(
            string systemPrompt, List<AiMessage> messages,
            Action<string> onToken, CancellationToken cancel)
        {
            var msgArray = new List<object>
            {
                new { role = "system", content = systemPrompt }
            };
            foreach (var m in messages)
                msgArray.Add(new { role = m.Role, content = m.Content });

            var body = new
            {
                model    = _model,
                messages = msgArray,
                stream   = true
            };

            var request = new HttpRequestMessage(HttpMethod.Post,
                "https://api.openai.com/v1/chat/completions");
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            request.Content = new StringContent(
                JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

            using var response = await Http.SendAsync(request,
                HttpCompletionOption.ResponseHeadersRead, cancel);

            if (!response.IsSuccessStatusCode)
            {
                string errBody = await response.Content.ReadAsStringAsync(cancel);
                throw new HttpRequestException(
                    $"OpenAI API error {(int)response.StatusCode}: {errBody}");
            }

            var sb = new StringBuilder();
            using var stream = await response.Content.ReadAsStreamAsync(cancel);
            using var reader = new StreamReader(stream);

            while (!reader.EndOfStream)
            {
                cancel.ThrowIfCancellationRequested();
                string? line = await reader.ReadLineAsync();
                if (line == null) continue;
                if (!line.StartsWith("data: ")) continue;
                string json = line["data: ".Length..];
                if (json == "[DONE]") break;

                try
                {
                    using var doc = JsonDocument.Parse(json);
                    var choices = doc.RootElement.GetProperty("choices");
                    foreach (var choice in choices.EnumerateArray())
                    {
                        if (choice.TryGetProperty("delta", out var delta) &&
                            delta.TryGetProperty("content", out var content))
                        {
                            string token = content.GetString() ?? "";
                            sb.Append(token);
                            onToken(token);
                        }
                    }
                }
                catch { }
            }
            return sb.ToString();
        }
    }

    // ── Gemini (Google Generative AI API) ──────────────────────────────────────

    public class GeminiProvider : IAiProvider
    {
        public AiProviderType Type => AiProviderType.Gemini;
        public string DisplayName => "Gemini (Google)";

        private readonly string _apiKey;
        private readonly string _model;
        private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };

        public GeminiProvider(string apiKey, string model = "gemini-3-flash")
        {
            _apiKey = apiKey;
            _model  = model;
        }

        public async Task<string> SendAsync(
            string systemPrompt, List<AiMessage> messages,
            Action<string> onToken, CancellationToken cancel)
        {
            // Gemini uses a different message format: contents[].parts[].text
            var contents = new List<object>();

            // System instruction is separate in Gemini
            foreach (var m in messages)
            {
                string role = m.Role == "assistant" ? "model" : "user";
                contents.Add(new
                {
                    role  = role,
                    parts = new[] { new { text = m.Content } }
                });
            }

            var body = new
            {
                system_instruction = new
                {
                    parts = new[] { new { text = systemPrompt } }
                },
                contents = contents,
                generationConfig = new { maxOutputTokens = 4096 }
            };

            string url = $"https://generativelanguage.googleapis.com/v1beta/models/{_model}:streamGenerateContent?alt=sse&key={_apiKey}";

            var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Content = new StringContent(
                JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");

            using var response = await Http.SendAsync(request,
                HttpCompletionOption.ResponseHeadersRead, cancel);

            if (!response.IsSuccessStatusCode)
            {
                string errBody = await response.Content.ReadAsStringAsync(cancel);
                throw new HttpRequestException(
                    $"Gemini API error {(int)response.StatusCode}: {errBody}");
            }

            var sb = new StringBuilder();
            using var stream = await response.Content.ReadAsStreamAsync(cancel);
            using var reader = new StreamReader(stream);

            while (!reader.EndOfStream)
            {
                cancel.ThrowIfCancellationRequested();
                string? line = await reader.ReadLineAsync();
                if (line == null) continue;
                if (!line.StartsWith("data: ")) continue;
                string json = line["data: ".Length..];

                try
                {
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("candidates", out var candidates))
                    {
                        foreach (var cand in candidates.EnumerateArray())
                        {
                            if (cand.TryGetProperty("content", out var content) &&
                                content.TryGetProperty("parts", out var parts))
                            {
                                foreach (var part in parts.EnumerateArray())
                                {
                                    if (part.TryGetProperty("text", out var text))
                                    {
                                        string token = text.GetString() ?? "";
                                        sb.Append(token);
                                        onToken(token);
                                    }
                                }
                            }
                        }
                    }
                }
                catch { }
            }
            return sb.ToString();
        }
    }

    // ── Factory ───────────────────────────────────────────────────────────────

    public static class AiProviderFactory
    {
        public static IAiProvider Create(AiProviderType type, string apiKey)
        {
            return type switch
            {
                AiProviderType.Claude  => new ClaudeProvider(apiKey),
                AiProviderType.ChatGPT => new ChatGptProvider(apiKey),
                AiProviderType.Gemini  => new GeminiProvider(apiKey),
                _ => throw new ArgumentException($"Unknown provider: {type}")
            };
        }
    }
}
