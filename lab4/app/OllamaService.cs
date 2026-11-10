using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using ExcelDataReader;
using Microsoft.Extensions.Logging;

namespace StreetFighterML.Components.Services
{
    public sealed record ChatTurn(string Role, string Content)
    {
        public const string User = "user";
        public const string Assistant = "assistant";
    }

    public class OllamaService
    {
        private const string ModelName = "ingrid";
        private const int TipsMaxSentences = 3;
        private const int ChatMaxSentences = 4;
        private const int SmallTalkMaxSentences = 3;
        private const int MaxAttempts = 2;
        private const int MaxHistoryTurns = 8;    
        private const int MaxUserMessageLength = 500;
        private const int ContextWindow = 8192;    

        private const string MovesReminder =
            "\n\n(Answer as Ingrid in 2-4 short sentences. Name moves only if they really help answer the question, " +
            "and only moves from the MOVES OF lists. Never write button codes or notation; the app adds inputs itself.)";

        private const string NoMovesReminder =
            "\n\n(This question does not need move advice: answer as Ingrid in 1-3 short sentences and do NOT name " +
            "any move, combo or button. Never write button codes or notation.)";

        private static readonly Regex GameplayRx = new(
            @"\b(combo\w*|moves?|special\w*|super\w*|sa[123]|arts?|jump\w*|anti[- ]?air\w*|wake[- ]?up|oki|neutral|" +
            @"defen[sc]e\w*|offen[sc]e\w*|pressure|attack\w*|block\w*|parry|drive|impact|rush|reversal|throw\w*|" +
            @"tech\w*|frames?|frame[- ]?data|option\w*|fireball\w*|projectile\w*|zon(?:e|ing)|poke\w*|whiff\w*|" +
            @"punish\w*|counter\w*|beat|fight\w*|play\w*|match-?ups?|strateg\w*|tactic\w*|tips?|advice|corner|" +
            @"knockdown|safe|unsafe|cancel\w*|set-?ups?|buttons?|kicks?|punch\w*|normals?|inputs?|learn|practi[cs]e|" +
            @"train\w*|improve|win|lose|losing|stuck|bait\w*|mix-?ups?|dodge|escape|approach|spacing|range|damage)\b",
            RegexOptions.IgnoreCase | RegexOptions.Compiled);


        private static readonly SemaphoreSlim LoadLock = new(1, 1);
        private static GuideBook? _guides;
        private static MoveCatalog? _catalog;
        private static readonly ConcurrentDictionary<string, MoveMarkup> MarkupCache = new();

        private readonly HttpClient _httpClient;
        private readonly ILogger<OllamaService> _logger;

        public OllamaService(HttpClient httpClient, ILogger<OllamaService> logger)
        {
            _httpClient = httpClient;
            _logger = logger;

            if (_httpClient.BaseAddress == null)
            {
                _httpClient.BaseAddress = new Uri("http://localhost:11434/");
            }
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

        public Task<string> GetMatchupTipsAsync(string playerChar, string opponentChar, CancellationToken ct = default)
        {
            _logger.LogInformation("[OllamaService] Matchup tips requested for {Player} vs {Opponent}", playerChar, opponentChar);
            return RunCoachAsync(playerChar, opponentChar, Array.Empty<ChatTurn>(), null, ct);
        }

        public Task<string> GetChatReplyAsync(
            string playerChar, string opponentChar,
            IReadOnlyList<ChatTurn> history, string userMessage,
            CancellationToken ct = default)
        {
            userMessage = (userMessage ?? "").Trim();
            if (userMessage.Length > MaxUserMessageLength) userMessage = userMessage[..MaxUserMessageLength];

            _logger.LogInformation("[OllamaService] Chat question for {Player} vs {Opponent}: {Question}", playerChar, opponentChar, userMessage);
            return RunCoachAsync(playerChar, opponentChar, history, userMessage, ct);
        }


        private static string? CheckSpecialPhrases(string? userMessage)
        {
            if (string.IsNullOrWhiteSpace(userMessage)) return null;

            var cleaned = userMessage.Trim().TrimEnd('!', '?', '.').ToLowerInvariant();

            if (cleaned == "слава україні")
            {
                return "Героям Слава!";
            }

            if (cleaned == "glory to ukraine")
            {
                return "Glory to the heroes.";
            }

            if (cleaned == "слава нації")
            {
                return "Смерть ворогам!";
            }

            if (cleaned == "glory to the nation")
            {
                return "Death to enemies.";
            }

            if (cleaned == "україна")
            {
                return "Понад усе!";
            }

            if (cleaned == "ukraine")
            {
                return "Above all!";
            }

            if (cleaned.Contains("путін") || cleaned.Contains("хто такий путін"))
            {
                return "Хуйло.";
            }

            if (cleaned.Contains("putin"))
            {
                return "Huilo.";
            }

            return null;
        }

        private async Task<string> RunCoachAsync(
            string playerChar, string opponentChar,
            IReadOnlyList<ChatTurn> history, string? userMessage,
            CancellationToken ct)
        {
            string? specialReply = CheckSpecialPhrases(userMessage);
            if (specialReply != null)
            {
                return specialReply;
            }

            bool movesAllowed = userMessage == null || GameplayRx.IsMatch(userMessage);
            int maxSentences = userMessage == null ? TipsMaxSentences
                             : movesAllowed ? ChatMaxSentences
                             : SmallTalkMaxSentences;

            MoveCatalog catalog;
            MoveMarkup markup;
            string contextPrompt;
            try
            {
                var guides = await LoadGuidesAsync();
                catalog = await LoadCatalogAsync();
                markup = GetMarkup(catalog, playerChar, opponentChar);

                if (movesAllowed)
                {
                    string facts = catalog.BuildMoveLists(playerChar, opponentChar) +
                                   markup.ToPlainLanguage(guides.BuildNotes(playerChar, opponentChar));
                    contextPrompt = BuildPrompt(playerChar, opponentChar, facts);
                }
                else
                {
                    contextPrompt = BuildSmallTalkPrompt(playerChar, opponentChar);
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[OllamaService] Error reading Excel file.");
                return "Ingrid cannot find her scrolls right now. Try again in a moment, little one.";
            }

            if (catalog.IsEmpty)
                _logger.LogWarning("[OllamaService] Move catalog is empty: replies will have NO input markup and no move validation.");
            _logger.LogInformation("[OllamaService] Moves allowed in answer: {Allowed}", movesAllowed);

            var messages = new List<ChatTurn> { new(ChatTurn.User, contextPrompt) };

            var recent = history.Count > MaxHistoryTurns
                ? history.Take(1).Concat(history.TakeLast(MaxHistoryTurns - 1))
                : history;

            foreach (var turn in recent)
            {
                string content = turn.Role == ChatTurn.Assistant ? MoveMarkup.StripMarkup(turn.Content) : turn.Content;
                messages.Add(new ChatTurn(turn.Role, content));
            }
            if (!string.IsNullOrWhiteSpace(userMessage))
                messages.Add(new ChatTurn(ChatTurn.User, userMessage + (movesAllowed ? MovesReminder : NoMovesReminder)));

            string lastClean = "";
            string? retryNote = null;
            try
            {
                for (int attempt = 0; attempt < MaxAttempts; attempt++)
                {
                    ct.ThrowIfCancellationRequested();

                    double? temperature = attempt == 0 ? null : 0.15;
                    var outgoing = attempt == 0 || retryNote == null
                        ? messages
                        : WithNoteOnLastMessage(messages, retryNote);

                    string raw = await CallOllamaAsync(outgoing, temperature, ct);
                    string clean = ReplyCleaner.Clean(raw, maxSentences);

                    string? wrongMove = catalog.FindMoveNotOwnedBy(clean, playerChar, opponentChar);
                    bool unwantedMove = !movesAllowed && markup.ContainsMove(clean);

                    if (clean.Length > 0 && wrongMove == null && !unwantedMove)
                    {
                        string answer = markup.Annotate(clean);
                        _logger.LogInformation("[OllamaService] Attempt {Attempt} raw: {Raw} | Final: {Answer}", attempt + 1, raw, answer);
                        return answer;
                    }

                    _logger.LogWarning(
                        "[OllamaService] Rejected answer (attempt {Attempt}, wrong move: {Move}, unwanted move: {Unwanted}). Raw: {Raw}",
                        attempt + 1, wrongMove ?? "none", unwantedMove, raw);

                    lastClean = clean;
                    retryNote = wrongMove != null
                        ? $"\nImportant: do not mention \"{wrongMove}\"; it is not a move of {playerChar} or {opponentChar}.\n"
                        : unwantedMove
                            ? "\nImportant: this question needs no move advice. Do not name any move, combo or button.\n"
                            : null;
                }
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[OllamaService] Error in Ollama communication.");
                return $"Ollama connection error: {ex.Message}";
            }

            var safeSentences = ReplyCleaner.SplitSentences(lastClean)
                .Where(s => catalog.FindMoveNotOwnedBy(s, playerChar, opponentChar) == null &&
                            (movesAllowed || !markup.ContainsMove(s)));
            string safeText = string.Join(" ", safeSentences);

            return safeText.Length > 0
                ? markup.Annotate(safeText)
                : "Hmph, the answer slipped away from me. Ask me once more, dear one.";
        }

        private static List<ChatTurn> WithNoteOnLastMessage(List<ChatTurn> messages, string note)
        {
            var copy = new List<ChatTurn>(messages);
            var last = copy[^1];
            copy[^1] = last with { Content = last.Content + note };
            return copy;
        }

        private static string BuildPrompt(string playerChar, string opponentChar, string facts) =>
            $"Player character: {playerChar}\n" +
            $"Opponent: {opponentChar}\n\n" +
            "CRITICAL RULES:\n" +
            "1. Mention only moves listed in the MOVES OF lists (use their exact names) or described in the GUIDE NOTES.\n" +
            "2. Never give a move to a character whose list does not contain it.\n" +
            "3. Never write button codes, numbers, directions or notation (like 236P, 2MK, [4]6P, j.HP). " +
            "Describe normal attacks in plain words, e.g. \"crouching medium kick\" or \"heavy punch\". " +
            "The app attaches the exact inputs to every move itself.\n\n" +
            $"GUIDE NOTES:\n{facts}\n" +
            "Give your advice now.";

        private static string BuildSmallTalkPrompt(string playerChar, string opponentChar) =>
            $"You are Ingrid, a warm Street Fighter 6 coach. The player is playing {playerChar} against {opponentChar}.\n\n" +
            "RULES:\n" +
            "1. Answer in character, in 1-3 short sentences.\n" +
            "2. Do NOT name any move, combo, special, super or button, and never write button codes or notation, " +
            "unless the player explicitly asks about moves.\n" +
            "3. If the question is not about Street Fighter, answer very briefly or politely say you only coach " +
            "Street Fighter 6. Never write code or long texts.";

        private async Task<GuideBook> LoadGuidesAsync()
        {
            if (_guides != null) return _guides;

            await LoadLock.WaitAsync();
            try
            {
                if (_guides != null) return _guides;

                using var stream = await FileSystem.OpenAppPackageFileAsync("CharacterGuides.xlsx");
                using var reader = ExcelReaderFactory.CreateReader(stream);
                _guides = GuideBook.FromDataSet(reader.AsDataSet());
                return _guides;
            }
            finally
            {
                LoadLock.Release();
            }
        }

        private async Task<MoveCatalog> LoadCatalogAsync()
        {
            if (_catalog != null) return _catalog;

            await LoadLock.WaitAsync();
            try
            {
                if (_catalog != null) return _catalog;

                try
                {
                    using var stream = await FileSystem.OpenAppPackageFileAsync("sf6_moves.xlsx");
                    using var reader = ExcelReaderFactory.CreateReader(stream);
                    var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration { UseHeaderRow = true }
                    });

                    var catalog = MoveCatalog.FromTable(dataSet.Tables.Count > 0 ? dataSet.Tables[0] : null);
                    _logger.LogInformation("[OllamaService] Loaded {Count} moves from sf6_moves.xlsx.", catalog.Count);

                    if (catalog.IsEmpty) return catalog;
                    MarkupCache.Clear();
                    _catalog = catalog;
                    return catalog;
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "[OllamaService] Could not load sf6_moves.xlsx; move markup disabled for this request.");
                    return MoveCatalog.Empty;
                }
            }
            finally
            {
                LoadLock.Release();
            }
        }

        private static MoveMarkup GetMarkup(MoveCatalog catalog, string player, string opponent)
        {
            if (catalog.IsEmpty) return new MoveMarkup(catalog, player, opponent);

            string key = MoveNames.Norm(player) + "|" + MoveNames.Norm(opponent);
            return MarkupCache.GetOrAdd(key, _ => new MoveMarkup(catalog, player, opponent));
        }
        private async Task<string> CallOllamaAsync(IReadOnlyList<ChatTurn> messages, double? temperature, CancellationToken ct)
        {
            var requestBody = new
            {
                model = ModelName,
                stream = false,
                keep_alive = "30m",
                options = temperature.HasValue
                    ? (object)new { temperature = temperature.Value, num_ctx = ContextWindow }
                    : new { num_ctx = ContextWindow },
                messages = messages.Select(m => new { role = m.Role, content = m.Content }).ToArray()
            };

            using var content = new StringContent(JsonSerializer.Serialize(requestBody), Encoding.UTF8, "application/json");
            using var response = await _httpClient.PostAsync("api/chat", content, ct);
            response.EnsureSuccessStatusCode();

            var json = await response.Content.ReadAsStringAsync(ct);
            using var doc = JsonDocument.Parse(json);
            return doc.RootElement.GetProperty("message").GetProperty("content").GetString() ?? "";
        }
    }
}