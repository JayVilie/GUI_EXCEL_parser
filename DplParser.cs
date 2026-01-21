using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace GUI_EXCEL_parser
{
    /// <summary>
    /// Слой работы с DPL: поиск строк, извлечение точки, комментария и шкафа.
    /// </summary>
    internal static class DplParser
    {
        /// <summary>
        /// Ищет строки с заданными паттернами в LANG:10027.
        /// </summary>
        public static List<string> GetMatchingLinesFromDplFile(string filePath)
        {
            // Список паттернов оставляем как был, но делаем токенизацию безопасной.
            var patterns = new List<string>
            {
                "Шлюз_ Пожар ППКиУП",
                "Шлюз_ Пред_ тревога ППКиУП",
                "Шлюз_ Неиправность ППКиУП"
            };

            // Предварительная токенизация шаблонов.
            var patternTokens = patterns.Select(SplitToTokens).ToList();
            var matchingLines = new List<string>();

            foreach (string line in File.ReadLines(filePath))
            {
                var tokens = SplitLine(line);
                int langIndex = tokens.IndexOf("LANG:10027");

                if (langIndex == -1 || langIndex + 1 >= tokens.Count)
                    continue;

                string comment = tokens[langIndex + 1].Trim('"');
                var commentTokens = SplitToTokens(comment);

                // Сопоставление по токенам — устойчиво к разным пробелам/подчёркиваниям.
                if (patternTokens.Any(pt => StartsWithTokens(commentTokens, pt)))
                {
                    matchingLines.Add(line);
                }
            }

            return matchingLines;
        }

        /// <summary>
        /// Извлекает datapoint, comment и cabinetCode из строки DPL.
        /// </summary>
        public static (string dataPoint, string comment, string cabinetCode) ExtractData(string line)
        {
            string dataPoint = string.Empty;
            string comment = string.Empty;
            string cabinetCode = string.Empty;

            try
            {
                var tokens = SplitLine(line);
                dataPoint = tokens.FirstOrDefault(t => t.StartsWith("GmsDevice"));

                for (int i = 0; i < tokens.Count; i++)
                {
                    if (tokens[i] != "LANG:10027" || i + 1 >= tokens.Count)
                        continue;

                    comment = tokens[i + 1].Trim('"');
                    var commentParts = comment.Split(new[] { ' ', '_' }, StringSplitOptions.RemoveEmptyEntries);
                    cabinetCode = commentParts.LastOrDefault();

                    // Если код некорректный — помечаем как «Неизвестно».
                    if (!string.IsNullOrEmpty(cabinetCode) && IsCabinetCode(cabinetCode))
                    {
                        int index = comment.LastIndexOf(cabinetCode, StringComparison.Ordinal);
                        if (index > 0)
                            comment = comment.Substring(0, index).Trim(new[] { ' ', '_' });
                    }
                    else
                    {
                        cabinetCode = "Неизвестно";
                    }

                    break;
                }
            }
            catch
            {
                // При любой ошибке возвращаем пустые строки, чтобы не ломать вставку.
                dataPoint = string.Empty;
                comment = string.Empty;
                cabinetCode = string.Empty;
            }

            return (dataPoint, comment, cabinetCode);
        }

        /// <summary>
        /// Разбиение строки на токены с учётом кавычек.
        /// </summary>
        public static List<string> SplitLine(string line)
        {
            var tokens = new List<string>();
            bool inQuotes = false;
            string currentToken = string.Empty;

            foreach (char c in line)
            {
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    currentToken += c;
                }
                else if (char.IsWhiteSpace(c) && !inQuotes)
                {
                    if (!string.IsNullOrEmpty(currentToken))
                    {
                        tokens.Add(currentToken);
                        currentToken = string.Empty;
                    }
                }
                else
                {
                    currentToken += c;
                }
            }

            if (!string.IsNullOrEmpty(currentToken))
                tokens.Add(currentToken);

            return tokens;
        }

        /// <summary>
        /// Проверка формата KKS шкафа.
        /// </summary>
        public static bool IsCabinetCode(string code)
        {
            return !string.IsNullOrWhiteSpace(code) &&
                   code.Length == 7 &&
                   code.All(c => char.IsUpper(c) || char.IsDigit(c));
        }

        /// <summary>
        /// Нормализованное разбиение строки на токены.
        /// </summary>
        private static string[] SplitToTokens(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return Array.Empty<string>();

            return text
                .Split(new[] { ' ', '_' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => t.Length > 0)
                .ToArray();
        }

        /// <summary>
        /// Проверяет, что commentTokens начинается с patternTokens.
        /// </summary>
        private static bool StartsWithTokens(string[] commentTokens, string[] patternTokens)
        {
            if (commentTokens == null || patternTokens == null)
                return false;
            if (commentTokens.Length < patternTokens.Length)
                return false;

            for (int i = 0; i < patternTokens.Length; i++)
            {
                if (!string.Equals(commentTokens[i], patternTokens[i], StringComparison.OrdinalIgnoreCase))
                    return false;
            }

            return true;
        }
    }
}
