using System;
using System.Collections.Generic;

namespace GUI_EXCEL_parser
{
    /// <summary>
    /// Класс‑резолвер для определения модели механизма по KKS‑строке.
    /// Вынесен отдельно, чтобы CSV/БД использовали единый источник логики.
    /// </summary>
    internal static class MechanismModelResolver
    {
        /// <summary>
        /// Возвращает модель механизма по его KKS.
        /// </summary>
        public static string GetObjectModel(string mechanism)
        {
            // Защита от пустых/битых значений — безопасный дефолт.
            if (string.IsNullOrWhiteSpace(mechanism))
            {
                return "AKKUYU_NO_MODEL";
            }

            // Словарь соответствия кодов оборудования и типов механизмов.
            // Комментарии подробные, чтобы проще было расширять список.
            var equipmentMapping = new Dictionary<string, Dictionary<string, string>>
            {
                // Коды оборудования "AA" и соответствующие системы (клапаны/ворота)
                {
                    "AA",
                    new Dictionary<string, string>
                    {
                        { "SAC", "AKKUYU_GATE" },  // Ворота
                        { "SGC", "AKKUYU_VALVE" }, // Клапан
                        { "SGK", "AKKUYU_VALVE" },
                        { "SGA", "AKKUYU_VALVE" },
                        { "GKE", "AKKUYU_VALVE" },
                        { "KLE", "AKKUYU_GATE" },
                        { "SAN", "AKKUYU_GATE" },
                        { "SAS", "AKKUYU_GATE" },
                        { "SAU", "AKKUYU_GATE" },
                        { "KLA", "AKKUYU_GATE" },
                        { "KLC", "AKKUYU_GATE" },
                        { "KLF", "AKKUYU_GATE" },
                        { "KLS", "AKKUYU_GATE" },
                        { "SAH", "AKKUYU_GATE" },
                        { "KLL", "AKKUYU_GATE" },
                        { "SAB", "AKKUYU_GATE" },
                        { "SAT", "AKKUYU_GATE" },
                        { "SAK", "AKKUYU_GATE" },
                        { "SAF", "AKKUYU_GATE" },
                        { "SAR", "AKKUYU_GATE" },
                        { "KLB", "AKKUYU_GATE" },
                        { "SAD", "AKKUYU_GATE" },
                        { "SAM", "AKKUYU_GATE" },
                        { "SAQ", "AKKUYU_GATE" },
                        { "SAE", "AKKUYU_GATE" }
                    }
                },
                // Коды оборудования "AN" для вентиляторов
                {
                    "AN",
                    new Dictionary<string, string>
                    {
                        { "SAC", "AKKUYU_FAN" },
                        { "SAM", "AKKUYU_FAN" },
                        { "SAF", "AKKUYU_FAN" },
                        { "SAS", "AKKUYU_FAN" },
                        { "SAN", "AKKUYU_FAN" },
                        { "SAT", "AKKUYU_FAN" },
                        { "SAK", "AKKUYU_FAN" },
                        { "KLS", "AKKUYU_FAN" },
                        { "BLF", "AKKUYU_FAN" },
                        { "SAB", "AKKUYU_FAN" },
                        { "SAH", "AKKUYU_FAN" },
                        { "SAU", "AKKUYU_FAN" },
                        { "SAQ", "AKKUYU_FAN" },
                        { "BKV", "AKKUYU_FAN" },
                        { "SGC", "AKKUYU_FAN" },
                        { "KLE", "AKKUYU_FAN" },
                        { "KLF", "AKKUYU_FAN" },
                        { "KLC", "AKKUYU_FAN" },
                        { "SAD", "AKKUYU_FAN" }
                    }
                },
                // Коды оборудования "AP" для насосов
                {
                    "AP",
                    new Dictionary<string, string>
                    {
                        { "SGA", "AKKUYU_MOTOR" },
                        { "SGK", "AKKUYU_MOTOR" }
                    }
                },
                // Заглушка под будущие группы оборудования
                {
                    "CP",
                    new Dictionary<string, string>()
                }
            };

            // Поиск соответствия: сначала код оборудования, потом код системы.
            foreach (var equipmentEntry in equipmentMapping)
            {
                if (mechanism.IndexOf(equipmentEntry.Key, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                foreach (var systemEntry in equipmentEntry.Value)
                {
                    if (mechanism.IndexOf(systemEntry.Key, StringComparison.OrdinalIgnoreCase) >= 0)
                        return systemEntry.Value;
                }
            }

            // Ничего не нашли — возвращаем дефолт.
            return "AKKUYU_NO_MODEL";
        }
    }
}
