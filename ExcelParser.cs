using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace GUI_EXCEL_parser
{
    /// <summary>
    /// Слой чтения Excel. Здесь вся логика доступа к таблицам/колонкам.
    /// </summary>
    internal static class ExcelParser
    {
        /// <summary>
        /// Собирает список строк‑«пакетов»: шкаф + его свойства + механизмы.
        /// Формат сохранён, чтобы не ломать существующую логику разбора.
        /// </summary>
        public static List<string> GetCabinetsInfo(string filePath)
        {
            var cabinetsSheetName = "Шкафы";
            var mechSheetName = "Мех-мы";

            var mechData = ReadMechData(filePath, mechSheetName);
            var cabinetData = ReadCabinetData(filePath, cabinetsSheetName);
            var cabinetsInfo = new List<string>();

            foreach (var kks in mechData.Keys)
            {
                if (!cabinetData.ContainsKey(kks))
                    continue;

                // ВНИМАНИЕ: сохраняем строковый формат, как было (не меняем поведение).
                var line = $"{kks},{cabinetData[kks]}";

                foreach (var mech in mechData[kks])
                {
                    line += $", {mech}";
                }

                cabinetsInfo.Add(line);
            }

            return cabinetsInfo;
        }

        /// <summary>
        /// Получает алгоритмические данные для конкретного шкафа из листа «Алгоритмы».
        /// </summary>
        public static List<Dictionary<string, string>> GetAlgorithmDataByKKS(string filePath, string kksCabinet)
        {
            var result = new List<Dictionary<string, string>>();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet("Алгоритмы");

                    // Поиск колонок по заголовкам — безопаснее, чем фиксированные индексы.
                    var buildingColumn = FindColumn(worksheet, "Здание");
                    var floorColumn = FindColumn(worksheet, "Этаж (отметка)");
                    var roomColumn = FindColumn(worksheet, "Помещение");
                    var fullRoomColumn = FindColumn(worksheet, "Полное название помещения");
                    var zoneColumn = FindColumn(worksheet, "№ зоны");
                    var sourceKKSColumn = FindColumn(worksheet, "ККС шкафа источника");
                    var destKKSColumn = FindColumn(worksheet, "ККС шкафа назначения");

                    // Если чего‑то нет — возвращаем пустой список (без исключений).
                    if (!buildingColumn.HasValue || !floorColumn.HasValue || !roomColumn.HasValue ||
                        !fullRoomColumn.HasValue || !zoneColumn.HasValue || !sourceKKSColumn.HasValue ||
                        !destKKSColumn.HasValue)
                        return result;

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var destKKSValue = SafeCellString(row, destKKSColumn.Value);
                        if (!string.Equals(destKKSValue, kksCabinet, StringComparison.OrdinalIgnoreCase))
                            continue;

                        var dataRow = new Dictionary<string, string>
                        {
                            ["Здание"] = SafeCellString(row, buildingColumn.Value),
                            ["Этаж (отметка)"] = SafeCellString(row, floorColumn.Value),
                            ["Помещение"] = SafeCellString(row, roomColumn.Value),
                            ["Полное название помещения"] = SafeCellString(row, fullRoomColumn.Value),
                            ["№ зоны"] = SafeCellString(row, zoneColumn.Value),
                            ["ККС шкафа источника"] = SafeCellString(row, sourceKKSColumn.Value),
                            ["ККС шкафа назначения"] = destKKSValue
                        };

                        result.Add(dataRow);
                    }
                }
            }
            catch (Exception ex)
            {
                // Не прерываем работу приложения — пишем в консоль.
                Console.WriteLine($"Произошла ошибка при чтении файла '{filePath}': {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Читает данные о механизмах из Excel.
        /// </summary>
        public static Dictionary<string, List<string>> ReadMechData(string filePath, string sheetName)
        {
            var data = new Dictionary<string, List<string>>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(sheetName);

                var kksColumn = FindColumn(worksheet, "KKS шкафа");
                var mechColumn = FindColumn(worksheet, "KKS механизма");
                var mechNumColumn = FindColumn(worksheet, "Desigo number");
                var buildingNumColumn = FindColumn(worksheet, "Здание");

                // Если критических колонок нет — возвращаем пусто.
                if (!kksColumn.HasValue || !mechColumn.HasValue || !mechNumColumn.HasValue)
                    return data;

                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    var kksValue = SafeCellString(row, kksColumn.Value);
                    var mechValue = SafeCellString(row, mechColumn.Value);
                    var mechNum = SafeCellString(row, mechNumColumn.Value);
                    var buildingNum = buildingNumColumn.HasValue ? SafeCellString(row, buildingNumColumn.Value) : string.Empty;

                    if (string.IsNullOrWhiteSpace(kksValue) || string.IsNullOrWhiteSpace(mechValue))
                        continue;

                    if (!data.ContainsKey(kksValue))
                        data[kksValue] = new List<string>();

                    // Формат строки оставляем прежним.
                    data[kksValue].Add($"{mechValue} {mechNum} {buildingNum}".Trim());
                }
            }

            return data;
        }

        /// <summary>
        /// Читает данные по шкафам из Excel (IP/Здание/Тип/ВЕКШ).
        /// </summary>
        public static Dictionary<string, (string IP, string Building, string Type, string VEKSH)> ReadCabinetData(string filePath, string sheetName)
        {
            var data = new Dictionary<string, (string IP, string Building, string Type, string VEKSH)>();

            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(sheetName);

                    var kksColumn = FindColumn(worksheet, "KKS");
                    var ipColumn = FindColumn(worksheet, "IP S7_A1");
                    var buildingColumn = FindColumn(worksheet, "Здание");
                    var typeColumn = FindColumn(worksheet, "Тип");
                    var vekshColumn = FindColumn(worksheet, "ВЕКШ");

                    if (!kksColumn.HasValue || !ipColumn.HasValue || !buildingColumn.HasValue ||
                        !typeColumn.HasValue || !vekshColumn.HasValue)
                        return data;

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var kksValue = SafeCellString(row, kksColumn.Value);
                        var ipValue = SafeCellString(row, ipColumn.Value);
                        var buildingValue = SafeCellString(row, buildingColumn.Value).Replace(" ", "");
                        var typeValue = SafeCellString(row, typeColumn.Value);
                        var vekshValue = SafeCellString(row, vekshColumn.Value);

                        if (string.IsNullOrWhiteSpace(kksValue) || data.ContainsKey(kksValue))
                            continue;

                        // Значения по умолчанию.
                        if (string.IsNullOrWhiteSpace(ipValue))
                            ipValue = "NO IP";
                        if (string.IsNullOrWhiteSpace(buildingValue))
                            buildingValue = "UNKNOWN BUILDING";
                        if (string.IsNullOrWhiteSpace(vekshValue))
                            vekshValue = "NO VEKSH";

                        var processedType = ProcessType(typeValue, vekshValue);

                        data[kksValue] = (ipValue, buildingValue, processedType, vekshValue);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка при чтении файла '{filePath}': {ex.Message}");
            }

            return data;
        }

        /// <summary>
        /// Приватная обработка типа шкафа — логика как в исходном коде.
        /// </summary>
        private static string ProcessType(string typeValue, string vekshValue)
        {
            string vekshSuffix = "00";

            if (!string.IsNullOrWhiteSpace(typeValue) &&
                typeValue.Contains("АПТС") &&
                !string.IsNullOrWhiteSpace(vekshValue) &&
                vekshValue.StartsWith("ВЕКШ.", StringComparison.OrdinalIgnoreCase))
            {
                var parts = vekshValue.Split('-');
                if (parts.Length > 1)
                    vekshSuffix = parts[parts.Length - 1];
            }

            if (!string.IsNullOrWhiteSpace(typeValue) && typeValue.Contains("АПТС"))
                return $"АПТС-{vekshSuffix}ЗАМЕНА";

            if (!string.IsNullOrWhiteSpace(typeValue) && typeValue.StartsWith("Type", StringComparison.OrdinalIgnoreCase))
            {
                typeValue = typeValue.Replace("Type", "ТипЗАМЕНА")
                                     .Replace('.', '_')
                                     .Replace("*", "");
                return typeValue;
            }

            return "Неизвестный тип";
        }

        /// <summary>
        /// Безопасное получение номера колонки по названию.
        /// </summary>
        private static int? FindColumn(IXLWorksheet worksheet, string header)
        {
            return worksheet.FirstRowUsed()
                ?.CellsUsed()
                .FirstOrDefault(c => c.Value.ToString().Contains(header))
                ?.Address.ColumnNumber;
        }

        /// <summary>
        /// Безопасное чтение значения ячейки как строки.
        /// </summary>
        private static string SafeCellString(IXLRow row, int columnNumber)
        {
            try
            {
                return row.Cell(columnNumber).GetValue<string>().Trim();
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}
