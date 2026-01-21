using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace GUI_EXCEL_parser
{
    /// <summary>
    /// Слой работы с SQLite: схемы, вставки, проверки уникальности.
    /// </summary>
    internal static class DatabaseWriter
    {
        /// <summary>
        /// Создаёт БД и заполняет её данными по всем шкафам.
        /// Схема создаётся один раз на файл.
        /// </summary>
        public static async Task CreateAndFillDatabaseAsync(string dbPath, List<string> cabinetsInfo)
        {
            using (var connection = new SQLiteConnection($"Data Source={dbPath}; Version=3;"))
            {
                await connection.OpenAsync();

                // 1) Схема БД — только один раз.
                await EnsureMainSchemaAsync(connection);

                // 2) Заполняем таблицы.
                var existingIdx = new HashSet<string>();

                foreach (var cabinetInfo in cabinetsInfo)
                {
                    if (!TryParseCabinetInfo(cabinetInfo, out var cabinetName, out var _, out var kkss))
                        continue;

                    var binaryEntries = new List<(string idx, string tag, string comment)>();
                    var integerEntries = new List<(string idx, string tag, string comment)>();

                    foreach (var raw in kkss)
                    {
                        if (!TryParseMechEntry(raw, out var mechKks, out var kksNumber, out _))
                            continue;

                        string model = MechanismModelResolver.GetObjectModel(mechKks);
                        if (model == "AKKUYU_NO_MODEL")
                            continue;

                        var tags = BuildDefaultTags();

                        foreach (var tag in tags)
                        {
                            string baseIdx = $"{cabinetName}_{mechKks}_{tag.Item1}";
                            string idx = baseIdx;
                            string wccTag = $"System00_1:ManagementView_FieldNetworks_S7Plus_{cabinetName}_{kksNumber}.{tag.Item1}";
                            string comment = $"{tag.Item1} {cabinetName}_{mechKks}";

                            int counter = 1;
                            while (existingIdx.Contains(idx) || await IdxExistsAsync(connection, idx))
                            {
                                idx = $"{baseIdx}_{counter}";
                                counter++;
                            }

                            existingIdx.Add(idx);

                            if (tag.Item2 == "bool")
                                binaryEntries.Add((idx, wccTag, comment));
                            else
                                integerEntries.Add((idx, wccTag, comment));
                        }
                    }

                    // Вставляем пакетно (с параметрами).
                    await InsertEntriesAsync(connection, "messages_binary", binaryEntries);
                    await InsertEntriesAsync(connection, "messages_integer", integerEntries);
                }
            }
        }

        /// <summary>
        /// Возвращает количество строк в таблице.
        /// </summary>
        public static async Task<int> GetTableCountAsync(string dbPath, string tableName)
        {
            using (var connection = new SQLiteConnection($"Data Source={dbPath}; Version=3;"))
            {
                await connection.OpenAsync();
                using (var command = new SQLiteCommand($"SELECT COUNT(*) FROM {tableName}", connection))
                {
                    return Convert.ToInt32(await command.ExecuteScalarAsync());
                }
            }
        }

        /// <summary>
        /// Создаёт вторую БД (маршрутизация) по данным из первой БД.
        /// </summary>
        public static async Task CreateAndFillSecondDatabaseAsync(string dbPath, string sourceDbPath)
        {
            using (var connection = new SQLiteConnection($"Data Source={dbPath}; Version=3;"))
            {
                await connection.OpenAsync();

                // Схема второй БД.
                const string schemaSql = @"
BEGIN TRANSACTION;
CREATE TABLE IF NOT EXISTS types (
    id INTEGER UNIQUE,
    name TEXT,
    comment TEXT
);
CREATE TABLE IF NOT EXISTS managers (
    id INTEGER UNIQUE,
    name TEXT UNIQUE,
    sending_socket TEXT NOT NULL UNIQUE,
    receiving_socket TEXT UNIQUE,
    description TEXT,
    PRIMARY KEY(id AUTOINCREMENT)
);
CREATE TABLE IF NOT EXISTS routing (
    id INTEGER NOT NULL UNIQUE,
    from_manager_id INTEGER NOT NULL,
    from_idx TEXT NOT NULL,
    to_manager_id INTEGER NOT NULL,
    to_idx TEXT NOT NULL,
    PRIMARY KEY(id AUTOINCREMENT),
    FOREIGN KEY(from_manager_id) REFERENCES managers(id)
);
INSERT OR IGNORE INTO types (id, name, comment) VALUES 
    (0, 'ANA_VT', 'Аналоговый тип.'), 
    (1, 'BIN_VT', 'Булевый тип'), 
    (2, 'INT_VT', 'Целочисленный тип');
INSERT OR IGNORE INTO managers (id, name, sending_socket, receiving_socket, description) VALUES 
    (1, 'DtsManager', 'ipc:///tmp/S_dts', 'ipc:///tmp/R_dts', 'Менеджер канала DTS.'),
    (2, 'WccManager', 'ipc:///tmp/S_wcc', 'ipc:///tmp/R_wcc', 'Менеджер взаимодействия с WinCC OA. '),
    (3, 'WCCOA_Mg_Gateway', 'tcp://192.168.11.151:5555', NULL, 'Менеджер взаимодействия с Desigo CC');
CREATE VIEW IF NOT EXISTS ROUTING_VIEW AS 
    SELECT MAN_FROM.name as from_manager, route.from_idx as from_idx, MAN_TO.name as to_manager, route.to_idx as to_idx 
    FROM routing as route 
    LEFT JOIN managers as MAN_FROM on route.from_manager_id = MAN_FROM.id 
    LEFT JOIN managers as MAN_TO on route.to_manager_id = MAN_TO.id;
COMMIT;";

                using (var command = new SQLiteCommand(schemaSql, connection))
                    await command.ExecuteNonQueryAsync();

                // Получаем данные из первой БД.
                var allIdx = await GetAllIdxFromFirstDatabaseAsync(sourceDbPath);

                // Готовим вставку маршрутов.
                using (var transaction = connection.BeginTransaction())
                using (var insertCmd = new SQLiteCommand("INSERT INTO routing (id, from_manager_id, from_idx, to_manager_id, to_idx) VALUES (@id, @fromId, @fromIdx, @toId, @toIdx)", connection, transaction))
                {
                    insertCmd.Parameters.Add("@id", DbType.Int32);
                    insertCmd.Parameters.Add("@fromId", DbType.Int32);
                    insertCmd.Parameters.Add("@fromIdx", DbType.String);
                    insertCmd.Parameters.Add("@toId", DbType.Int32);
                    insertCmd.Parameters.Add("@toIdx", DbType.String);

                    int id = 1;
                    int fromManagerId = 3;
                    int toManagerId = 1;

                    foreach (var item in allIdx)
                    {
                        string fromIdx = item.Item1;
                        int tagType = item.Item2;
                        string toIdx = $"{tagType}{item.Item3}";

                        insertCmd.Parameters["@id"].Value = id;
                        insertCmd.Parameters["@fromId"].Value = fromManagerId;
                        insertCmd.Parameters["@fromIdx"].Value = fromIdx;
                        insertCmd.Parameters["@toId"].Value = toManagerId;
                        insertCmd.Parameters["@toIdx"].Value = toIdx;

                        await insertCmd.ExecuteNonQueryAsync();
                        id++;
                    }

                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// Создаёт 4 базы (System00_1..System10_2) в указанной папке.
        /// </summary>
        public static async Task CreateAndFillSystemDatabasesAsync(string outputDirectory, string excelFilePath)
        {
            // Системы фиксированы — порядок сохраняем.
            string[] systems = { "System00_1", "System00_2", "System10_1", "System10_2" };

            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            // Считываем данные один раз — экономим время.
            var cabinetsInfo = ExcelParser.GetCabinetsInfo(excelFilePath);

            foreach (string sys in systems)
            {
                string dbFile = Path.Combine(outputDirectory, sys + ".db");

                using (var connection = new SQLiteConnection($"Data Source={dbFile}; Version=3;"))
                {
                    await connection.OpenAsync();
                    await EnsureMainSchemaAsync(connection);

                    var binEntries = new List<(string idx, string tag, string comment)>();
                    var intEntries = new List<(string idx, string tag, string comment)>();
                    var idxSet = new HashSet<string>();

                    foreach (var cabinetInfo in cabinetsInfo)
                    {
                        if (!TryParseCabinetInfo(cabinetInfo, out var cabinet, out var _, out var kkss))
                            continue;

                        foreach (string raw in kkss)
                        {
                            if (!TryParseMechEntry(raw, out var mechKks, out var kksNum, out var kksBuild))
                                continue;

                            string model = MechanismModelResolver.GetObjectModel(mechKks);
                            if (model == "AKKUYU_NO_MODEL")
                                continue;

                            var tags = BuildDefaultTags();

                            foreach (var tag in tags)
                            {
                                string baseIdx = $"{cabinet}_{mechKks}_{tag.Item1}";
                                string idx = baseIdx;

                                int counter = 1;
                                while (idxSet.Contains(idx) || await IdxExistsAsync(connection, idx))
                                {
                                    idx = $"{baseIdx}_{counter}";
                                    counter++;
                                }

                                idxSet.Add(idx);

                                string wccTag = $"{sys}:ManagementView_FieldNetworks_S7Plus_{kksBuild}_{cabinet}_{kksNum}.{tag.Item1}";
                                string comment = $"{tag.Item1} {kksBuild}_{cabinet}_{mechKks}";

                                if (tag.Item2 == "bool")
                                    binEntries.Add((idx, wccTag, comment));
                                else
                                    intEntries.Add((idx, wccTag, comment));
                            }
                        }
                    }

                    await InsertEntriesAsync(connection, "messages_binary", binEntries);
                    await InsertEntriesAsync(connection, "messages_integer", intEntries);

                    Debug.WriteLine($"[{sys}] DB готова: {binEntries.Count} bool + {intEntries.Count} int");
                }
            }
        }

        /// <summary>
        /// Вставка точек из DPL в messages_integer (с параметрами).
        /// </summary>
        public static void AddDataPointsToDatabase(string databasePath, List<(string dataPoint, string comment, string cabinetCode)> matchingData)
        {
            using (var connection = new SQLiteConnection($"Data Source={databasePath}; Version=3;"))
            {
                connection.Open();

                using (var transaction = connection.BeginTransaction())
                using (var command = new SQLiteCommand("INSERT INTO messages_integer (idx, wcc_oa_tag, comment) VALUES (@idx, @tag, @comment)", connection, transaction))
                {
                    command.Parameters.Add("@idx", DbType.String);
                    command.Parameters.Add("@tag", DbType.String);
                    command.Parameters.Add("@comment", DbType.String);

                    foreach (var data in matchingData)
                    {
                        // Если datapoint пустой — пропускаем, чтобы не ломать БД.
                        if (string.IsNullOrWhiteSpace(data.dataPoint))
                            continue;

                        string combinedComment = $"{data.comment} {data.cabinetCode}".Trim();

                        string idxSystem001 = GetUniqueIdx(connection, $"{data.dataPoint}_System001");
                        string idxSystem002 = GetUniqueIdx(connection, $"{data.dataPoint}_System002");
                        string idxSystem003 = GetUniqueIdx(connection, $"{data.dataPoint}_System1001");
                        string idxSystem004 = GetUniqueIdx(connection, $"{data.dataPoint}_System1002");

                        InsertPoint(command, idxSystem001, $"System00_1:{data.dataPoint}.Present_Value", combinedComment);
                        InsertPoint(command, idxSystem002, $"System00_2:{data.dataPoint}.Present_Value", combinedComment);
                        InsertPoint(command, idxSystem003, $"System10_1:{data.dataPoint}.Present_Value", combinedComment);
                        InsertPoint(command, idxSystem004, $"System10_2:{data.dataPoint}.Present_Value", combinedComment);
                    }

                    transaction.Commit();
                }
            }
        }

        /// <summary>
        /// Вставка одной строки с уже подготовленными параметрами.
        /// </summary>
        private static void InsertPoint(SQLiteCommand command, string idx, string tag, string comment)
        {
            command.Parameters["@idx"].Value = idx;
            command.Parameters["@tag"].Value = tag;
            command.Parameters["@comment"].Value = comment;
            command.ExecuteNonQuery();
        }

        /// <summary>
        /// Генерирует уникальный idx (с проверкой по БД).
        /// </summary>
        private static string GetUniqueIdx(SQLiteConnection connection, string baseIdx)
        {
            string idx = baseIdx;
            int counter = 1;
            while (IdxExists(connection, idx))
            {
                idx = $"{baseIdx}_{counter}";
                counter++;
            }
            return idx;
        }

        /// <summary>
        /// Проверка существования idx (синхронно).
        /// </summary>
        private static bool IdxExists(SQLiteConnection connection, string idx)
        {
            using (var command = new SQLiteCommand("SELECT COUNT(*) FROM messages_integer WHERE idx = @idx", connection))
            {
                command.Parameters.AddWithValue("@idx", idx);
                return Convert.ToInt32(command.ExecuteScalar()) > 0;
            }
        }

        /// <summary>
        /// Проверка существования idx (асинхронно, по трём таблицам).
        /// </summary>
        private static async Task<bool> IdxExistsAsync(SQLiteConnection connection, string idx)
        {
            using (var command = new SQLiteCommand("SELECT COUNT(*) FROM (SELECT idx FROM messages_analog WHERE idx = @idx UNION ALL SELECT idx FROM messages_binary WHERE idx = @idx UNION ALL SELECT idx FROM messages_integer WHERE idx = @idx)", connection))
            {
                command.Parameters.AddWithValue("@idx", idx);
                return Convert.ToInt32(await command.ExecuteScalarAsync()) > 0;
            }
        }

        /// <summary>
        /// Создаёт основные таблицы + types (если ещё нет).
        /// </summary>
        private static async Task EnsureMainSchemaAsync(SQLiteConnection connection)
        {
            const string schemaSql = @"
BEGIN TRANSACTION;
CREATE TABLE IF NOT EXISTS commands (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idx TEXT NOT NULL,
    type_id INTEGER NOT NULL DEFAULT 3,
    wcc_oa_tag TEXT NOT NULL,
    comment TEXT
);
CREATE TABLE IF NOT EXISTS messages_analog (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idx TEXT UNIQUE,
    wcc_oa_tag TEXT NOT NULL,
    comment TEXT
);
CREATE TABLE IF NOT EXISTS messages_binary (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idx TEXT UNIQUE,
    wcc_oa_tag TEXT NOT NULL,
    comment TEXT
);
CREATE TABLE IF NOT EXISTS messages_integer (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idx TEXT UNIQUE,
    wcc_oa_tag TEXT NOT NULL,
    comment TEXT
);
CREATE TABLE IF NOT EXISTS types (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    type_id INTEGER NOT NULL,
    type_name TEXT NOT NULL,
    comment INTEGER
);
INSERT OR IGNORE INTO types (type_id, type_name, comment) VALUES 
    (0, 'Ana_VT', 'Структура данных аналогового значения. AData (quality : 4 bytes; value : 4 bytes)'), 
    (1, 'Bin_VT', 'Структура данных двоичного значения. BData (quality : 4 bytes; value : 4 bytes)'), 
    (2, 'Int_VT', 'Структура данных целочисленного значения. IData (quality : 4 bytes; value : 4 bytes)'), 
    (3, 'Grp_VT', 'Структура данных группового значения. GData (size : 4 bytes; groupType : 4 bytes; value : 4100 bytes max)');
COMMIT;";

            using (var command = new SQLiteCommand(schemaSql, connection))
                await command.ExecuteNonQueryAsync();
        }

        /// <summary>
        /// Вставка набора записей с параметризацией.
        /// </summary>
        private static async Task InsertEntriesAsync(SQLiteConnection connection, string tableName, List<(string idx, string tag, string comment)> entries)
        {
            if (entries == null || entries.Count == 0)
                return;

            using (var transaction = connection.BeginTransaction())
            using (var command = new SQLiteCommand($"INSERT INTO {tableName} (idx, wcc_oa_tag, comment) VALUES (@idx, @tag, @comment)", connection, transaction))
            {
                command.Parameters.Add("@idx", DbType.String);
                command.Parameters.Add("@tag", DbType.String);
                command.Parameters.Add("@comment", DbType.String);

                foreach (var entry in entries)
                {
                    command.Parameters["@idx"].Value = entry.idx;
                    command.Parameters["@tag"].Value = entry.tag;
                    command.Parameters["@comment"].Value = entry.comment;
                    await command.ExecuteNonQueryAsync();
                }

                transaction.Commit();
            }
        }

        /// <summary>
        /// Читает idx из первой БД (для маршрутизации).
        /// </summary>
        private static async Task<List<Tuple<string, int, int>>> GetAllIdxFromFirstDatabaseAsync(string sourceDbPath)
        {
            var allIdx = new List<Tuple<string, int, int>>();

            using (var connection = new SQLiteConnection($"Data Source={sourceDbPath}; Version=3;"))
            {
                await connection.OpenAsync();

                const string sql = @"
SELECT idx, 0 AS type, id 
FROM messages_analog 
UNION ALL 
SELECT idx, 1 AS type, id 
FROM messages_binary 
UNION ALL 
SELECT idx, 2 AS type, id 
FROM messages_integer";

                using (var command = new SQLiteCommand(sql, connection))
                using (var reader = await command.ExecuteReaderAsync())
                {
                    while (await reader.ReadAsync())
                    {
                        string idx = reader.GetString(0);
                        int type = reader.GetInt32(1);
                        int id = reader.GetInt32(2);

                        allIdx.Add(new Tuple<string, int, int>(idx, type, id));
                    }
                }
            }

            return allIdx;
        }

        /// <summary>
        /// Безопасный парсинг строки шкафа (формат сохранён).
        /// </summary>
        private static bool TryParseCabinetInfo(string cabinetInfo, out string cabinet, out string controllerIp, out List<string> kkss)
        {
            cabinet = string.Empty;
            controllerIp = string.Empty;
            kkss = new List<string>();

            if (string.IsNullOrWhiteSpace(cabinetInfo))
                return false;

            var parts = cabinetInfo.Split(',');
            if (parts.Length < 2)
                return false;

            cabinet = CleanString(parts[0].Trim());
            controllerIp = CleanString(parts[1].Trim());

            // Механизмы (и всё, что после них) — оставляем как есть.
            kkss = parts.Skip(2).ToList();
            return true;
        }

        /// <summary>
        /// Безопасный разбор записи механизма.
        /// </summary>
        private static bool TryParseMechEntry(string raw, out string mechKks, out string mechNum, out string mechBld)
        {
            mechKks = string.Empty;
            mechNum = string.Empty;
            mechBld = string.Empty;

            if (string.IsNullOrWhiteSpace(raw))
                return false;

            var parts = raw.Split(' ');
            if (parts.Length >= 4 && string.IsNullOrWhiteSpace(parts[0]))
            {
                mechKks = parts[1].Trim();
                mechNum = parts[2].Trim();
                mechBld = parts[3].Trim();
                return true;
            }

            var tokens = raw.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length < 3)
                return false;

            mechKks = tokens[0].Trim();
            mechNum = tokens[1].Trim();
            mechBld = tokens[2].Trim();
            return true;
        }

        /// <summary>
        /// Стандартный набор тегов для каждого механизма.
        /// </summary>
        private static List<Tuple<string, string>> BuildDefaultTags()
        {
            return new List<Tuple<string, string>>
            {
                Tuple.Create("Alarm_bit0", "bool"),
                Tuple.Create("Alarm_word_1", "int"),
                Tuple.Create("Mode", "int"),
                Tuple.Create("OPRT", "int"),
                Tuple.Create("OPRT_INDEX", "int"),
                Tuple.Create("State", "int"),
                Tuple.Create("RASU_bits0", "bool"),
                Tuple.Create("RASU_bits1", "bool")
            };
        }

        /// <summary>
        /// Очистка строки от скобок (сохраняем поведение).
        /// </summary>
        private static string CleanString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;
            return input.Replace("(", "").Replace(")", "");
        }
    }
}
