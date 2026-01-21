using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace GUI_EXCEL_parser
{
    /// <summary>
    /// Слой экспорта CSV. Вся генерация файлов и строк — здесь.
    /// </summary>
    internal static class CsvExporter
    {
        /// <summary>
        /// Генерация отдельного CSV‑файла по одному шкафу.
        /// </summary>
        public static void GenerateCsv(List<string> lines, string outputDirectory)
        {
            GenerateCsv(lines, outputDirectory, null, null);
        }

        /// <summary>
        /// Генерация CSV с поддержкой логирования и резервного копирования.
        /// </summary>
        public static void GenerateCsv(List<string> lines, string outputDirectory, Action<string> logInfo, Action<string> logWarning)
        {
            // Безопасное извлечение заголовочных данных.
            string cabinetName = CleanString(lines.ElementAtOrDefault(0) ?? "Без_имени_шкафа");
            string controllerIp = CleanString(lines.ElementAtOrDefault(1) ?? "NO IP");
            string rawBuilding = lines.ElementAtOrDefault(2) ?? string.Empty;
            string buildingName = rawBuilding.Replace(" ", "");
            if (string.IsNullOrWhiteSpace(buildingName))
                buildingName = "Без_здания";
            string type = CleanString(lines.ElementAtOrDefault(3) ?? "Без_типа");

            // Механизмы идут после первых четырёх полей (формат сохраняем).
            List<string> kkss = lines.Skip(4).ToList();

            // Если механизмов нет — файл не создаём.
            if (!kkss.Any())
            {
                Console.WriteLine("Ошибка: Список механизмов пуст. CSV файл не будет создан.");
                logWarning?.Invoke("CSV не создан: список механизмов пуст.");
                return;
            }

            // Генерируем содержимое CSV.
            StringBuilder csvContent = GenerateCsvContent(kkss, cabinetName, controllerIp, buildingName, type);

            // Имя файла: учитываем отсутствие ключевых данных.
            string csvFileName =
                $"{cabinetName}_{(controllerIp.ToLower() == "no ip" ? "без_айпи" : controllerIp)}_{buildingName}_{(type == "Без_типа" ? "без_типа" : type)}.csv";

            if (cabinetName == "Без_имени_шкафа")
                csvFileName = "Без_имени_шкафа.csv";
            else if (controllerIp.ToLower() == "no ip" && buildingName == "Без_здания" && type == "Без_типа")
                csvFileName = $"{cabinetName}_без_айпи_и_здания_и_типа.csv";
            else if (controllerIp.ToLower() == "no ip" && buildingName == "Без_здания")
                csvFileName = $"{cabinetName}_без_айпи_и_здания.csv";
            else if (controllerIp.ToLower() == "no ip")
                csvFileName = $"{cabinetName}_без_айпи.csv";
            else if (buildingName == "Без_здания" && type == "Без_типа")
                csvFileName = $"{cabinetName}_без_здания_и_типа.csv";
            else if (buildingName == "Без_здания")
                csvFileName = $"{cabinetName}_без_здания.csv";
            else if (type == "Без_типа")
                csvFileName = $"{cabinetName}_без_типа.csv";

            string csvFilePath = Path.Combine(outputDirectory, csvFileName);

            try
            {
                if (File.Exists(csvFilePath))
                {
                    string backupPath = BackupFileIfExists(csvFilePath);
                    logWarning?.Invoke($"Существующий CSV сохранён в резерв: {backupPath}");
                }

                // Записываем с BOM (как было).
                using (var writer = new StreamWriter(csvFilePath, false, new UTF8Encoding(true)))
                {
                    writer.Write(csvContent.ToString());
                }
                Console.WriteLine($"CSV файл создан: {csvFilePath}");
                logInfo?.Invoke($"CSV файл создан: {csvFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при создании CSV файла: {ex.Message}");
                logWarning?.Invoke($"Ошибка при создании CSV файла: {ex.Message}");
            }
        }

        private static string BackupFileIfExists(string filePath)
        {
            string dir = Path.GetDirectoryName(filePath) ?? string.Empty;
            string name = Path.GetFileNameWithoutExtension(filePath);
            string ext = Path.GetExtension(filePath);
            string backupName = $"{name}_backup_{DateTime.Now:yyyyMMdd_HHmmss}{ext}";
            string backupPath = Path.Combine(dir, backupName);
            File.Move(filePath, backupPath);
            return backupPath;
        }

        /// <summary>
        /// Генерирует содержимое CSV (для одного шкафа).
        /// </summary>
        public static StringBuilder GenerateCsvContent(List<string> KKSIds, string cabinetName, string controllerIp, string buildingName, string type)
        {
            var csvContent = new StringBuilder();

            // Общая шапка.
            AppendStandardHeader(csvContent);

            // DEVICES
            AppendDevicesHeader(csvContent);
            AppendCabinetDeviceLine(csvContent, cabinetName, controllerIp, buildingName, type, KKSIds);
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");

            // POINTS
            AppendPointsHeader(csvContent);
            AppendCabinetPoints(csvContent, cabinetName, buildingName, type, KKSIds);

            // Хвост
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[VALUE_ASSIGNMENT],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"#[ObjectPath],[Property Name],[Property Value],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");

            return csvContent;
        }

        /// <summary>
        /// Добавляет строку DEVICES для одного шкафа (формат как в обычной конвертации).
        /// </summary>
        public static void AppendCabinetDeviceLine(StringBuilder csvContent, string cabinetName, string controllerIp, string buildingName, string type, List<string> kkss)
        {
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},{GetDistinctBuildingsWithExclusion(kkss, buildingName) + cabinetName + type},S7-1500,{controllerIp},S7ONLINE,S7Plus$Online,Online,TRUE,PLC_{buildingName + "_" + cabinetName},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
        }

        /// <summary>
        /// Добавляет блок POINTS для одного шкафа (формат как в обычной конвертации).
        /// </summary>
        public static void AppendCabinetPoints(StringBuilder csvContent, string cabinetName, string buildingName, string type, List<string> kkss)
        {
            AppendCabinetDiagnostics(csvContent, buildingName, cabinetName, type);

            foreach (string name in kkss)
            {
                if (!TryParseMechEntry(name, out var mechKks, out var kksNumber, out var buildingNumber))
                    continue;

                string model = MechanismModelResolver.GetObjectModel(mechKks);
                if (model == "AKKUYU_NO_MODEL")
                    continue;

                string dispName = GetBuildingDifference(buildingNumber, buildingName) + mechKks;

                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},{dispName},DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.KKS,String,IO,FALSE,COV,pollGr_3,{model},KKS,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingName}\\Mechanisms\\{buildingName + "_" + cabinetName}\\,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].OPRT.OPRT_INDEX,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT_INDEX,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.Mode,Uint,IO,FALSE,COV,pollGr_3,{model},Mode,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.State,Uint,IO,FALSE,COV,pollGr_3,{model},State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

                if (model.Contains("AKKUYU_VALVE"))
                {
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.Extinguishing_State,Uint,IO,FALSE,COV,pollGr_3,{model},Extinguishing_State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                }

                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].ALM.Alarm_word_1,Uint,IO,FALSE,COV,pollGr_3,{model},Alarm_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].OPRT.OPRT,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].ALM.Alarm.Alarm[0],bool,IO,FALSE,COV,pollGr_3,{model},Alarm_bit0,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.RASU_bits.RASU_bits[0],bool,IO,FALSE,COV,pollGr_3,{model},RASU_bits0,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.RASU_bits.RASU_bits[1],bool,IO,FALSE,COV,pollGr_3,{model},RASU_bits1,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                csvContent.AppendLine(" ");
            }
        }

        /// <summary>
        /// Общая шапка CSV (HEADER/DRIVER/FILEVERSION/etc).
        /// </summary>
        public static void AppendStandardHeader(StringBuilder csvContent)
        {
            csvContent.AppendLine(@"[HEADER],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"#ListSeparator,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("," + ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#DecimalSeparator" + ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("." + ",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[DRIVER],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"S7Plus,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"#For detailed description of tags and columns, please refer to Simatic_S7 EM documentation,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[FILEVERSION],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"# File Version (MP 4.0 Ver. 1),,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"4,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[TEXTS],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"# Table Name,# Index,# Text,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText1,0,mg,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText1,1,g,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText1,2,kg,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText1,3,quintal,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText1,4,ton,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText1,5,kton,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText2,0,Normal,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText2,1,Alarm,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText2,2,Alarm Reset,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"TxG_StateText2,3,Alarm Closed,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[HIERARCHY],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"#[Name],[Description],[LogicalHierarchy],[UserHierarchy],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[SMOOTHING],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"#[SmoothingConfigName],[SmoothingType],[Deadband],[IsRelative],[Seconds],[Milliseconds],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"SmoothingConfig_01,oldnew,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"SmoothingConfig_02,oldnewandtime,,,10,10,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"SmoothingConfig_03,oldnewortime,,,9,10,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"SmoothingConfig_05,time,,,10,10,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"SmoothingConfig_04,value,4,FALSE,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"SmoothingConfig_06,valueandtime,3,TRUE,10,10,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"SmoothingConfig_07,valueortime,4,TRUE,10,10,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[POLLGROUPS],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"#[PollGroup Name],[PollInterval in ms],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"pollGr_19,2100,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[POLLGROUPS_COV],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"#[PollGroup Name],[PollInterval in ms],[Unsubscribe],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"PollGr_3,2000,FALSE,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[LIBRARY],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"BA_Device_S7Plus_HQ_1,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
        }

        /// <summary>
        /// Шапка DEVICES.
        /// </summary>
        public static void AppendDevicesHeader(StringBuilder csvContent)
        {
            csvContent.AppendLine("[DEVICES],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#[DeviceName],[DeviceDescription],[S7_Type],[IP_address],[AccessPointS7],[Projectname],[StationName],[Establish Connection],[Alias],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],,,,,,,,,,,,,,,,,,,,,,,,,,");
        }

        /// <summary>
        /// Шапка POINTS.
        /// </summary>
        public static void AppendPointsHeader(StringBuilder csvContent)
        {
            csvContent.AppendLine("[POINTS],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,");
            csvContent.AppendLine("#[ParentDeviceName],[Name],[Description],[Address],[DataType],[Direction],[LowLevelComparison],[PollType],[PollGroup],[ObjectModel],[Property],[Alias],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],[Min],[Max],[MinRaw],[MaxRaw],[MinEng],[MaxEng],[Resolution],[Unit],[UnitTextGroup],[StateText],[ActivityLog],[ValueLog],[Smoothing],[AlarmClass],[AlarmType],[AlarmValue],[EventText],[NormalText],[UpperHysteresis],[LowerHysteresis],[NoAlarmOn],[LogicalHierarchy],[UserHierarchy]");
        }

        /// <summary>
        /// Диагностические точки шкафа + строки зон.
        /// </summary>
        public static void AppendCabinetDiagnostics(StringBuilder csvContent, string buildingName, string cabinetName, string type)
        {
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,ДиагностикаЗАМЕНА,DB_HMI_IO_Mechanisms.Cabinet.KKS,String,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,KKS,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL1_Power_On,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit0,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL2_Automation_Disabled,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit1,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL3_Malfunction,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit2,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL4_Fire,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit3,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL5_Start,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit4,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL6_Maintenance,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit5,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Door_1_Opened,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit6,,,,,,,,,,,,,,,,TxG_AKK_CabDoor1,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.State_word_1,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,State_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

            // Если тип содержит «АПТС», добавляем дополнительные диагностические строки.
            if (!string.IsNullOrWhiteSpace(type) && type.Contains("АПТС"))
            {
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Door_2_Opened,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit7,,,,,,,,,,,,,,,,TxG_AKK_CabDoor2,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HMI_Alarm_Word_2,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,Alarm_word_2,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
            }

            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HMI_Alarm_Word_1,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,Alarm_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Watchdog,Dint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,ConnectionPLC,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.OPRT,int,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,OPRT,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

            // Зоны (две строки).
            csvContent.AppendLine();
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Zones_string_1,Помещение_ПожарЗАМЕНА,FB_Zones_DB.Zones_string_1,String,IO,FALSE,COV,pollGr_3,,,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Zones_string_2,Помещение_Пожар_ПодтверждениеЗАМЕНА,FB_Zones_DB.Zones_string_2,String,IO,FALSE,COV,pollGr_3,,,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine();
        }

        /// <summary>
        /// Склейка описания устройства с корректными пробелами.
        /// </summary>
        public static string BuildDeviceDescription(string distinct, string cabinet, string type)
        {
            string res = string.Empty;
            if (!string.IsNullOrWhiteSpace(distinct)) res = distinct.Trim();
            if (!string.IsNullOrWhiteSpace(cabinet)) res += (res.Length > 0 ? " " : "") + cabinet.Trim();
            if (!string.IsNullOrWhiteSpace(type)) res += (res.Length > 0 ? " " : "") + type.Trim();
            return res;
        }

        /// <summary>
        /// Парсинг триплета "KKS / номер / здание".
        /// </summary>
        public static (string mechKks, string mechNum, string mechBld) ParseMechTriplet(string raw, string fallbackBuilding = "")
        {
            // Сохраняем пустые элементы (как раньше), чтобы не ломать формат.
            var parts = raw.Split(new[] { ' ' }, StringSplitOptions.None);
            string mechKks = parts.ElementAtOrDefault(0)?.Trim() ?? "";
            string mechNum = parts.ElementAtOrDefault(1)?.Trim() ?? "";
            string mechBld = parts.ElementAtOrDefault(2)?.Trim() ?? fallbackBuilding;
            return (mechKks, mechNum, mechBld);
        }

        /// <summary>
        /// Проверка «числовости» строки.
        /// </summary>
        public static bool IsUint(string s) => !string.IsNullOrWhiteSpace(s) && s.All(char.IsDigit);

        /// <summary>
        /// Возвращает блок ("00","10","20","30","40") по зданию/ККС.
        /// </summary>
        public static string GetBlockFromBuildingOrKks(string building, string cabinetKks)
        {
            string Normalize(string src)
            {
                if (string.IsNullOrWhiteSpace(src)) return null;
                char d0 = src[0];
                if (!char.IsDigit(d0)) return null;

                switch (d0)
                {
                    case '0': return "00";
                    case '1': return "10";
                    case '2': return "20";
                    case '3': return "30";
                    case '4': return "40";
                    default: return null;
                }
            }

            var byBuilding = Normalize(building);
            if (byBuilding != null) return byBuilding;
            return Normalize(cabinetKks);
        }

        /// <summary>
        /// Подпись блока для имени файла.
        /// </summary>
        public static string GetBlockLabel(string block)
        {
            switch (block)
            {
                case "00": return "общестанционка";
                case "10": return "первый_блок";
                case "20": return "второй_блок";
                case "30": return "третий_блок";
                case "40": return "четвёртый_блок";
                default: return "неизвестно";
            }
        }

        /// <summary>
        /// Разница зданий для отображения.
        /// </summary>
        public static string GetBuildingDifference(string buildingNumber, string buildingName)
        {
            if (buildingNumber != buildingName)
                return $"({buildingNumber})";
            return string.Empty;
        }

        /// <summary>
        /// Возвращает список отличающихся зданий (в скобках), кроме текущего.
        /// </summary>
        public static string GetDistinctBuildingsWithExclusion(List<string> KKSIds, string buildingName)
        {
            var buildingNumbers = new HashSet<string>();

            foreach (var name in KKSIds)
            {
                if (!TryParseMechEntry(name, out _, out _, out var bld))
                    continue;
                if (!string.IsNullOrWhiteSpace(bld))
                    buildingNumbers.Add(bld);
            }

            buildingNumbers.Remove(buildingName);

            return buildingNumbers.Count > 0
                ? "(" + string.Join(" ", buildingNumbers) + ")"
                : string.Empty;
        }

        /// <summary>
        /// Используется для экспортов по блокам — исключение «своего» здания.
        /// </summary>
        public static string GetDistinctBuildingsForBlockExport(List<string> kkss, string deviceBuilding)
        {
            var buildings = new HashSet<string>();
            foreach (var raw in kkss)
            {
                var (_, _, bld) = ParseMechTriplet(raw, deviceBuilding);
                if (!string.IsNullOrWhiteSpace(bld))
                    buildings.Add(bld);
            }

            buildings.Remove(deviceBuilding);
            return buildings.Count == 0 ? string.Empty : "(" + string.Join(" ", buildings) + ")";
        }

        /// <summary>
        /// Очистка строки от скобок.
        /// </summary>
        public static string CleanString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;
            return input.Replace("(", "").Replace(")", "");
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

            // Вариант 1: старая разбивка с пустым первым элементом.
            var parts = raw.Split(' ');
            if (parts.Length >= 4 && string.IsNullOrWhiteSpace(parts[0]))
            {
                mechKks = parts[1].Trim();
                mechNum = parts[2].Trim();
                mechBld = parts[3].Trim();
                return true;
            }

            // Вариант 2: нормальная разбивка.
            var tokens = raw.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length < 3)
                return false;

            mechKks = tokens[0].Trim();
            mechNum = tokens[1].Trim();
            mechBld = tokens[2].Trim();
            return true;
        }
    }
}
