using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.SQLite;
using System.Threading.Tasks;
using System.Diagnostics;


namespace GUI_EXCEL_parser
{
    public partial class Converter : Form
    {
        public Converter()
        {
            InitializeComponent();
            InitializeCustomComponents();
        }

        // Метод для инициализации пользовательских компонентов
        private void InitializeCustomComponents()
        {


            this.btnParseDpl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnConvert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateSecondDatabase.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
        }

        // Обработчик кнопки выбора файла Excel
        private void btnSelectExcelFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtExcelFilePath.Text = openFileDialog.FileName;
                }
            }
        }

        // Обработчик кнопки выбора папки для сохранения CSV файлов
        private void btnSelectOutputDirectory_Click(object sender, EventArgs e)
        {
            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    txtOutputDirectory.Text = folderDialog.SelectedPath;
                }
            }
        }

        // Обработчик кнопки начала конвертации
        private async void btnConvert_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Вы уверены, что хотите создать файлы CSV?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }

            string excelFilePath = txtExcelFilePath.Text;
            string outputDirectory = txtOutputDirectory.Text;
            try
            {
                // Обновление текста, состояния кнопки и отображение PictureBox
                btnConvert.Text = "Конвертация...";
                btnConvert.Enabled = false;
                progressBar1.Value = 0;

                var cabinetsInfo = await Task.Run(() => GetCabinetsInfo(excelFilePath));
                int totalFiles = 0;
                int totalMechanisms = 0;
                int cabinetsWithoutIP = 0;
                int cabinetsWithoutMechanisms = 0;
                progressBar1.Maximum = cabinetsInfo.Count;

                foreach (var cabinetInfo in cabinetsInfo)
                {
                    var lines = cabinetInfo.Split(',');
                    if (lines.Length < 3)
                    {
                        cabinetsWithoutMechanisms++;
                        continue;
                    }
                    if (lines[1].ToLower() == "no ip")
                    {
                        cabinetsWithoutIP++;
                    }
                    await Task.Run(() => GenerateCsv(lines.ToList(), outputDirectory));
                    totalFiles++;
                    totalMechanisms += lines.Length - 2;
                    progressBar1.Value++;
                }

                // Восстановление текста, состояния кнопки и скрытие PictureBox
                btnConvert.Text = "Конвертировать";
                btnConvert.Enabled = true;


                // Сообщение об успешном завершении
                MessageBox.Show($"Преобразование завершено.\n" +
                                $"Всего создано файлов: {totalFiles}\n" +
                                $"Всего механизмов: {totalMechanisms}\n" +
                                $"Шкафов без IP: {cabinetsWithoutIP}\n" +
                                $"Шкафов без механизмов: {cabinetsWithoutMechanisms}",
                                "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnConvert.Text = "Конвертировать";
                btnConvert.Enabled = true;
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Метод для получения информации о шкафах и механизмах
        static List<string> GetCabinetsInfo(string filePath)
        {
            var cabinetsSheetName = "Шкафы";
            var mechSheetName = "Мех-мы";

            var mechData = ReadMechData(filePath, mechSheetName);
            var cabinetData = ReadCabinetData(filePath, cabinetsSheetName);
            var cabinetsInfo = new List<string>();

            foreach (var kks in mechData.Keys)
            {
                if (cabinetData.ContainsKey(kks))
                {
                    var line = $"{kks},{cabinetData[kks]}";
                    foreach (var mech in mechData[kks])
                    {
                        line += $", {mech}";
                    }
                    cabinetsInfo.Add(line);
                }
            }

            return cabinetsInfo;
        }


        // Функция для очистки строки от нежелательных символов и пробелов
        private static string CleanString(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // Удаляем пробелы и символы '(' и ')'
            return input.Replace("(", "").Replace(")", "");
        }


        // Метод для создания CSV файла
        static void GenerateCsv(List<string> lines, string outputDirectory)
        {
            // Проверяем наличие данных и сразу очищаем от пробелов и скобок
            string cabinetName = CleanString(lines.ElementAtOrDefault(0) ?? "Без_имени_шкафа");
            string controllerIp = CleanString(lines.ElementAtOrDefault(1) ?? "NO IP");
            string buildingName = lines.ElementAtOrDefault(2).Replace(" ", "") ?? "Без_здания";
            string type = CleanString(lines.ElementAtOrDefault(3) ?? "Без_типа");


            List<string> kkss = lines.Skip(4).ToList();

            // Проверка на наличие механизмов
            if (!kkss.Any())
            {
                Console.WriteLine("Ошибка: Список механизмов пуст. CSV файл не будет создан.");
                return;
            }

            // Генерируем содержимое CSV с учетом всех данных
            StringBuilder csvContent = GenerateCsvContent(kkss, cabinetName, controllerIp, buildingName, type);

            // Формируем имя файла с учетом отсутствующих данных
            string csvFileName = $"{cabinetName}_{(controllerIp.ToLower() == "no ip" ? "без_айпи" : controllerIp)}_{buildingName}_{(type == "Без_типа" ? "без_типа" : type)}.csv";

            // Если какие-то ключевые данные отсутствуют, добавляем отметки в имени файла
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
                // Открываем StreamWriter с указанием кодировки UTF-8 с BOM
                using (StreamWriter writer = new StreamWriter(csvFilePath, false, new System.Text.UTF8Encoding(true)))
                {
                    writer.Write(csvContent.ToString());
                }

                Console.WriteLine($"CSV файл создан: {csvFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при создании CSV файла: {ex.Message}");
            }
        }

        // Метод для получения данных по ККС шкафа из листа "Алгоритмы"
        public static List<Dictionary<string, string>> GetAlgorithmDataByKKS(string filePath, string kksCabinet)
        {
            var result = new List<Dictionary<string, string>>();
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet("Алгоритмы");

                    // Поиск номеров колонок
                    var buildingColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Здание"))?.Address.ColumnNumber;
                    var floorColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Этаж (отметка)"))?.Address.ColumnNumber;
                    var roomColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Помещение"))?.Address.ColumnNumber;
                    var fullRoomColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Полное название помещения"))?.Address.ColumnNumber;
                    var zoneColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("№ зоны"))?.Address.ColumnNumber;
                    var sourceKKSColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("ККС шкафа источника"))?.Address.ColumnNumber;
                    var destKKSColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("ККС шкафа назначения"))?.Address.ColumnNumber;

                    // Проверка наличия всех нужных колонок
                    if (!buildingColumn.HasValue || !floorColumn.HasValue || !roomColumn.HasValue || !fullRoomColumn.HasValue || !zoneColumn.HasValue || !sourceKKSColumn.HasValue || !destKKSColumn.HasValue)
                        return result;

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        var destKKSValue = row.Cell(destKKSColumn.Value).GetValue<string>().Trim();

                        if (destKKSValue == kksCabinet)
                        {
                            var dataRow = new Dictionary<string, string>
                            {
                                ["Здание"] = row.Cell(buildingColumn.Value).GetValue<string>().Trim(),
                                ["Этаж (отметка)"] = row.Cell(floorColumn.Value).GetValue<string>().Trim(),
                                ["Помещение"] = row.Cell(roomColumn.Value).GetValue<string>().Trim(),
                                ["Полное название помещения"] = row.Cell(fullRoomColumn.Value).GetValue<string>().Trim(),
                                ["№ зоны"] = row.Cell(zoneColumn.Value).GetValue<string>().Trim(),
                                ["ККС шкафа источника"] = row.Cell(sourceKKSColumn.Value).GetValue<string>().Trim(),
                                ["ККС шкафа назначения"] = destKKSValue
                            };

                            result.Add(dataRow);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка при чтении файла '{filePath}': {ex.Message}");
            }

            return result;
        }


        // Метод для чтения данных о механизмах
        static Dictionary<string, List<string>> ReadMechData(string filePath, string sheetName)
        {
            var data = new Dictionary<string, List<string>>();
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(sheetName);
                var kksColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("KKS шкафа"))?.Address.ColumnNumber;
                var mechColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("KKS механизма"))?.Address.ColumnNumber;
                var mechNumColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Desigo number"))?.Address.ColumnNumber;
                var BuildingNumColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Здание"))?.Address.ColumnNumber;

                if (!kksColumn.HasValue || !mechColumn.HasValue || !mechNumColumn.HasValue) return data;

                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    var kksValue = row.Cell(kksColumn.Value).GetValue<string>().Trim();
                    var mechValue = row.Cell(mechColumn.Value).GetValue<string>().Trim();
                    var mechNum = row.Cell(mechNumColumn.Value).GetValue<string>().Trim();
                    var BuildingNum = row.Cell(BuildingNumColumn.Value).GetValue<string>().Trim();

                    if (!string.IsNullOrWhiteSpace(kksValue) && !string.IsNullOrWhiteSpace(mechValue))
                    {
                        if (!data.ContainsKey(kksValue))
                        {
                            data[kksValue] = new List<string>();
                        }
                        data[kksValue].Add(mechValue + " " + mechNum + " " + BuildingNum);
                    }
                }
            }
            return data;
        }

        // Генерация содержимого CSV файла
        static StringBuilder GenerateCsvContent(List<string> KKSIds, string cabinetName, string controllerIp, string buildingName, string type)
        {
            StringBuilder csvContent = new StringBuilder();
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
            //csvContent.AppendLine(@"Bldg1,Building 1,\,\,Measure,0,1,200,201,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Bldg2,Building 2,\,\,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Flr1,Floor 1,\Base1\Bldg1\,\Base1\Bldg1\,Measure,0,1,200,201,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Flr2,Floor 2,\Base2\Bldg2\,\Base2\Bldg2\,BoilerDigital,,,400,402,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Base1,Base 1,\,\,SensorRoomCO2,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Base2,Base 2,\,\,,20,21,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Base3,Base 3,\,\,Measure,0,1,200,201,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Bldg3,Building 3,\Base3\,\Base3\,Measure,0,1,200,201,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Flr3,Floor 3,\Base3\Bldg3\,\Base3\Bldg3\,Measure,0,1,200,201,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            //csvContent.AppendLine(@"Section3,Section 3,\Base3\Bldg3\Flr3\,\Base3\Bldg3\Flr3\,Measure,0,1,200,201,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
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

            // Статические заголовки DEVICE
            csvContent.AppendLine("[DEVICES],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#[DeviceName],[DeviceDescription],[S7_Type],[IP_address],[AccessPointS7],[Projectname],[StationName],[Establish Connection],[Alias],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},{GetDistinctBuildingsWithExclusion(KKSIds, buildingName) + cabinetName + type},S7-1500,{controllerIp},S7ONLINE,S7Plus$Online,Online,TRUE,PLC_{buildingName + "_" + cabinetName},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");


            //00UAC_00CMM11,00CMM11 ТипЗАМЕНА 008_25,S7-1500,10.100.11.51,S7ONLINE,S7Plus$Online,Online,TRUE,PLC_00UAC_00CMM11,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
            //00CMM11 ТипЗАМЕНА 008_25





            // Данные POINTS
            csvContent.AppendLine("[POINTS],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,");
            csvContent.AppendLine("#[ParentDeviceName],[Name],[Description],[Address],[DataType],[Direction],[LowLevelComparison],[PollType],[PollGroup],[ObjectModel],[Property],[Alias],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],[Min],[Max],[MinRaw],[MaxRaw],[MinEng],[MaxEng],[Resolution],[Unit],[UnitTextGroup],[StateText],[ActivityLog],[ValueLog],[Smoothing],[AlarmClass],[AlarmType],[AlarmValue],[EventText],[NormalText],[UpperHysteresis],[LowerHysteresis],[NoAlarmOn],[LogicalHierarchy],[UserHierarchy]");

            // Добавляем переменные для диагностики шкафа (Diagnostic_Cab)                                 
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,ДиагностикаЗАМЕНА,DB_HMI_IO_Mechanisms.Cabinet.KKS,String,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,KKS,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL1_Power_On,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit0,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL2_Automation_Disabled,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit1,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL3_Malfunction,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit2,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL4_Fire,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit3,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL5_Start,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit4,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL6_Maintenance,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit5,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Door_1_Opened,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit6,,,,,,,,,,,,,,,,TxG_AKK_CabDoor1,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.State_word_1,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,State_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");


            // Проверяем наличие "АПТС" в типе и добавляем строку при соблюдении условия
            if (type.Contains("АПТС"))
            {
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Door_2_Opened,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit7,,,,,,,,,,,,,,,,TxG_AKK_CabDoor2,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
                csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HMI_Alarm_Word_2,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,Alarm_word_2,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
            }
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HMI_Alarm_Word_1,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,Alarm_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Watchdog,Dint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,ConnectionPLC,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.OPRT,int,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,OPRT,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

            // Пропускаем пустую строку и добавляем переменные для зон
            csvContent.AppendLine();
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Zones_string_1,Помещение_ПожарЗАМЕНА,FB_Zones_DB.Zones_string_1,String,IO,FALSE,COV,pollGr_3,,,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            csvContent.AppendLine($"{buildingName + "_" + cabinetName},Zones_string_2,Помещение_Пожар_ПодтверждениеЗАМЕНА,FB_Zones_DB.Zones_string_2,String,IO,FALSE,COV,pollGr_3,,,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");

            csvContent.AppendLine();
            int device_number = 1;


            foreach (string name in KKSIds)
            {
                string model = GetObjectModel(name.Split(' ')[1]);


                if (model != "AKKUYU_NO_MODEL")
                {
                    string[] parts = name.Split(' ');

                    string kksname = parts[1];
                    string kksnumber = parts[2];
                    string buildingnumber = parts[3];

                    // csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},{GetBuildingDifference(buildingnumber, buildingName) + kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.KKS,String,IO,FALSE,COV,pollGr_3,{model},KKS,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{buildingnumber +"_"+ cabinetName}\\,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},{GetBuildingDifference(buildingnumber, buildingName) + kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.KKS,String,IO,FALSE,COV,pollGr_3,{model},KKS,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingName}\\Mechanisms\\{buildingName + "_" + cabinetName}\\,");
                    //10UBA_10CMS01,18,10SGK33AA007,DB_HMI_IO_Mechanisms.Devices.Devices[18].PRM.KKS,String,IO,FALSE,COV,pollGr_3,AKKUYU_VALVE                                                                         ,KKS,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\       10UBA       \Mechanisms\              10UBA_10CMS01\,



                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].OPRT.OPRT_INDEX,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT_INDEX,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.Mode,Uint,IO,FALSE,COV,pollGr_3,{model},Mode,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.State,Uint,IO,FALSE,COV,pollGr_3,{model},State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

                    // Проверяем наличие "AKKUYU_VALVE" в типе и добавляем строку при соблюдении условия
                    if (model.Contains("AKKUYU_VALVE"))
                    {
                        csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.Extinguishing_State,Uint,IO,FALSE,COV,pollGr_3,{model},Extinguishing_State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                     //   csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.Extinguishing_State,Uint,IO,FALSE,COV,pollGr_3,{model},Extinguishing_State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,"); необходимо  добавить тег РЕДУНДАНТ


                    }

                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].ALM.Alarm_word_1,Uint,IO,FALSE,COV,pollGr_3,{model},Alarm_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].OPRT.OPRT,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].ALM.Alarm.Alarm[0],bool,IO,FALSE,COV,pollGr_3,{model},Alarm_bit0,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.RASU_bits.RASU_bits[0],bool,IO,FALSE,COV,pollGr_3,{model},RASU_bits0,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.RASU_bits.RASU_bits[1],bool,IO,FALSE,COV,pollGr_3,{model},RASU_bits1,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,,");



                    csvContent.AppendLine(" ");

                    device_number++;
                }
            }
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"[VALUE_ASSIGNMENT],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine(@"#[ObjectPath],[Property Name],[Property Value],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");

            return csvContent;
        }

        // Метод для получения объектной модели по механизму
        static string GetObjectModel(string mechanism)
        {
            // Если механизм пустой или null, возвращаем значение по умолчанию
            if (string.IsNullOrEmpty(mechanism))
            {
                return "AKKUYU_NO_MODEL";
            }

            // Словарь для соответствия кодов оборудования и моделей механизмов
            var equipmentMapping = new Dictionary<string, Dictionary<string, string>>
    {
        // Коды оборудования "AA" и соответствующие им системы
        { "AA", new Dictionary<string, string>
            {
                { "SAC", "AKKUYU_GATE" },  // Ворота
                { "SGC", "AKKUYU_VALVE" }, // Клапан
                { "SGK", "AKKUYU_VALVE" },

                { "SGA", "AKKUYU_VALVE" }, // Клапан
                { "GKE", "AKKUYU_VALVE" }, // Клапан

                { "KLE", "AKKUYU_GATE" }, // Клапан

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
        { "AN", new Dictionary<string, string>
            {
                { "SAC", "AKKUYU_FAN" },   // Вентилятор
                { "SAM", "AKKUYU_FAN" },   // Вентилятор
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
        { "AP", new Dictionary<string, string>
            {
                { "SGA", "AKKUYU_MOTOR" },
                { "SGK", "AKKUYU_MOTOR" }   // Насос
            }
        },
        // Коды оборудования "CP" могут быть добавлены в будущем
        { "CP", new Dictionary<string, string>
            {
                // Пример: { "SAC", "AKKUYU_CUSTOM_CP" }
            }
        }
    };

            // Проходим по всем кодам оборудования (например, "AA", "AN", "AP")
            foreach (var equipmentEntry in equipmentMapping)
            {
                // Если код оборудования (например, "AA") присутствует в строке механизма
                if (mechanism.Contains(equipmentEntry.Key))
                {
                    // Проходим по кодам систем (например, "SAC", "SGA", "KLS")
                    foreach (var systemEntry in equipmentEntry.Value)
                    {
                        // Если код системы присутствует в строке механизма
                        if (mechanism.Contains(systemEntry.Key))
                        {
                            // Возвращаем соответствующую модель механизма
                            return systemEntry.Value;
                        }
                    }
                }
            }

            // Если механизм не найден, возвращаем значение по умолчанию
            return "AKKUYU_NO_MODEL";
        }

        // Метод для проверки здания
        public static string GetBuildingDifference(string buildingNumber, string buildingName)
        {
            if (buildingNumber != buildingName)
            {
                return $"({buildingNumber})";
            }

            return string.Empty; // Возвращаем пустую строку, если здания одинаковы
        }

        // Метод для проверки уникальности зданий
        public static string GetDistinctBuildingsWithExclusion(List<string> KKSIds, string buildingName)
        {
            var buildingNumbers = new HashSet<string>();

            foreach (var name in KKSIds)
            {
                string[] parts = name.Split(' ');
                if (parts.Length > 3)
                {
                    buildingNumbers.Add(parts[3]);
                }
            }

            buildingNumbers.Remove(buildingName); // Исключаем указанное здание

            if (buildingNumbers.Count > 0)
            {
                return "(" + string.Join(" ", buildingNumbers) + ")";
            }

            return string.Empty; // Возвращаем пустую строку, если все здания одинаковы
        }


        //Метод получения данных для Шкафа
        public static Dictionary<string, (string IP, string Building, string Type, string VEKSH)> ReadCabinetData(string filePath, string sheetName)
        {
            var data = new Dictionary<string, (string IP, string Building, string Type, string VEKSH)>();
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(sheetName);

                    // Поиск номеров колонок
                    var kksColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("KKS"))?.Address.ColumnNumber;
                    var ipColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains(@"IP S7_A1"))?.Address.ColumnNumber;
                    var buildingColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Здание"))?.Address.ColumnNumber;
                    var typeColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("Тип"))?.Address.ColumnNumber;
                    var vekshColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("ВЕКШ"))?.Address.ColumnNumber;

                    // Проверка наличия всех нужных колонок
                    if (!kksColumn.HasValue || !ipColumn.HasValue || !buildingColumn.HasValue || !typeColumn.HasValue || !vekshColumn.HasValue)
                        return data;

                    foreach (var row in worksheet.RowsUsed().Skip(1))
                    {
                        // Получение данных из строк
                        var kksValue = row.Cell(kksColumn.Value).GetValue<string>().Trim();
                        var ipValue = row.Cell(ipColumn.Value).GetValue<string>().Trim();
                        var buildingValue = row.Cell(buildingColumn.Value).GetValue<string>().Trim().Replace(" ", "");
                        var typeValue = row.Cell(typeColumn.Value).GetValue<string>().Trim();
                        var vekshValue = row.Cell(vekshColumn.Value).GetValue<string>().Trim();

                        // Заполнение значений по умолчанию, если данные пустые
                        if (!string.IsNullOrWhiteSpace(kksValue) && !data.ContainsKey(kksValue))
                        {
                            if (string.IsNullOrWhiteSpace(ipValue))
                                ipValue = "NO IP";
                            if (string.IsNullOrWhiteSpace(buildingValue))
                                buildingValue = "UNKNOWN BUILDING";
                            if (string.IsNullOrWhiteSpace(vekshValue))
                                vekshValue = "NO VEKSH";

                            // Обработка типа
                            var processedType = ProcessType(typeValue, vekshValue);

                            // Добавление данных в словарь
                            data[kksValue] = (ipValue, buildingValue, processedType, vekshValue);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Произошла ошибка при чтении файла '{filePath}': {ex.Message}");
            }

            return data;
        }

        // Метод обработки типа
        private static string ProcessType(string typeValue, string vekshValue)
        {
            string vekshSuffix = "00"; // Значение по умолчанию

            if (typeValue.Contains("АПТС") && !string.IsNullOrWhiteSpace(vekshValue) && vekshValue.StartsWith("ВЕКШ."))
            {
                var parts = vekshValue.Split('-');
                if (parts.Length > 1)
                {
                    vekshSuffix = parts[parts.Length - 1]; // Последняя часть после тире
                }
            }

            if (typeValue.Contains("АПТС"))
            {
                return $"АПТС-{vekshSuffix}ЗАМЕНА";
            }

            if (typeValue.StartsWith("Type"))
            {
                typeValue = typeValue.Replace("Type", "ТипЗАМЕНА")
                                     .Replace('.', '_')
                                     .Replace("*", "");

                return typeValue;
            }

            return "Неизвестный тип";
        }

        // Обработчик кнопки для отображения информации о программе
        private void btnInfo_Click(object sender, EventArgs e)
        {
            string infoMessage = "EXCEL_parser - программа для конвертации данных из Excel в CSV, а также для создания и заполнения баз данных на основе этих данных.\n\n" +
"Функциональность программы:\n\n" +
"1. Конвертация данных из Excel в CSV-файлы.\n" +
"   - Выбор файла Excel: Нажмите кнопку 'Выбрать файл Excel' и выберите Excel-файл, содержащий данные о шкафах и механизмах.\n" +
"   - Выбор папки для сохранения CSV-файлов: Нажмите кнопку 'Выбрать папку' и укажите папку, в которую будут сохранены сгенерированные CSV-файлы.\n" +
"   - Конвертация: Нажмите кнопку 'Конвертировать', чтобы начать процесс конвертации данных из Excel в CSV-файлы. Программа обработает файл Excel и создаст CSV-файлы в выбранной папке.\n\n" +
"2. Создание базы данных.\n" +
"   - Создать базу данных: Нажмите кнопку 'Создать базу данных', чтобы создать новую базу данных 'MyDatabase.db' в выбранной папке. База данных будет заполнена данными, извлеченными из файла Excel.\n" +
"   - Структура базы данных: База данных содержит таблицы messages_binary, messages_integer, messages_analog и другие, которые хранят информацию о точках данных (tags), извлеченных из Excel-файла.\n\n" +
"3. Создание второй базы данных.\n" +
"   - Создать вторую базу данных: Нажмите кнопку 'Создать вторую базу данных', чтобы создать дополнительную базу данных 'SecondDatabase.db'. Эта база данных используется для маршрутизации данных и содержит таблицы managers, routing и другие.\n" +
"   - Заполнение второй базы данных: Вторая база данных наполняется на основе данных из первой базы данных и добавляет информацию о маршрутизации точек данных между различными менеджерами.\n\n" +
"4. Создание базы данных с дублированием точек данных.\n" +
"   - Создать базу данных с дублированием: Нажмите кнопку 'Создать базу данных с дублированием', чтобы создать базу данных 'DuplicatedDatabase.db', в которой точки данных дублируются для System00_1 и System00_2.\n" +
"   - Дублирование точек данных: Программа создает копии точек данных для двух систем, что может быть полезно для определенных задач мониторинга или управления.\n\n" +
"5. Парсинг DPL-файла и добавление данных в базу данных.\n" +
"   - Парсить DPL: Нажмите кнопку 'Парсить DPL', чтобы выбрать файл DPL, который будет проанализирован. Программа извлечет соответствующие точки данных из файла DPL и добавит их в базу данных.\n" +
"   - Добавление данных в базу: Извлеченные точки данных будут добавлены в базу данных, что позволит их дальнейшее использование в системе.\n\n" +
"Инструкция по использованию программы:\n\n" +
"1. Конвертация данных из Excel в CSV:\n" +
"   - Нажмите 'Выбрать файл Excel' и выберите файл Excel с данными (убедитесь, что в файле присутствуют листы 'Мех-мы' и 'Шкафы').\n" +
"   - Нажмите 'Выбрать папку' и укажите папку для сохранения CSV-файлов.\n" +
"   - Нажмите 'Конвертировать', чтобы начать процесс конвертации.\n\n" +
"2. Создание базы данных:\n" +
"   - Убедитесь, что выбран файл Excel и папка для вывода.\n" +
"   - Нажмите кнопку 'Создать базу данных'.\n" +
"   - Подтвердите создание базы данных в появившемся окне.\n\n" +
"3. Создание второй базы данных:\n" +
"   - После создания первой базы данных нажмите 'Создать вторую базу данных'.\n" +
"   - Подтвердите создание второй базы данных.\n\n" +
"4. Создание базы данных с дублированием точек данных:\n" +
"   - Нажмите 'Создать базу данных с дублированием'.\n" +
"   - Подтвердите создание базы данных с дублированием точек данных.\n\n" +
"5. Парсинг DPL-файла и добавление данных в базу данных:\n" +
"   - Нажмите 'Парсить DPL' и выберите DPL-файл.\n" +
"   - Выберите базу данных, в которую необходимо добавить данные.\n" +
"   - Подтвердите добавление данных в базу.\n\n" +
"Примечания:\n\n" +
"- Структура Excel-файла: Убедитесь, что ваш Excel-файл содержит необходимые листы и столбцы:\n" +
"  - Лист 'Мех-мы' должен содержать информацию о механизмах.\n" +
"  - Лист 'Шкафы' должен содержать информацию о шкафах, включая KKS и IP-адреса.\n\n" +
"- Кнопки управления:\n" +
"  - 'Выбрать файл Excel': позволяет выбрать исходный файл Excel.\n" +
"  - 'Выбрать папку': позволяет выбрать папку для сохранения результатов.\n" +
"  - 'Конвертировать': запускает процесс конвертации Excel в CSV.\n" +
"  - 'Создать базу данных': создает базу данных на основе данных из Excel.\n" +
"  - 'Создать вторую базу данных': создает дополнительную базу данных для маршрутизации.\n" +
"  - 'Создать базу данных с дублированием': создает базу данных с дублированием точек данных.\n" +
"  - 'Парсить DPL': позволяет выбрать DPL-файл для парсинга и добавления данных в базу.\n\n" +
"- Работа с базами данных:\n" +
"  - При создании баз данных убедитесь, что файлы базы данных не заняты другими приложениями.\n" +
"  - Если база данных уже существует, программа предложит перезаписать ее.\n\n" +
"- Дополнительная информация:\n" +
"  - Программа автоматически определяет модель механизма на основе кодов из Excel-файла.\n" +
"  - Поддерживаются модели: AKKUYU_FAN, AKKUYU_GATE, AKKUYU_VALVE, AKKUYU_MOTOR.\n" +
"  - В случае отсутствия соответствия используется модель AKKUYU_NO_MODEL.\n\n" +
"- Обработка ошибок:\n" +
"  - Если произойдет ошибка, программа выведет сообщение с описанием проблемы.\n" +
"  - Убедитесь, что все файлы доступны и не используются другими программами.\n\n" +
"- Контактная информация:\n" +
"  - При возникновении вопросов или предложений обращайтесь к разработчикам программы.\n\n" +
"Спасибо за использование EXCEL_parser!";

            MessageBox.Show(infoMessage, "Информация о программе", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Обработчик кнопки для создания базы данных
        private async void button5_Click(object sender, EventArgs e)
        {
            string databasePath = Path.Combine(txtOutputDirectory.Text, "MyDatabase.db");

            var result = MessageBox.Show("Вы уверены, что хотите создать новую базу данных?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }

            if (File.Exists(databasePath))
            {
                var overwriteResult = MessageBox.Show("База данных уже существует. Перезаписать?", "Подтверждение", MessageBoxButtons.YesNo);
                if (overwriteResult == DialogResult.No)
                {
                    return;
                }

                File.Delete(databasePath);
            }

            try
            {
                button5.Text = "Создание базы данных...";
                button5.Enabled = false;

                var cabinetsInfo = GetCabinetsInfo(txtExcelFilePath.Text);
                int totalPoints = 0;

                foreach (var cabinetInfo in cabinetsInfo)
                {
                    var lines = cabinetInfo.Split(',');
                    List<string> kkss = lines.Skip(2).ToList();
                    string cabinetName = lines[0];
                    string controllerIp = lines[1];

                    totalPoints += kkss.Count;

                    await CreateAndFillDatabaseAsync(databasePath, kkss, cabinetName, controllerIp);
                }

                var fileInfo = new FileInfo(databasePath);
                long fileSize = fileInfo.Length;

                int binaryCount = await GetTableCountAsync(databasePath, "messages_binary");
                int integerCount = await GetTableCountAsync(databasePath, "messages_integer");
                int analogCount = await GetTableCountAsync(databasePath, "messages_analog");

                button5.Text = "Создать базу данных";
                button5.Enabled = true;

                MessageBox.Show($"База данных успешно создана!\nФайл: {databasePath}\nРазмер: {fileSize} байт\nКоличество точек данных:\n - messages_binary: {binaryCount}\n - messages_integer: {integerCount}\n - messages_analog: {analogCount}\nОбщее количество точек данных: {totalPoints}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                button5.Text = "Создать базу данных";
                button5.Enabled = true;
                MessageBox.Show($"Ошибка при создании базы данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод для получения количества записей в таблице
        private async Task<int> GetTableCountAsync(string dbPath, string tableName)
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

        // Метод для создания и заполнения базы данных
        private async Task CreateAndFillDatabaseAsync(string dbPath, List<string> KKSIds, string cabinetName, string controllerIp)
        {
            using (var connection = new SQLiteConnection($"Data Source={dbPath}; Version=3;"))
            {
                await connection.OpenAsync();

                string sql = @"
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
                INSERT INTO types (type_id, type_name, comment) VALUES 
                    (0, 'Ana_VT', 'Структура данных аналогового значения. AData (quality : 4 bytes; value : 4 bytes)'), 
                    (1, 'Bin_VT', 'Структура данных двоичного значения. BData (quality : 4 bytes; value : 4 bytes)'), 
                    (2, 'Int_VT', 'Структура данных целочисленного значения. IData (quality : 4 bytes; value : 4 bytes)'), 
                    (3, 'Grp_VT', 'Структура данных группового значения. GData (size : 4 bytes; groupType : 4 bytes; value : 4100 bytes max)');
                COMMIT;";

                using (var command = new SQLiteCommand(sql, connection))
                {
                    await command.ExecuteNonQueryAsync();
                }

                var binaryEntries = new List<string>();
                var integerEntries = new List<string>();
                var existingIdx = new HashSet<string>();

                foreach (var name in KKSIds)
                {
                    string model = GetObjectModel(name.Trim());
                    if (model != "AKKUYU_NO_MODEL")
                    {
                        string[] parts = name.Split(' ');
                        string kksname = parts[1];
                        string kksnumber = parts[2];

                        var tags = new List<Tuple<string, string>>()
                        {
                            new Tuple<string, string>("Alarm_bit0", "bool"),
                            new Tuple<string, string>("Alarm_word_1", "int"),
                            new Tuple<string, string>("Mode", "int"),
                            new Tuple<string, string>("OPRT", "int"),
                            new Tuple<string, string>("OPRT_INDEX", "int"),
                            new Tuple<string, string>("State", "int"),
                            new Tuple<string, string>("RASU_bits0", "bool"),
                            new Tuple<string, string>("RASU_bits1", "bool")
                        };

                        foreach (var tag in tags)
                        {
                            string baseIdx = $"{cabinetName}_{kksname}_{tag.Item1}";
                            string idx = baseIdx;
                            string wcc_oa_tag = $"System00_1:ManagementView_FieldNetworks_S7Plus_{cabinetName}_{kksnumber}.{tag.Item1}";
                            string comment = $"{tag.Item1} {cabinetName}_{kksname}";



                            int counter = 1;
                            while (existingIdx.Contains(idx) || await IdxExistsAsync(connection, idx))
                            {
                                idx = $"{baseIdx}_{counter}";
                                counter++;
                            }

                            existingIdx.Add(idx);

                            string entry = $"('{idx}', '{wcc_oa_tag}', '{comment}')";

                            if (tag.Item2 == "bool")
                            {
                                binaryEntries.Add(entry);
                            }
                            else
                            {
                                integerEntries.Add(entry);
                            }
                        }
                    }
                }

                if (binaryEntries.Count > 0)
                {
                    string binarySql = $"INSERT INTO messages_binary (idx, wcc_oa_tag, comment) VALUES {string.Join(", ", binaryEntries)};";
                    using (var command = new SQLiteCommand(binarySql, connection))
                    {
                        try
                        {
                            await command.ExecuteNonQueryAsync();
                        }
                        catch (SQLiteException ex)
                        {
                            MessageBox.Show($"Ошибка при вставке данных в messages_binary: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }

                if (integerEntries.Count > 0)
                {
                    string integerSql = $"INSERT INTO messages_integer (idx, wcc_oa_tag, comment) VALUES {string.Join(", ", integerEntries)};";
                    using (var command = new SQLiteCommand(integerSql, connection))
                    {
                        try
                        {
                            await command.ExecuteNonQueryAsync();
                        }
                        catch (SQLiteException ex)
                        {
                            MessageBox.Show($"Ошибка при вставке данных в messages_integer: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        // Метод для проверки существования idx в базе данных
        private async Task<bool> IdxExistsAsync(SQLiteConnection connection, string idx)
        {
            using (var command = new SQLiteCommand("SELECT COUNT(*) FROM (SELECT idx FROM messages_analog WHERE idx = @idx UNION ALL SELECT idx FROM messages_binary WHERE idx = @idx UNION ALL SELECT idx FROM messages_integer WHERE idx = @idx)", connection))
            {
                command.Parameters.AddWithValue("@idx", idx);
                return Convert.ToInt32(await command.ExecuteScalarAsync()) > 0;
            }
        }

        private void txtOutputDirectory_TextChanged(object sender, EventArgs e)
        {

        }

        // Обработчик кнопки для создания второй базы данных
        private async void btnCreateSecondDatabase_Click(object sender, EventArgs e)
        {
            string secondDatabasePath = Path.Combine(txtOutputDirectory.Text, "SecondDatabase.db");

            var result = MessageBox.Show("Вы уверены, что хотите создать новую базу данных?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }

            if (File.Exists(secondDatabasePath))
            {
                var overwriteResult = MessageBox.Show("База данных уже существует. Перезаписать?", "Подтверждение", MessageBoxButtons.YesNo);
                if (overwriteResult == DialogResult.No)
                {
                    return;
                }

                File.Delete(secondDatabasePath);
            }

            try
            {
                btnCreateSecondDatabase.Text = "Создание базы данных...";
                btnCreateSecondDatabase.Enabled = false;

                await CreateAndFillSecondDatabaseAsync(secondDatabasePath);

                btnCreateSecondDatabase.Text = "Создать вторую базу данных";
                btnCreateSecondDatabase.Enabled = true;

                MessageBox.Show($"Вторая база данных успешно создана!\nФайл: {secondDatabasePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnCreateSecondDatabase.Text = "Создать вторую базу данных";
                btnCreateSecondDatabase.Enabled = true;
                MessageBox.Show($"Ошибка при создании второй базы данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод для создания и заполнения второй базы данных
        private async Task CreateAndFillSecondDatabaseAsync(string dbPath)
        {
            using (var connection = new SQLiteConnection($"Data Source={dbPath}; Version=3;"))
            {
                await connection.OpenAsync();

                string sql = @"
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
                INSERT INTO types (id, name, comment) VALUES 
                    (0, 'ANA_VT', 'Аналоговый тип.'), 
                    (1, 'BIN_VT', 'Булевый тип'), 
                    (2, 'INT_VT', 'Целочисленный тип');
                INSERT INTO managers (id, name, sending_socket, receiving_socket, description) VALUES 
                    (1, 'DtsManager', 'ipc:///tmp/S_dts', 'ipc:///tmp/R_dts', 'Менеджер канала DTS.'),
                    (2, 'WccManager', 'ipc:///tmp/S_wcc', 'ipc:///tmp/R_wcc', 'Менеджер взаимодействия с WinCC OA. '),
                    (3, 'WCCOA_Mg_Gateway', 'tcp://192.168.11.151:5555', NULL, 'Менеджер взаимодействия с Desigo CC');
                CREATE VIEW ROUTING_VIEW AS 
                    SELECT MAN_FROM.name as from_manager, route.from_idx as from_idx, MAN_TO.name as to_manager, route.to_idx as to_idx 
                    FROM routing as route 
                    LEFT JOIN managers as MAN_FROM on route.from_manager_id = MAN_FROM.id 
                    LEFT JOIN managers as MAN_TO on route.to_manager_id = MAN_TO.id;
                COMMIT;";

                using (var command = new SQLiteCommand(sql, connection))
                {
                    await command.ExecuteNonQueryAsync();
                }

                var routingEntries = new List<string>();
                int id = 1;
                int fromManagerId = 3;
                int toManagerId = 1;

                var allIdx = await GetAllIdxFromFirstDatabaseAsync();

                foreach (var item in allIdx)
                {
                    string fromIdx = item.Item1;
                    int tagType = item.Item2;
                    string toIdx = $"{tagType}{item.Item3}";

                    routingEntries.Add($"({id}, {fromManagerId}, '{fromIdx}', {toManagerId}, '{toIdx}')");
                    id++;
                }

                if (routingEntries.Count > 0)
                {
                    string routingSql = $"INSERT INTO routing (id, from_manager_id, from_idx, to_manager_id, to_idx) VALUES {string.Join(", ", routingEntries)};";
                    using (var command = new SQLiteCommand(routingSql, connection))
                    {
                        try
                        {
                            await command.ExecuteNonQueryAsync();
                        }
                        catch (SQLiteException ex)
                        {
                            MessageBox.Show($"Ошибка при вставке данных в routing: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        // Метод для получения всех idx из первой базы данных
        private async Task<List<Tuple<string, int, int>>> GetAllIdxFromFirstDatabaseAsync()
        {
            var allIdx = new List<Tuple<string, int, int>>();

            using (var connection = new SQLiteConnection($"Data Source={Path.Combine(txtOutputDirectory.Text, "MyDatabase.db")}; Version=3;"))
            {
                await connection.OpenAsync();

                string sql = @"
                SELECT idx, 0 AS type, id 
                FROM messages_analog 
                UNION ALL 
                SELECT idx, 1 AS type, id 
                FROM messages_binary 
                UNION ALL 
                SELECT idx, 2 AS type, id 
                FROM messages_integer";

                using (var command = new SQLiteCommand(sql, connection))
                {
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
            }

            return allIdx;
        }

        // Обработчик кнопки для создания базы данных с дублированием точек данных для System00_1 и System00_2
        private async void btnCreateDuplicatedDatabase_Click(object sender, EventArgs e)
        {
            string duplicatedDatabasePath = Path.Combine(txtOutputDirectory.Text, "DuplicatedDatabase.db");

            var result = MessageBox.Show("Вы уверены, что хотите создать новую базу данных с дублированными точками данных?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                return;
            }

            if (File.Exists(duplicatedDatabasePath))
            {
                var overwriteResult = MessageBox.Show("База данных уже существует. Перезаписать?", "Подтверждение", MessageBoxButtons.YesNo);
                if (overwriteResult == DialogResult.No)
                {
                    return;
                }

                File.Delete(duplicatedDatabasePath);
            }

            try
            {
                btnCreateDuplicatedDatabase.Text = "Создание базы данных...";
                btnCreateDuplicatedDatabase.Enabled = false;

                await CreateAndFillSystemDatabasesAsync(duplicatedDatabasePath);

                btnCreateDuplicatedDatabase.Text = "Создать базу данных";
                btnCreateDuplicatedDatabase.Enabled = true;

                MessageBox.Show($"База данных успешно создана!\nФайл: {duplicatedDatabasePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnCreateDuplicatedDatabase.Text = "Создать базу данных";
                btnCreateDuplicatedDatabase.Enabled = true;
                MessageBox.Show($"Ошибка при создании базы данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Генератор баз для четырёх «логических» систем WCC OA.  
        /// <para>‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑‑</para>
        /// ▸ Вместо одной «общей» БД теперь формируются **четыре**
        ///   отдельные БД‑файла:  
        ///   • System00_1.db  
        ///   • System00_2.db  
        ///   • System10_1.db  
        ///   • System10_2.db  
        /// ▸ В каждом файле содержатся ТОЛЬКО точки своей системы  
        /// ▸ Структура таблиц и вся остальная «идеология» сохранена  
        /// ▸ Метод рассчитан на C# 7.3 (без records, init‑set и т.п.)  
        /// </summary>
        /// <param name="dirPath">
        ///   Папка, куда будут сложены БД.  
        ///   Если её нет – будет создана.
        /// </param>
        private async Task CreateAndFillSystemDatabasesAsync(string dirPath)
        {
            /* -----------------------------------------------------------
               0.  Подготовка: имена систем и будущих БД‑файлов
            ----------------------------------------------------------- */
            string[] systems = { "System00_1", "System00_2", "System10_1", "System10_2" };

            // Гарантируем существование каталога
            if (!System.IO.Directory.Exists(dirPath))
                System.IO.Directory.CreateDirectory(dirPath);

            /* -----------------------------------------------------------
               1.  Читаем Excel → получаем перечень шкафов и ККС‑объектов,
                   это общая операция, делаем один раз
            ----------------------------------------------------------- */
            var cabinetsInfo = GetCabinetsInfo(txtExcelFilePath.Text); // ← ваша существующая функция

            /* -----------------------------------------------------------
               2.  Цикл по системам: для каждой создаём СВОЙ БД‑файл
                   и наполняем только «своими» точками
            ----------------------------------------------------------- */
            foreach (string sys in systems)
            {
                string dbFile = System.IO.Path.Combine(dirPath, sys + ".db");

                // ───────── 2.1  Открываем (или создаём) БД
                using (var connection = new SQLiteConnection($"Data Source={dbFile}; Version=3;"))
                {
                    await connection.OpenAsync();

                    /* ---------- 2.1.1  Создание таблиц (если их ещё нет) */
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
                    using (var cmd = new SQLiteCommand(schemaSql, connection))
                        await cmd.ExecuteNonQueryAsync();

                    /* ---------- 2.1.2  Подготовка коллекций для bulk‑insert */
                    var binValues = new List<string>();           // bool‑точки
                    var intValues = new List<string>();           // int‑точки
                    var idxSet = new HashSet<string>();        // контроль уникальности в рамках ЭТОЙ БД

                    /* -----------------------------------------------------
                       2.2  Проход по всем шкафам → по всем ККС → по тегам
                    ----------------------------------------------------- */
                    foreach (var cabinetInfo in cabinetsInfo)
                    {
                        var parts = cabinetInfo.Split(',');
                        var cabinet = parts[0];
                        var ctrlIp = parts[1];                // не используется, но оставим
                        var kkss = parts.Skip(2).ToList();  // все ККС данного шкафа

                        foreach (string raw in kkss)
                        {
                            string model = GetObjectModel(raw.Trim());
                            if (model == "AKKUYU_NO_MODEL") continue; // пропускаем «пустышки»

                            string[] k = raw.Split(' ');
                            string kksName = k[1];
                            string kksNum = k[2];
                            string kksBuild = k[3];

                            /* стандартный набор тегов: имя + тип */
                            var tags = new List<Tuple<string, string>>
                    {
                        Tuple.Create("Alarm_bit0" , "bool"),
                        Tuple.Create("Alarm_word_1","int"),
                        Tuple.Create("Mode"       , "int"),
                        Tuple.Create("OPRT"       , "int"),
                        Tuple.Create("OPRT_INDEX" , "int"),
                        Tuple.Create("State"      , "int"),
                        Tuple.Create("RASU_bits0" , "bool"),
                        Tuple.Create("RASU_bits1" , "bool")
                    };

                            foreach (var tag in tags)
                            {
                                /* ---------- формируем IDX и WCCOA‑tag */
                                string baseIdx = $"{cabinet}_{kksName}_{tag.Item1}";
                                string idx = baseIdx;

                                // контроль уникальности IDX внутри одной БД
                                int counter = 1;
                                while (idxSet.Contains(idx) || await IdxExistsAsync(connection, idx))
                                {
                                    idx = $"{baseIdx}_{counter}";
                                    counter++;
                                }
                                idxSet.Add(idx);

                                string wccTag =
                                    $"{sys}:ManagementView_FieldNetworks_S7Plus_{kksBuild}_{cabinet}_{kksNum}.{tag.Item1}";
                                string comment = $"{tag.Item1} {kksBuild}_{cabinet}_{kksName}";

                                string sqlValue = $"('{idx}','{wccTag}','{comment}')";
                                if (tag.Item2 == "bool") binValues.Add(sqlValue);
                                else intValues.Add(sqlValue);
                            }
                        }
                    }

                    /* ---------- 2.3  Bulk‑insert  (если есть что вставлять) */
                    if (binValues.Count > 0)
                    {
                        string sql = $"INSERT INTO messages_binary (idx,wcc_oa_tag,comment) VALUES {string.Join(",", binValues)};";
                        using (var cmd = new SQLiteCommand(sql, connection))
                            await cmd.ExecuteNonQueryAsync();
                    }
                    if (intValues.Count > 0)
                    {
                        string sql = $"INSERT INTO messages_integer (idx,wcc_oa_tag,comment) VALUES {string.Join(",", intValues)};";
                        using (var cmd = new SQLiteCommand(sql, connection))
                            await cmd.ExecuteNonQueryAsync();
                    }

                    /* ---------- 2.4  Информационное сообщение в Debug‑лог */
                    Debug.WriteLine($"[{sys}] DB готова: {binValues.Count} bool + {intValues.Count} int");
                } // using connection
            }     // foreach system
        }








        // Обработчик кнопки для парсинга DPL и добавления данных в базу
        private void btnParseDpl_Click(object sender, EventArgs e)
        {
            string dplFilePath = GetDplFilePath();
            if (string.IsNullOrEmpty(dplFilePath) || !File.Exists(dplFilePath))
            {
                MessageBox.Show($"Ошибка: Входной файл '{dplFilePath}' не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string databasePath = GetDatabaseFilePath();
            if (string.IsNullOrEmpty(databasePath) || !File.Exists(databasePath))
            {
                MessageBox.Show($"Ошибка: Файл базы данных '{databasePath}' не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                List<string> matchingLines = GetMatchingLinesFromDplFile(dplFilePath);
                if (matchingLines.Count == 0)
                {
                    MessageBox.Show("Не найдено строк, соответствующих заданным паттернам.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                List<(string dataPoint, string comment, string cabinetCode)> matchingData = matchingLines.Select(ExtractData).ToList();
                AddDataPointsToDatabase(databasePath, matchingData);
                MessageBox.Show("Данные успешно добавлены в базу данных.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении данных в базу: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Метод для получения пути к файлу DPL от пользователя
        private string GetDplFilePath()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "DPL Files|*.dpl";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return string.Empty;
        }

        // Метод для получения пути к базе данных от пользователя
        private string GetDatabaseFilePath()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Database Files|*.db";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
            }
            return string.Empty;
        }

        // Метод для получения строк из файла DPL, содержащих заданные паттерны и ККС после LANG:10027
        private List<string> GetMatchingLinesFromDplFile(string filePath)
        {
            List<string> patterns = new List<string>
            {
                "Шлюз_ Пожар",
                "Шлюз_ Неисправность",
                "Шлюз_ Дверь открыта"
            };

            List<string> matchingLines = new List<string>();

            string[] lines = File.ReadAllLines(filePath);
            foreach (string line in lines)
            {
                List<string> tokens = SplitLine(line);
                int langIndex = tokens.IndexOf("LANG:10027");

                if (langIndex != -1 && langIndex + 1 < tokens.Count)
                {
                    string comment = tokens[langIndex + 1].Trim('"');
                    foreach (string pattern in patterns)
                    {
                        if (comment.StartsWith(pattern))
                        {
                            string[] commentParts = comment.Split(new char[] { ' ', '_' }, StringSplitOptions.RemoveEmptyEntries);
                            string cabinetCode = commentParts.LastOrDefault();

                            if (!string.IsNullOrEmpty(cabinetCode) && IsCabinetCode(cabinetCode))
                            {
                                matchingLines.Add(line);
                                break;
                            }
                        }
                    }
                }
            }

            return matchingLines;
        }

        // Метод для добавления точек данных в базу данных
        private void AddDataPointsToDatabase(string databasePath, List<(string dataPoint, string comment, string cabinetCode)> matchingData)
        {
            using (var connection = new SQLiteConnection($"Data Source={databasePath}; Version=3;"))
            {
                connection.Open();

                foreach (var data in matchingData)
                {
                    string combinedComment = $"{data.comment} {data.cabinetCode}";
                    string idxSystem001 = $"{data.dataPoint}_System001";
                    string idxSystem002 = $"{data.dataPoint}_System002";

                    idxSystem001 = GetUniqueIdx(connection, idxSystem001);
                    idxSystem002 = GetUniqueIdx(connection, idxSystem002);

                    string sql1 = $"INSERT INTO messages_integer (idx, wcc_oa_tag, comment) VALUES ('{idxSystem001}', 'System00_1:{data.dataPoint}.Present_Value', '{combinedComment}');";
                    string sql2 = $"INSERT INTO messages_integer (idx, wcc_oa_tag, comment) VALUES ('{idxSystem002}', 'System00_2:{data.dataPoint}.Present_Value', '{combinedComment}');";

                    InsertDataPoint(connection, sql1);
                    InsertDataPoint(connection, sql2);
                }
            }
        }

        // Метод для получения уникального значения idx
        private string GetUniqueIdx(SQLiteConnection connection, string baseIdx)
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

        // Метод для проверки существования idx в базе данных
        private bool IdxExists(SQLiteConnection connection, string idx)
        {
            using (var command = new SQLiteCommand("SELECT COUNT(*) FROM messages_integer WHERE idx = @idx", connection))
            {
                command.Parameters.AddWithValue("@idx", idx);
                return Convert.ToInt32(command.ExecuteScalar()) > 0;
            }
        }


        // Метод для разбиения строки на токены, учитывая кавычки
        static List<string> SplitLine(string line)
        {
            List<string> tokens = new List<string>();
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
            {
                tokens.Add(currentToken);
            }

            return tokens;
        }

        // Метод для проверки, является ли строка кодом ККС
        static bool IsCabinetCode(string code)
        {
            return code.Length == 7 && code.All(c => char.IsUpper(c) || char.IsDigit(c));
        }

        // Метод для извлечения точки данных, комментария и кода ККС из строки
        static (string dataPoint, string comment, string cabinetCode) ExtractData(string line)
        {
            string dataPoint = string.Empty;
            string comment = string.Empty;
            string cabinetCode = string.Empty;

            try
            {
                List<string> tokens = SplitLine(line);
                dataPoint = tokens.FirstOrDefault(t => t.StartsWith("GmsDevice"));

                for (int i = 0; i < tokens.Count; i++)
                {
                    if (tokens[i] == "LANG:10027" && i + 1 < tokens.Count)
                    {
                        comment = tokens[i + 1].Trim('"');
                        string[] commentParts = comment.Split(new char[] { ' ', '_' }, StringSplitOptions.RemoveEmptyEntries);
                        cabinetCode = commentParts.LastOrDefault();

                        if (!string.IsNullOrEmpty(cabinetCode) && IsCabinetCode(cabinetCode))
                        {
                            int index = comment.LastIndexOf(cabinetCode);
                            if (index > 0)
                            {
                                comment = comment.Substring(0, index).Trim(new char[] { ' ', '_' });
                            }
                        }
                        else
                        {
                            cabinetCode = "Неизвестно";
                        }
                        break;
                    }
                }
            }
            catch (Exception)
            {
                dataPoint = string.Empty;
                comment = string.Empty;
                cabinetCode = string.Empty;
            }

            return (dataPoint, comment, cabinetCode);
        }

        // Метод для вставки точки данных в базу данных
        static void InsertDataPoint(SQLiteConnection connection, string sql)
        {
            using (var command = new SQLiteCommand(sql, connection))
            {
                command.ExecuteNonQuery();
            }
        }


        /////////////////////////////




        // Экспорт: 1 файл = 1 блок (00/10/20/30/40); внутри — несколько DEVICES (по шкафам) + POINTS (по их механизмам)
        private async void btnExportByBuilding_Click(object sender, EventArgs e)
        {
            string excelFilePath = txtExcelFilePath.Text;
            string outputDirectory = txtOutputDirectory.Text;

            if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
            {
                MessageBox.Show("Выбери корректный Excel-файл.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(outputDirectory) || !Directory.Exists(outputDirectory))
            {
                MessageBox.Show("Выбери существующую папку для вывода.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show("Создать CSV-файлы по блокам (00/10/20/30/40)?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No) return;

            try
            {
                btnExportByBuilding.Enabled = false;
                btnExportByBuilding.Text = "Экспорт по блокам...";

                // Сырые данные
                var mechData = ReadMechData(excelFilePath, "Мех-мы");     // KKS шкафа -> List("10SGK33AA007 18 10UBA", ...)
                var cabinetData = ReadCabinetData(excelFilePath, "Шкафы");   // KKS шкафа -> (IP, Building, Type, VEKSH)

                // Группировка шкафов по БЛОКУ
                var byBlock = new Dictionary<string, List<string>>(); // "00"/"10"/... -> List<KKS шкафа>
                foreach (var kv in cabinetData)
                {
                    string cabKks = kv.Key;                 // KKS шкафа (например, "10CMM11")
                    string building = kv.Value.Building;      // например, "10UBA"
                    string block = GetBlockFromBuildingOrKks(building, cabKks);

                    if (block == null) continue;              // пропускаем неизвестные
                    if (!byBlock.ContainsKey(block)) byBlock[block] = new List<string>();
                    byBlock[block].Add(cabKks);
                }

                // Оставляем только существующие блоки в порядке 00..40
                string[] order = { "00", "10", "20", "30", "40" };
                var blocksOrdered = order.Where(b => byBlock.ContainsKey(b)).ToList();

                int filesCreated = 0;
                int blocksTotal = blocksOrdered.Count;
                progressBar1.Value = 0;
                progressBar1.Maximum = Math.Max(1, blocksTotal);

                foreach (var block in blocksOrdered)
                {
                    List<string> cabinetKks = byBlock[block];
                    if (cabinetKks == null || cabinetKks.Count == 0)
                    {
                        progressBar1.Value++;
                        continue;
                    }

                    var sb = new StringBuilder();
                    AppendStandardHeader(sb);

                    // DEVICES
                    AppendDevicesHeader(sb);
                    foreach (var cabKks in cabinetKks)
                    {
                        // Если у шкафа нет механизмов — пропускаем девайс
                        if (!mechData.TryGetValue(cabKks, out var kkss) || kkss == null || kkss.Count == 0)
                            continue;

                        var cab = cabinetData[cabKks];
                        string building = cab.Building;                  // ВНИМАНИЕ: используем ФАКТИЧЕСКОЕ "Здание" шкафа (а не "блок")
                        string cabinetNm = CleanString(cabKks);
                        string ip = CleanString(cab.IP);
                        string type = CleanString(cab.Type);

                        //string desc = GetDistinctBuildingsWithExclusion(kkss, building) + cabinetNm + type;
                        //string desc = BuildDeviceDescription(GetDistinctBuildingsWithExclusion(kkss, building),cabinetNm,type);
                        string desc = BuildDeviceDescription(GetDistinctBuildingsForBlockExport(kkss, building), cabinetNm, type);



                        sb.AppendLine($"{building + "_" + cabinetNm},{desc},S7-1500,{ip},S7ONLINE,S7Plus$Online,Online,TRUE,PLC_{building + "_" + cabinetNm},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
                    }
                    sb.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");

                    // POINTS — шапка
                    AppendPointsHeader(sb);

                    // Диагностика + точки по механизмам
                    foreach (var cabKks in cabinetKks)
                    {
                        if (!mechData.TryGetValue(cabKks, out var kkss) || kkss == null || kkss.Count == 0)
                            continue;

                        var cab = cabinetData[cabKks];
                        string building = cab.Building;       // ФАКТИЧЕСКОЕ "Здание" шкафа
                        string cabinetNm = CleanString(cabKks);
                        string type = CleanString(cab.Type);

                        // Диагностика шкафа
                        AppendCabinetDiagnostics(sb, building, cabinetNm, type);

                        // Две строки «зон»
                        sb.AppendLine();
                        sb.AppendLine($"{building + "_" + cabinetNm},Zones_string_1,Помещение_ПожарЗАМЕНА,FB_Zones_DB.Zones_string_1,String,IO,FALSE,COV,pollGr_3,,,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
                        sb.AppendLine($"{building + "_" + cabinetNm},Zones_string_2,Помещение_Пожар_ПодтверждениеЗАМЕНА,FB_Zones_DB.Zones_string_2,String,IO,FALSE,COV,pollGr_3,,,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
                        sb.AppendLine();

                        // Механизмы
                        foreach (string raw in kkss)
                        {
                            // raw: "10SGK33AA007 18 10UBA"
                            var parts = raw.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                            string mechKks = parts.ElementAtOrDefault(0) ?? "";
                            string kksNumber = parts.ElementAtOrDefault(1) ?? "0";
                            string mechBld = parts.ElementAtOrDefault(2) ?? building;

                            string model = GetObjectModel(mechKks);
                            if (model == "AKKUYU_NO_MODEL") continue;

                            string dispName = GetBuildingDifference(mechBld, building) + mechKks;

                            // KKS
                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},{dispName},DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.KKS,String,IO,FALSE,COV,pollGr_3,{model},KKS,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\\{building}\\Mechanisms\\{building + "_" + cabinetNm}\\,");

                            // Базовый набор
                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].OPRT.OPRT_INDEX,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT_INDEX,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.Mode,Uint,IO,FALSE,COV,pollGr_3,{model},Mode,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.State,Uint,IO,FALSE,COV,pollGr_3,{model},State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

                            // Доп. поле для клапанов
                            if (model.Contains("AKKUYU_VALVE"))
                                sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.Extinguishing_State,Uint,IO,FALSE,COV,pollGr_3,{model},Extinguishing_State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].ALM.Alarm_word_1,Uint,IO,FALSE,COV,pollGr_3,{model},Alarm_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].OPRT.OPRT,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].ALM.Alarm.Alarm[0],bool,IO,FALSE,COV,pollGr_3,{model},Alarm_bit0,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.RASU_bits.RASU_bits[0],bool,IO,FALSE,COV,pollGr_3,{model},RASU_bits0,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                            sb.AppendLine($"{building + "_" + cabinetNm},{kksNumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksNumber}].PRM.RASU_bits.RASU_bits[1],bool,IO,FALSE,COV,pollGr_3,{model},RASU_bits1,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                            sb.AppendLine(" ");
                        }
                    }

                    // Хвост
                    sb.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
                    sb.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
                    sb.AppendLine(@"[VALUE_ASSIGNMENT],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
                    sb.AppendLine(@"#[ObjectPath],[Property Name],[Property Value],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");

                    // Имя файла: <Блок>_<метка>_devices.csv, например "10_первый_блок_devices.csv"
                    string label = GetBlockLabel(block);
                    string safeLabel = label.Replace(' ', '_');
                    string fileName = $"{block}_{safeLabel}_devices.csv";
                    string path = Path.Combine(outputDirectory, fileName);

                    using (var writer = new StreamWriter(path, false, new UTF8Encoding(true)))
                        await writer.WriteAsync(sb.ToString());

                    filesCreated++;
                    progressBar1.Value++;
                }

                btnExportByBuilding.Enabled = true;
                btnExportByBuilding.Text = "CSV по блокам";

                MessageBox.Show($"Готово. Создано файлов: {filesCreated} (по {blocksTotal} блокам).",
                    "Экспорт по блокам", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnExportByBuilding.Enabled = true;
                btnExportByBuilding.Text = "CSV по блокам";
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // Склеивает части с одинарным пробелом и без двойных пробелов
        private static string BuildDeviceDescription(string distinct, string cabinet, string type)
        {
            string res = string.Empty;
            if (!string.IsNullOrWhiteSpace(distinct)) res = distinct.Trim();
            if (!string.IsNullOrWhiteSpace(cabinet)) res += (res.Length > 0 ? " " : "") + cabinet.Trim();
            if (!string.IsNullOrWhiteSpace(type)) res += (res.Length > 0 ? " " : "") + type.Trim();
            return res;
        }

        private static (string mechKks, string mechNum, string mechBld) ParseMechTriplet(string raw, string fallbackBuilding = "")
        {
            // сохраняем пустые элементы между пробелами
            var parts = raw.Split(new[] { ' ' }, StringSplitOptions.None);
            string mechKks = parts.ElementAtOrDefault(0)?.Trim() ?? "";
            string mechNum = parts.ElementAtOrDefault(1)?.Trim() ?? "";   // может быть пустым — это ОК
            string mechBld = parts.ElementAtOrDefault(2)?.Trim() ?? fallbackBuilding;
            return (mechKks, mechNum, mechBld);
        }

        private static bool IsUint(string s) => !string.IsNullOrWhiteSpace(s) && s.All(char.IsDigit);




        // Возвращает "00","10","20","30","40" либо null, если не распознано
        // Возвращает "00","10","20","30","40" с нормализацией префиксов (0x→00, 1x→10, 2x→20, 3x→30, 4x→40)
        private static string GetBlockFromBuildingOrKks(string building, string cabinetKks)
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

            // 1) Пытаемся по "Здание" (напр., 11UBN → 10)
            var byBuilding = Normalize(building);
            if (byBuilding != null) return byBuilding;

            // 2) Фолбэк — по KKS шкафа (напр., 11CMM05 → 10)
            return Normalize(cabinetKks);
        }


        private static string GetBlockLabel(string block)
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

        /// <param name="csvContent"></param>
        // Общая «шапка» CSV (как в GenerateCsvContent, но один раз на файл)
        private static void AppendStandardHeader(StringBuilder csvContent)
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

        private static void AppendDevicesHeader(StringBuilder csvContent)
        {
            csvContent.AppendLine("[DEVICES],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#[DeviceName],[DeviceDescription],[S7_Type],[IP_address],[AccessPointS7],[Projectname],[StationName],[Establish Connection],[Alias],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],,,,,,,,,,,,,,,,,,,,,,,,,,");
        }

        private static void AppendPointsHeader(StringBuilder csvContent)
        {
            csvContent.AppendLine("[POINTS],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
            csvContent.AppendLine("#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,");
            csvContent.AppendLine("#[ParentDeviceName],[Name],[Description],[Address],[DataType],[Direction],[LowLevelComparison],[PollType],[PollGroup],[ObjectModel],[Property],[Alias],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],[Min],[Max],[MinRaw],[MaxRaw],[MinEng],[MaxEng],[Resolution],[Unit],[UnitTextGroup],[StateText],[ActivityLog],[ValueLog],[Smoothing],[AlarmClass],[AlarmType],[AlarmValue],[EventText],[NormalText],[UpperHysteresis],[LowerHysteresis],[NoAlarmOn],[LogicalHierarchy],[UserHierarchy]");
        }

        // Диагностический «блок шкафа» — как в твоём GenerateCsvContent
        private static void AppendCabinetDiagnostics(StringBuilder sb, string buildingName, string cabinetName, string type)
        {
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,ДиагностикаЗАМЕНА,DB_HMI_IO_Mechanisms.Cabinet.KKS,String,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,KKS,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL1_Power_On,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit0,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL2_Automation_Disabled,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit1,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL3_Malfunction,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit2,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL4_Fire,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit3,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL5_Start,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit4,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HL6_Maintenance,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit5,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Door_1_Opened,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit6,,,,,,,,,,,,,,,,TxG_AKK_CabDoor1,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.State_word_1,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,State_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

            // Если АПТС — добавляем 2 строки
            if (type.Contains("АПТС"))
            {
                sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Door_2_Opened,bool,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,bit7,,,,,,,,,,,,,,,,TxG_AKK_CabDoor2,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
                sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HMI_Alarm_Word_2,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,Alarm_word_2,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
            }

            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.HMI_Alarm_Word_1,Uint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,Alarm_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.Watchdog,Dint,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,ConnectionPLC,,,,,,,,,,,,,,,,,,,SmoothingConfig_01,,,,,,,,DriverFail,,");
            sb.AppendLine($"{buildingName + "_" + cabinetName},Diagnostics,,DB_HMI_IO_Mechanisms.Cabinet.OPRT,int,IO,FALSE,COV,pollGr_3,Diagnostic_Cab,OPRT,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
        }

        // безопасно достаёт список чужих зданий (кроме текущего шкафа)
        private static string GetDistinctBuildingsForBlockExport(List<string> kkss, string deviceBuilding)
        {
            var set = new HashSet<string>();
            foreach (var raw in kkss)
            {
                var (_, _, bld) = ParseMechTriplet(raw, deviceBuilding); // уже учитывает пустой номер
                if (!string.IsNullOrWhiteSpace(bld))
                    set.Add(bld);
            }
            set.Remove(deviceBuilding); // оставляем только "чужие" здания
            return set.Count > 0 ? "(" + string.Join(" ", set) + ")" : string.Empty;
        }









        private void Converter_Load(object sender, EventArgs e)
        {

        }
    }

}


