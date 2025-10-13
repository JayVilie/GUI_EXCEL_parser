# GUI_EXCEL_parser
## Описание проекта

**GUI_EXCEL_parser** — это Windows Forms приложение, предназначенное для конвертации данных из Excel-файлов в CSV-файлы и для создания баз данных SQLite. Программа также поддерживает создание дублированных баз данных с точками данных для System00_1 и System00_2.

## Установка и запуск

### Требования
- .NET Framework 4.7.2 или выше
- Библиотеки ClosedXML и SQLite

### Шаги по установке

1. Склонируйте репозиторий или скачайте архив с кодом.
2. Откройте проект в Visual Studio.
3. Убедитесь, что все зависимости установлены (ClosedXML, SQLite).
4. Постройте и запустите проект.

### Запуск
Для запуска приложения выполните следующие действия:

1. Откройте проект в Visual Studio.
2. Нажмите на кнопку "Start" (F5) для запуска приложения.

## Структура проекта

### Основные файлы и классы

- `Program.cs`: Точка входа в приложение.
- `Converter.cs`: Основной класс формы, который отвечает за пользовательский интерфейс и взаимодействие с пользователем.
- `Converter.Designer.cs`: Автоматически сгенерированный файл для дизайна формы.
- `Converter.resx`: Ресурсы формы (например, строки, иконки).

### Методы и функции

#### Пример: `InitializeComponent()`
Этот метод автоматически генерируется конструктором форм и используется для инициализации компонентов формы.

```csharp
private void InitializeComponent()
{
    System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Converter));
    // Здесь идет код инициализации компонентов...
}
```

## Основные функции

### Выбор файла Excel

#### Метод: `btnSelectExcelFile_Click`

Этот метод открывает диалоговое окно для выбора файла Excel.

```csharp
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
```

### Выбор директории для сохранения CSV файлов

#### Метод: `btnSelectOutputDirectory_Click`

Этот метод открывает диалоговое окно для выбора директории, в которую будут сохранены CSV файлы.

```csharp
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
```

### Конвертация данных из Excel в CSV

#### Метод: `btnConvert_Click`

Этот метод запускает процесс конвертации данных из Excel файла в CSV файлы.

```csharp
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

        btnConvert.Text = "Конвертировать";
        btnConvert.Enabled = true;

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
```

### Получение информации о шкафах и механизмах

#### Метод: `GetCabinetsInfo`

Этот метод извлекает информацию о шкафах и механизмах из Excel файла.

```csharp
static List<string> GetCabinetsInfo(string filePath)
{
    var mechSheetName = "Мех-мы";
    var cabinetsSheetName = "Шкафы";
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
```

### Создание CSV файла

#### Метод: `GenerateCsv`

Этот метод генерирует CSV файл на основе переданных данных.

```csharp
static void GenerateCsv(List<string> lines, string outputDirectory)
{
    string cabinetName = lines[0];
    string controllerIp = lines[1];
    List<string> kkss = lines.Skip(2).ToList();
    StringBuilder csvContent = GenerateCsvContent(kkss, cabinetName, controllerIp);
    string csvFileName = controllerIp.ToLower() == "no ip" ? $"{cabinetName}_без_айпи.csv" : $"{cabinetName}.csv";
    string csvFilePath = Path.Combine(outputDirectory, csvFileName);
    File.WriteAllText(csvFilePath, csvContent.ToString());
    Console.WriteLine($"CSV файл создан: {csvFilePath}");
}
```

### Чтение данных о механизмах

#### Метод: `ReadMechData`

Этот метод читает данные о механизмах из Excel файла.

```csharp
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
```

### Генерация содержимого CSV файла

#### Метод: `GenerateCsvContent`

Этот метод генерирует содержимое CSV файла.

```csharp
static StringBuilder GenerateCsvContent(List<string> KKSIds, string cabinetName, string controllerIp)
{
    StringBuilder csvContent = new StringBuilder();
    // Статические заголовки
    csvContent.AppendLine("[HEADER],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
    csvContent.AppendLine("#ListSeparator,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,

,,,");
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

    // Статические заголовки DEVICE
    csvContent.AppendLine("[DEVICES],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
    csvContent.AppendLine("#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,,,,,,,,,,,,,,,,,,,,,,,,,,");
    csvContent.AppendLine("#[DeviceName],[DeviceDescription],[S7_Type],[IP_address],[AccessPointS7],[Projectname],[StationName],[Establish Connection],[Alias],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],,,,,,,,,,,,,,,,,,,,,,,,,,");
    csvContent.AppendLine($"{cabinetName},{cabinetName},S7-1500,{controllerIp},S7ONLINE,S7Plus$Online,Online,TRUE,PLC_{cabinetName},,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
    csvContent.AppendLine(@",,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");

    // Данные POINTS
    csvContent.AppendLine("[POINTS],,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,");
    csvContent.AppendLine("#MandatoryCommon,#MandatoryCommon,#MandatoryCommon,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#MandatorySubsys,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,#OptionalCommon,");
    csvContent.AppendLine("#[ParentDeviceName],[Name],[Description],[Address],[DataType],[Direction],[LowLevelComparison],[PollType],[PollGroup],[ObjectModel],[Property],[Alias],[Function],[Discipline],[Subdiscipline],[Type],[Subtype],[Min],[Max],[MinRaw],[MaxRaw],[MinEng],[MaxEng],[Resolution],[Unit],[UnitTextGroup],[StateText],[ActivityLog],[ValueLog],[Smoothing],[AlarmClass],[AlarmType],[AlarmValue],[EventText],[NormalText],[UpperHysteresis],[LowerHysteresis],[NoAlarmOn],[LogicalHierarchy],[UserHierarchy]");

    int device_number = 1;

    foreach (string name in KKSIds)
    {
        string model = GetObjectModel(name.Trim());
        if (model != "AKKUYU_NO_MODEL")
        {
            string[] parts = name.Split(' ');

            string kksname = parts[1];
            string kksnumber = parts[2];
            string buildingnumber = parts[3];

            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.KKS,String,IO,FALSE,COV,pollGr_3,{model},KKS,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{cabinetName}\\,");
            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].OPRT.OPRT_INDEX,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT_INDEX,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{

cabinetName}\\,");
            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.Mode,Uint,IO,FALSE,COV,pollGr_3,{model},Mode,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{cabinetName}\\,");
            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.State,Uint,IO,FALSE,COV,pollGr_3,{model},State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{cabinetName}\\,");
            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].ALM.Alarm_word_1,Uint,IO,FALSE,COV,pollGr_3,{model},Alarm_word_1,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,Alarm,EQ,1,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{cabinetName}\\,");
            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].OPRT.OPRT,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{cabinetName}\\,");
            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].ALM.Alarm.Alarm[0],bool,IO,FALSE,COV,pollGr_3,{model},Alarm_bit0,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{cabinetName}\\,");

            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.RASU_bits.RASU_bits[0],bool,IO,FALSE,COV,pollGr_3,{model},RASU_bits0,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{cabinetName}\\,");
            csvContent.AppendLine($"{cabinetName},{kksnumber},{kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.RASU_bits.RASU_bits[1],bool,IO,FALSE,COV,pollGr_3,{model},RASU_bits1,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingnumber}\\Mechanisms\\{cabinetName}\\,");
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
```

### Получение объектной модели по механизму

#### Метод: `GetObjectModel`

Этот метод возвращает модель объекта в зависимости от механизма.

```csharp
static string GetObjectModel(string mechanism)
{
    if (mechanism.Contains("AN"))
    {
        return "AKKUYU_MOTOR";
    }
    else if (mechanism.Contains("AA"))
    {
        return "AKKUYU_VALVE";
    }
    else
    {
        return "AKKUYU_NO_MODEL";
    }
}
```

### Чтение данных о шкафах и их IP

#### Метод: `ReadCabinetData`

Этот метод читает данные о шкафах и их IP из Excel файла.

```csharp
public static Dictionary<string, string> ReadCabinetData(string filePath, string sheetName)
{
    var data = new Dictionary<string, string>();
    try
    {
        using (var workbook = new XLWorkbook(filePath))
        {
            var worksheet = workbook.Worksheet(sheetName);
            var kksColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains("KKS"))?.Address.ColumnNumber;
            var ipColumn = worksheet.FirstRowUsed().CellsUsed().FirstOrDefault(c => c.Value.ToString().Contains(@"IP S7_A1"))?.Address.ColumnNumber;

            if (!kksColumn.HasValue || !ipColumn.HasValue)
                return data;

            foreach (var row in worksheet.RowsUsed().Skip(1))
            {
                var kksValue = row.Cell(kksColumn.Value).GetValue<string>().Trim();
                var ipValue = row.Cell(ipColumn.Value).GetValue<string>().Trim();

                if (!string.IsNullOrWhiteSpace(kksValue) && !data.ContainsKey(kksValue))
                {
                    if (string.IsNullOrWhiteSpace(ipValue))
                        ipValue = "NO IP";

                    data[kksValue] = ipValue;
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
```

### Обработчик кнопки для отображения информации о программе

#### Метод: `btnInfo_Click`

Этот метод показывает информацию о программе в виде сообщения.

```csharp
private void btnInfo_Click(object sender, EventArgs e)
{
    string infoMessage = "EXCEL_parser - программа для конвертации данных из Excel в CSV.\n\n" +
                          "Инструкция:\n" +
                          "1. Нажмите 'Выбрать файл Excel' и выберите файл Excel.\n" +
                          "2. Нажмите 'Выбрать папку' и укажите папку для сохранения CSV файлов.\n" +
                          "3. Нажмите 'Конвертировать', чтобы начать процесс конвертации.\n\n";

    MessageBox.Show(infoMessage, "Информация о программе", MessageBoxButtons.OK, MessageBoxIcon.Information);
}
```

### Создание и заполнение базы данных

#### Метод: `button5_Click`

Этот метод создает базу данных и заполняет её данными из Excel файла.

```csharp
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
```

### Метод для получения количества записей в таблице

#### Метод: `GetTableCountAsync`

Этот метод возвращает количество записей в таблице базы данных.

```csharp
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
```

### Метод для создания и заполнения базы данных

#### Метод: `CreateAndFillDatabaseAsync`

Этот метод создает и заполняет базу данных.

```csharp
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
                    string wcc_oa_tag = $"System00_2:ManagementView_FieldNetworks_S7Plus_{cabinetName}_{kksnumber}.{tag.Item1}";
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
```

### Метод для проверки существования idx в базе данных

#### Метод: `IdxExistsAsync`

Этот метод проверяет, существует ли указанный idx в базе данных.

```csharp
private async Task<bool> IdxExistsAsync(SQLiteConnection connection, string idx)
{
    using (var command = new SQLiteCommand("SELECT COUNT(*) FROM (SELECT idx FROM messages_analog WHERE idx = @idx UNION ALL SELECT idx FROM messages_binary WHERE idx = @idx UNION ALL SELECT idx FROM messages_integer WHERE idx = @idx)", connection))
    {
        command.Parameters.AddWithValue("@idx", idx);
        return Convert.ToInt32(await command.ExecuteScalarAsync()) > 0;
    }
}
```

### Создание второй базы данных

#### Метод: `btnCreateSecondDatabase_Click`

Этот метод создает вторую базу данных и заполняет её данными.

```csharp
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
```

### Метод для создания и заполнения второй базы данных

#### Метод: `CreateAndFillSecondDatabaseAsync`

Этот метод создает и заполняет вторую базу данных.

```csharp
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
            (1, 'DtsManager', 'ipc:///tmp/S_dts', 'ipc:///tmp/R_dts', 'Менеджер кан

ала DTS.'),
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
```

### Метод для получения всех idx из первой базы данных

#### Метод: `GetAllIdxFromFirstDatabaseAsync`

Этот метод возвращает все idx из первой базы данных.

```csharp
private async Task<List<Tuple<string, int, int>>> GetAllIdxFromFirstDatabaseAsync()
{
    var allIdx = new List<Tuple<string, int, int>>();

    using (var connection = new SQLiteConnection($"Data Source={Path.Combine(txtOutputDirectory.Text, "DuplicatedDatabase.db")}; Version=3;"))
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
```

### Создание базы данных с дублированием точек данных

#### Метод: `btnCreateDuplicatedDatabase_Click`

Этот метод создает базу данных с дублированием точек данных.

```csharp
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

        await CreateAndFillDuplicatedDatabaseAsync(duplicatedDatabasePath);

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
```

### Метод для создания и заполнения базы данных с дублированием точек данных

#### Метод: `CreateAndFillDuplicatedDatabaseAsync`

Этот метод создает и заполняет базу данных с дублированием точек данных.

```csharp
private async Task CreateAndFillDuplicatedDatabaseAsync(string dbPath)
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

        var cabinetsInfo = GetCabinetsInfo(txtExcelFilePath.Text);

        foreach (var cabinetInfo in cabinetsInfo)
        {
            var lines = cabinetInfo.Split(',');
            List<string> kkss = lines.Skip(2).ToList();
            string cabinetName = lines[0];
            string controllerIp = lines[1];

            foreach (var name in kkss)
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
                        string idx1 = baseIdx;
                        string idx2 = $"{baseIdx}_2";
                        string wcc_oa_tag1 = $"System00_1:ManagementView_FieldNetworks_S7Plus_{cabinetName}_{kksnumber}.{tag.Item1}";
                        string wcc_oa_tag2 = $"System00_2:ManagementView_FieldNetworks_S7Plus_{cabinetName}_{kksnumber}.{tag.Item1}";
                        string comment = $"{tag.Item1} {cabinetName}_{kksname}";

                        int counter = 1;
                        while (existingIdx.Contains

(idx1) || await IdxExistsAsync(connection, idx1))
                        {
                            idx1 = $"{baseIdx}_{counter}";
                            idx2 = $"{baseIdx}_{counter}_2";
                            counter++;
                        }

                        existingIdx.Add(idx1);
                        existingIdx.Add(idx2);

                        string entry1 = $"('{idx1}', '{wcc_oa_tag1}', '{comment}')";
                        string entry2 = $"('{idx2}', '{wcc_oa_tag2}', '{comment}')";

                        if (tag.Item2 == "bool")
                        {
                            binaryEntries.Add(entry1);
                            binaryEntries.Add(entry2);
                        }
                        else
                        {
                            integerEntries.Add(entry1);
                            integerEntries.Add(entry2);
                        }
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
            string integerSql = $"INSERT INTO messages_integer (idx, wcc_oa_tag, comment) VALUES {string.join(", ", integerEntries)};";
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
```

