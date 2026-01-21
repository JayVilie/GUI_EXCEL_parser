using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using GUI_EXCEL_parser.Properties;

namespace GUI_EXCEL_parser
{
    public partial class Converter : Form
    {
        private CancellationTokenSource _cts;
        private bool _isBusy;
        private LogVerbosity _logVerbosity = LogVerbosity.Short;
        private readonly ValidationSummary _lastValidationSummary = new ValidationSummary();
        private readonly Stopwatch _operationStopwatch = new Stopwatch();
        private readonly System.Windows.Forms.Timer _statusTimer = new System.Windows.Forms.Timer();
        private int _lastProgressValue;
        private DateTime _lastProgressUpdate;
        private string _currentStep = "Ожидание";
        private int _errorCount;
        private int _warningCount;
        private Color _normalTextBackColor;
        private readonly Dictionary<string, Profile> _profiles = new Dictionary<string, Profile>(StringComparer.OrdinalIgnoreCase);

        private const long MinFreeSpaceBytes = 50L * 1024 * 1024; // 50 МБ
        private const int AutoSaveEvery = 50;

        private enum LogVerbosity
        {
            Short,
            Detailed
        }

        public Converter()
        {
            InitializeComponent();
            InitializeCustomComponents();
            LoadSettings();
            InitializeStatusTimer();
            this.KeyPreview = true;
            this.KeyDown += Converter_KeyDown;
            UpdateIssuesFromSummary(new ValidationSummary());
            UpdateStatusBar();
            InitializeLogFilter();
            UpdateRunHistory();
            LogInfo("Приложение запущено и готово к работе.");
        }

        private void InitializeLogFilter()
        {
            if (clbLogFilter == null || clbLogFilter.Items.Count == 0)
                return;

            for (int i = 0; i < clbLogFilter.Items.Count; i++)
                clbLogFilter.SetItemChecked(i, true);
        }

        private void UpdateRunHistory()
        {
            try
            {
                string raw = Settings.Default.LastRunHistory ?? string.Empty;
                var list = raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries).ToList();
                string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                list.Insert(0, now);
                list = list.Distinct().Take(5).ToList();
                Settings.Default.LastRunHistory = string.Join(";", list);
                Settings.Default.Save();

                if (list.Count > 1)
                    LogInfo($"Предыдущие запуски: {string.Join(", ", list.Skip(1))}");
            }
            catch
            {
                // История запусков не должна ломать запуск приложения.
            }
        }

        // Инициализация визуальных параметров формы и кнопок
        private void InitializeCustomComponents()
        {
            this.btnParseDpl.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnConvert.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button3.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button4.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.button5.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCreateSecondDatabase.FlatStyle = System.Windows.Forms.FlatStyle.Flat;

            // Применяем темную тему для всех контролов формы
            ApplyDarkTheme();
            _normalTextBackColor = txtExcelFilePath.BackColor;
        }

        // Применяет темную тему ко всем элементам формы
        private void ApplyDarkTheme()
        {
            this.BackColor = Color.FromArgb(18, 18, 18);
            this.ForeColor = Color.Gainsboro;
            ApplyDarkThemeToControls(this);
        }

        // Рекурсивно настраивает цвета контролов
        private void ApplyDarkThemeToControls(Control parent)
        {
            foreach (Control control in parent.Controls)
            {
                if (control is Label label)
                {
                    label.ForeColor = Color.Gainsboro;
                    label.BackColor = Color.Transparent;
                }
                else if (control is TextBox textBox)
                {
                    textBox.BackColor = Color.FromArgb(30, 30, 30);
                    textBox.ForeColor = Color.Gainsboro;
                    textBox.BorderStyle = BorderStyle.FixedSingle;
                }
                else if (control is Button button)
                {
                    button.BackColor = Color.FromArgb(45, 45, 48);
                    button.ForeColor = Color.Gainsboro;
                    button.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
                }
                else if (control is ComboBox comboBox)
                {
                    comboBox.BackColor = Color.FromArgb(30, 30, 30);
                    comboBox.ForeColor = Color.Gainsboro;
                    comboBox.FlatStyle = FlatStyle.Flat;
                }
                else if (control is RichTextBox richTextBox)
                {
                    richTextBox.BackColor = Color.FromArgb(24, 24, 24);
                    richTextBox.ForeColor = Color.Gainsboro;
                    richTextBox.BorderStyle = BorderStyle.FixedSingle;
                }
                else if (control is ProgressBar progress)
                {
                    progress.BackColor = Color.FromArgb(30, 30, 30);
                    progress.ForeColor = Color.Gainsboro;
                }
                else if (control is GroupBox groupBox)
                {
                    groupBox.BackColor = Color.FromArgb(18, 18, 18);
                    groupBox.ForeColor = Color.Gainsboro;
                }
                else if (control is StatusStrip statusStrip)
                {
                    statusStrip.BackColor = Color.FromArgb(18, 18, 18);
                    statusStrip.ForeColor = Color.Gainsboro;
                    foreach (ToolStripItem item in statusStrip.Items)
                    {
                        item.ForeColor = Color.Gainsboro;
                    }
                }
                else if (control is Panel || control is FlowLayoutPanel)
                {
                    control.BackColor = Color.FromArgb(18, 18, 18);
                    control.ForeColor = Color.Gainsboro;
                }
                else
                {
                    control.BackColor = Color.FromArgb(18, 18, 18);
                    control.ForeColor = Color.Gainsboro;
                }

                if (control.HasChildren)
                {
                    ApplyDarkThemeToControls(control);
                }
            }
        }

        // Загрузка сохранённых настроек пользователя
        private void LoadSettings()
        {
            txtExcelFilePath.Text = Settings.Default.LastExcelPath ?? string.Empty;
            txtOutputDirectory.Text = Settings.Default.LastOutputDirectory ?? string.Empty;

            cmbLogLevel.SelectedIndexChanged -= cmbLogLevel_SelectedIndexChanged;
            cmbLogLevel.SelectedIndex = Settings.Default.LogVerbosity == "Подробно" ? 1 : 0;
            cmbLogLevel.SelectedIndexChanged += cmbLogLevel_SelectedIndexChanged;

            _logVerbosity = cmbLogLevel.SelectedIndex == 1 ? LogVerbosity.Detailed : LogVerbosity.Short;
            LoadProfiles();
        }

        // Сохранение текущих настроек пользователя
        private void SaveSettings()
        {
            Settings.Default.LastExcelPath = txtExcelFilePath.Text;
            Settings.Default.LastOutputDirectory = txtOutputDirectory.Text;
            Settings.Default.LogVerbosity = _logVerbosity == LogVerbosity.Detailed ? "Подробно" : "Кратко";
            Settings.Default.Save();
        }

        private void LoadProfiles()
        {
            _profiles.Clear();
            cmbProfile.Items.Clear();

            string raw = Settings.Default.ProfilesSerialized ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(raw))
            {
                foreach (var entry in raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    var parts = entry.Split('|');
                    if (parts.Length != 4) continue;
                    string name = Decode(parts[0]);
                    var profile = new Profile
                    {
                        ExcelPath = Decode(parts[1]),
                        OutputDirectory = Decode(parts[2]),
                        LogVerbosity = Decode(parts[3])
                    };
                    if (!string.IsNullOrWhiteSpace(name))
                        _profiles[name] = profile;
                }
            }

            foreach (var name in _profiles.Keys.OrderBy(n => n))
                cmbProfile.Items.Add(name);

            if (!string.IsNullOrWhiteSpace(Settings.Default.LastProfileName))
            {
                cmbProfile.SelectedItem = Settings.Default.LastProfileName;
            }
        }

        private void SaveProfiles()
        {
            var list = new List<string>();
            foreach (var kv in _profiles)
            {
                string entry = string.Join("|",
                    Encode(kv.Key),
                    Encode(kv.Value.ExcelPath ?? string.Empty),
                    Encode(kv.Value.OutputDirectory ?? string.Empty),
                    Encode(kv.Value.LogVerbosity ?? "Кратко"));
                list.Add(entry);
            }
            Settings.Default.ProfilesSerialized = string.Join(";", list);
            Settings.Default.Save();
        }

        private static string Encode(string value)
        {
            if (value == null) return string.Empty;
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(value));
        }

        private static string Decode(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            try
            {
                return Encoding.UTF8.GetString(Convert.FromBase64String(value));
            }
            catch
            {
                return string.Empty;
            }
        }

        private void cmbProfile_SelectedIndexChanged(object sender, EventArgs e)
        {
            string name = cmbProfile.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(name) || !_profiles.TryGetValue(name, out var profile))
                return;

            txtExcelFilePath.Text = profile.ExcelPath ?? string.Empty;
            txtOutputDirectory.Text = profile.OutputDirectory ?? string.Empty;

            cmbLogLevel.SelectedIndexChanged -= cmbLogLevel_SelectedIndexChanged;
            cmbLogLevel.SelectedIndex = profile.LogVerbosity == "Подробно" ? 1 : 0;
            cmbLogLevel.SelectedIndexChanged += cmbLogLevel_SelectedIndexChanged;

            _logVerbosity = cmbLogLevel.SelectedIndex == 1 ? LogVerbosity.Detailed : LogVerbosity.Short;
            Settings.Default.LastProfileName = name;
            Settings.Default.Save();

            LogInfo($"Применён профиль: {name}");
        }

        private void btnSaveProfile_Click(object sender, EventArgs e)
        {
            string name = PromptForProfileName();
            if (string.IsNullOrWhiteSpace(name))
                return;

            _profiles[name] = new Profile
            {
                ExcelPath = txtExcelFilePath.Text,
                OutputDirectory = txtOutputDirectory.Text,
                LogVerbosity = _logVerbosity == LogVerbosity.Detailed ? "Подробно" : "Кратко"
            };

            SaveProfiles();
            LoadProfiles();
            cmbProfile.SelectedItem = name;
            LogInfo($"Профиль сохранён: {name}");
        }

        private void btnDeleteProfile_Click(object sender, EventArgs e)
        {
            string name = cmbProfile.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(name))
                return;

            if (!_profiles.ContainsKey(name))
                return;

            var confirm = MessageBox.Show($"Удалить профиль \"{name}\"?", "Подтверждение", MessageBoxButtons.YesNo);
            if (confirm != DialogResult.Yes)
                return;

            _profiles.Remove(name);
            SaveProfiles();
            LoadProfiles();
            LogInfo($"Профиль удалён: {name}");
        }

        private string PromptForProfileName()
        {
            using (var form = new Form())
            {
                form.Text = "Название профиля";
                form.StartPosition = FormStartPosition.CenterParent;
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.Width = 360;
                form.Height = 140;
                form.MaximizeBox = false;
                form.MinimizeBox = false;

                var textBox = new TextBox { Left = 12, Top = 12, Width = 320 };
                var btnOk = new Button { Text = "OK", Left = 172, Width = 80, Top = 50, DialogResult = DialogResult.OK };
                var btnCancel = new Button { Text = "Отмена", Left = 252, Width = 80, Top = 50, DialogResult = DialogResult.Cancel };

                form.Controls.Add(textBox);
                form.Controls.Add(btnOk);
                form.Controls.Add(btnCancel);
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                return form.ShowDialog(this) == DialogResult.OK ? textBox.Text.Trim() : string.Empty;
            }
        }

        private void cmbLogLevel_SelectedIndexChanged(object sender, EventArgs e)
        {
            _logVerbosity = cmbLogLevel.SelectedIndex == 1 ? LogVerbosity.Detailed : LogVerbosity.Short;
            SaveSettings();
            LogInfo($"Уровень логирования изменён: {Settings.Default.LogVerbosity}");
        }

        private void InitializeStatusTimer()
        {
            _statusTimer.Interval = 1000;
            _statusTimer.Tick += (s, e) => UpdateStatusBar();
            _statusTimer.Start();
        }

        private void Converter_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.O)
            {
                btnSelectExcelFile_Click(sender, e);
                e.Handled = true;
            }
            else if (e.Control && e.Shift && e.KeyCode == Keys.O)
            {
                btnSelectOutputDirectory_Click(sender, e);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.Enter)
            {
                btnConvert.PerformClick();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.B)
            {
                button5.PerformClick();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.E)
            {
                btnExportByBuilding.PerformClick();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.P)
            {
                btnParseDpl.PerformClick();
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.L)
            {
                btnOpenOutput.PerformClick();
                e.Handled = true;
            }
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
                    LogInfo($"Выбран Excel-файл: {openFileDialog.FileName}");
                    SaveSettings();
                }
                else
                {
                    LogInfo("Выбор Excel-файла отменён пользователем.");
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
                    LogInfo($"Выбрана папка вывода: {folderDialog.SelectedPath}");
                    SaveSettings();
                }
                else
                {
                    LogInfo("Выбор папки вывода отменён пользователем.");
                }
            }
        }

        // Обработчик кнопки начала конвертации
        private async void btnConvert_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Вы уверены, что хотите создать файлы CSV?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                LogInfo("Конвертация CSV отменена пользователем.");
                return;
            }

            string excelFilePath = txtExcelFilePath.Text;
            string outputDirectory = txtOutputDirectory.Text;
            try
            {
                if (!ValidateExcelAndOutput(excelFilePath, outputDirectory, out var validationMessage))
                {
                    LogError("Проверка входных данных не пройдена. Конвертация отменена.");
                    MessageBox.Show(validationMessage, "Ошибка проверки данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                StartOperation("Чтение Excel", 0);
                SetOperationState(true);
                LogInfo($"Старт конвертации CSV. Файл: {excelFilePath}. Папка: {outputDirectory}");
                // Обновление текста, состояния кнопки и отображение PictureBox
                btnConvert.Text = "Конвертация...";
                btnConvert.Enabled = false;
                progressBar1.Value = 0;

                // Читаем все шкафы и механизмы отдельно, чтобы статистика была корректной.
                var cabinetData = await Task.Run(() => ExcelParser.ReadCabinetData(excelFilePath, "Шкафы"));
                var mechData = await Task.Run(() => ExcelParser.ReadMechData(excelFilePath, "Мех-мы"));
                LogInfo($"Загружены данные Excel: шкафов={cabinetData.Count}, наборов механизмов={mechData.Count}");
                int totalFiles = 0;
                int totalMechanisms = 0;
                int cabinetsWithoutIP = 0;
                int cabinetsWithoutMechanisms = 0;
                progressBar1.Maximum = Math.Max(1, cabinetData.Count);
                UpdateProgress(0, progressBar1.Maximum, "Генерация CSV");

                foreach (var cabinetEntry in cabinetData)
                {
                    _cts?.Token.ThrowIfCancellationRequested();
                    string cabinetKks = cabinetEntry.Key;
                    var cabinet = cabinetEntry.Value;
                    LogInfoDetailed($"Обработка шкафа: {cabinetKks}");

                    if (!IsValidKks(cabinetKks))
                    {
                        LogWarning($"Некорректный KKS шкафа: {cabinetKks}");
                        _warningCount++;
                    }

                    bool hasIp = !string.IsNullOrWhiteSpace(cabinet.IP) &&
                                 !string.Equals(cabinet.IP.Trim(), "NO IP", StringComparison.OrdinalIgnoreCase);
                    if (!hasIp)
                    {
                        cabinetsWithoutIP++;
                        LogWarning($"Шкаф без IP: {cabinetKks}");
                    }
                    else if (!IsValidIp(cabinet.IP))
                    {
                        LogWarning($"Некорректный IP у шкафа {cabinetKks}: {cabinet.IP}");
                        _warningCount++;
                    }

                    bool hasMechanisms = mechData.TryGetValue(cabinetKks, out var mechanisms) &&
                                         mechanisms != null &&
                                         mechanisms.Count > 0;
                    if (!hasMechanisms)
                    {
                        cabinetsWithoutMechanisms++;
                        LogWarning($"Шкаф без механизмов: {cabinetKks}");
                    }

                    if (hasMechanisms)
                    {
                        // Формируем список строк в ожидаемом формате для CsvExporter.
                        var lines = new List<string>
                        {
                            cabinetKks,
                            cabinet.IP,
                            cabinet.Building,
                            cabinet.Type
                        };
                        lines.AddRange(mechanisms);

                        // Генерация CSV вынесена в CsvExporter.
                        await Task.Run(() => CsvExporter.GenerateCsv(lines, outputDirectory, LogInfoDetailed, LogWarning));
                        totalFiles++;
                        totalMechanisms += mechanisms.Count;
                        LogInfoDetailed($"CSV создан: {cabinetKks}, механизмов={mechanisms.Count}");
                    }

                    progressBar1.Value++;
                    UpdateProgress(progressBar1.Value, progressBar1.Maximum, "Генерация CSV");
                    labelProgressDetail.Text = $"Шкафы: {progressBar1.Value} / {progressBar1.Maximum}";

                    if (progressBar1.Value % AutoSaveEvery == 0)
                    {
                        AutoSaveProgress(outputDirectory, totalFiles, totalMechanisms, cabinetsWithoutIP, cabinetsWithoutMechanisms);
                    }
                }

                // Восстановление текста, состояния кнопки и скрытие PictureBox
                btnConvert.Text = "Конвертировать";
                btnConvert.Enabled = true;

                // Сообщение об успешном завершении
                LogInfo($"Конвертация завершена. Файлов={totalFiles}, механизмов={totalMechanisms}, шкафов без IP={cabinetsWithoutIP}, без механизмов={cabinetsWithoutMechanisms}");
                ShowSummaryReport("Итоговый отчёт: Конвертация CSV", totalFiles, totalMechanisms, cabinetsWithoutIP, cabinetsWithoutMechanisms);
                MessageBox.Show($"Преобразование завершено.\n" +
                                $"Всего создано файлов: {totalFiles}\n" +
                                $"Всего механизмов: {totalMechanisms}\n" +
                                $"Шкафов без IP: {cabinetsWithoutIP}\n" +
                                $"Шкафов без механизмов: {cabinetsWithoutMechanisms}",
                                "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                LogWarning("Конвертация CSV отменена пользователем.");
                MessageBox.Show("Операция конвертации была отменена.", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnConvert.Text = "Конвертировать";
                btnConvert.Enabled = true;
                LogError($"Ошибка конвертации CSV: {ex.Message}");
                MessageBox.Show($"Произошла ошибка: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnConvert.Text = "Конвертировать";
                btnConvert.Enabled = true;
                EndOperation("Ожидание");
                SetOperationState(false);
            }
        }

        // Метод для получения информации о шкафах и механизмах
        static List<string> GetCabinetsInfo(string filePath)
        {
            // Делегируем чтение Excel отдельному слою.
            return ExcelParser.GetCabinetsInfo(filePath);
        }

        // Функция для очистки строки от нежелательных символов и пробелов
        private static string CleanString(string input)
        {
            // Используем общую реализацию из CsvExporter.
            return CsvExporter.CleanString(input);
        }

        // Метод для создания CSV файла
        static void GenerateCsv(List<string> lines, string outputDirectory)
        {
            // Делегируем генерацию CSV отдельному слою.
            CsvExporter.GenerateCsv(lines, outputDirectory);
        }

        // Метод для получения данных по ККС шкафа из листа "Алгоритмы"
        public static List<Dictionary<string, string>> GetAlgorithmDataByKKS(string filePath, string kksCabinet)
        {
            // Делегируем чтение Excel в ExcelParser.
            return ExcelParser.GetAlgorithmDataByKKS(filePath, kksCabinet);
        }

        // Метод для чтения данных о механизмах
        static Dictionary<string, List<string>> ReadMechData(string filePath, string sheetName)
        {
            // Делегируем чтение Excel в ExcelParser.
            return ExcelParser.ReadMechData(filePath, sheetName);
        }

        // Генерация содержимого CSV файла
        static StringBuilder GenerateCsvContent(List<string> KKSIds, string cabinetName, string controllerIp, string buildingName, string type)
        {
            // Делегируем генерацию CSV отдельному слою.
            return CsvExporter.GenerateCsvContent(KKSIds, cabinetName, controllerIp, buildingName, type);

#if false
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

                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},{GetBuildingDifference(buildingnumber, buildingName) + kksname},DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.KKS,String,IO,FALSE,COV,pollGr_3,{model},KKS,,,,,,,,,,,,,,,,,FALSE,FALSE,SmoothingConfig_01,,,,,,,,DriverFail,\\{buildingName}\\Mechanisms\\{buildingName + "_" + cabinetName}\\,");

                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].OPRT.OPRT_INDEX,Uint,IO,FALSE,COV,pollGr_3,{model},OPRT_INDEX,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.Mode,Uint,IO,FALSE,COV,pollGr_3,{model},Mode,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
                    csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.State,Uint,IO,FALSE,COV,pollGr_3,{model},State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");

                    // Проверяем наличие "AKKUYU_VALVE" в типе и добавляем строку при соблюдении условия
                    if (model.Contains("AKKUYU_VALVE"))
                    {
                        csvContent.AppendLine($"{buildingName + "_" + cabinetName},{kksnumber},,DB_HMI_IO_Mechanisms.Devices.Devices[{kksnumber}].PRM.Extinguishing_State,Uint,IO,FALSE,COV,pollGr_3,{model},Extinguishing_State,,,,,,,,,,,,,,,,,TRUE,TRUE,SmoothingConfig_01,,,,,,,,DriverFail,,");
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
#endif
        }

        // Получение объектной модели по ККС механизма
        static string GetObjectModel(string mechanism)
        {
            // Делегируем определение модели в общий резолвер.
            return MechanismModelResolver.GetObjectModel(mechanism);
        }

        // Формирует префикс здания, если оно отличается от шкафа
        public static string GetBuildingDifference(string buildingNumber, string buildingName)
        {
            // Делегируем в CsvExporter.
            return CsvExporter.GetBuildingDifference(buildingNumber, buildingName);
        }

        // Возвращает список "чужих" зданий для описания устройства
        public static string GetDistinctBuildingsWithExclusion(List<string> KKSIds, string buildingName)
        {
            // Делегируем в CsvExporter.
            return CsvExporter.GetDistinctBuildingsWithExclusion(KKSIds, buildingName);
        }

        // Получение данных по шкафам с указанного листа Excel
        public static Dictionary<string, (string IP, string Building, string Type, string VEKSH)> ReadCabinetData(string filePath, string sheetName)
        {
            // Делегируем чтение Excel в ExcelParser.
            return ExcelParser.ReadCabinetData(filePath, sheetName);
        }

        // Нормализует тип шкафа с учетом ВЕКШ
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
            string infoMessage =
                "Наименование ПО: SKUPZ_Generation_P\n" +
                "Версия: см. свойства сборки\n" +
                "Назначение: подготовка CSV и баз данных для интеграции Desigo CC / WinCC OA\n" +
                "Область применения: конфигурация СКУПЗ\n\n" +

                "Требования:\n" +
                "  • ОС Windows, .NET Framework 4.7.2\n" +
                "  • Права на запись в папку вывода\n\n" +

                "Входные данные:\n" +
                "  • Excel-файл со следующими листами:\n" +
                "    - «Шкафы»: KKS шкафа, IP, здание, тип, ВЕКШ\n" +
                "    - «Мех-мы»: KKS шкафа, KKS механизма, Desigo number, здание\n" +
                "    - «Алгоритмы»: опционально (выгрузка зон)\n\n" +

                "Выходные данные:\n" +
                "  • CSV-файлы формата DEVICES/POINTS (S7Plus)\n" +
                "  • SQLite БД: MyDatabase.db, SecondDatabase.db, SystemXX_Y.db\n\n" +

                "Основные функции:\n" +
                "  • «Конвертировать» — CSV по шкафам\n" +
                "  • «CSV по блокам» — CSV по блокам 00/10/20/30/40 (тот же формат)\n" +
                "  • «Создать базу данных» — MyDatabase.db\n" +
                "  • «Создать вторую базу данных» — SecondDatabase.db (routing)\n" +
                "  • «Создать базу данных с дублированием» — System00_1/00_2/10_1/10_2\n" +
                "  • «Парсить DPL» — добавление точек из DPL\n\n" +

                "Порядок работы:\n" +
                "  1) Выберите Excel-файл\n" +
                "  2) Выберите папку вывода\n" +
                "  3) Запустите нужную операцию\n\n" +

                "Ограничения и ошибки:\n" +
                "  • Некорректные заголовки листов/колонок дают пустой результат\n" +
                "  • При наличии БД/CSV предлагается перезапись\n" +
                "  • При ошибках выводится сообщение с описанием";

            MessageBox.Show(infoMessage, "О программе", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Обработчик кнопки для создания базы данных
        private async void button5_Click(object sender, EventArgs e)
        {
            string databasePath = Path.Combine(txtOutputDirectory.Text, "MyDatabase.db");

            var result = MessageBox.Show("Вы уверены, что хотите создать новую базу данных?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                LogInfo("Создание БД отменено пользователем.");
                return;
            }

            if (!ValidateExcelAndOutput(txtExcelFilePath.Text, txtOutputDirectory.Text, out var validationMessage))
            {
                LogError("Проверка входных данных не пройдена. Создание БД отменено.");
                MessageBox.Show(validationMessage, "Ошибка проверки данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (File.Exists(databasePath))
            {
                var overwriteResult = MessageBox.Show("База данных уже существует. Перезаписать?", "Подтверждение", MessageBoxButtons.YesNo);
                if (overwriteResult == DialogResult.No)
                {
                    LogInfo("Перезапись БД отменена пользователем.");
                    return;
                }

                var backupPath = BackupFileIfExists(databasePath);
                LogWarning($"Существующая БД сохранена в резерв: {backupPath}");
            }

            try
            {
                StartOperation("Создание базы данных", 0);
                SetOperationState(true);
                LogInfo($"Старт создания базы данных: {databasePath}");
                button5.Text = "Создание базы данных...";
                button5.Enabled = false;

                // Читаем Excel через ExcelParser.
                var cabinetsInfo = ExcelParser.GetCabinetsInfo(txtExcelFilePath.Text);
                int totalPoints = 0;

                foreach (var cabinetInfo in cabinetsInfo)
                {
                    var lines = cabinetInfo.Split(',');
                    List<string> kkss = lines.Skip(2).ToList();
                    totalPoints += kkss.Count;
                }

                _cts?.Token.ThrowIfCancellationRequested();

                // Схема БД создаётся один раз на файл, дальше вставка точек.
                await DatabaseWriter.CreateAndFillDatabaseAsync(databasePath, cabinetsInfo);

                var fileInfo = new FileInfo(databasePath);
                long fileSize = fileInfo.Length;

                int binaryCount = await GetTableCountAsync(databasePath, "messages_binary");
                int integerCount = await GetTableCountAsync(databasePath, "messages_integer");
                int analogCount = await GetTableCountAsync(databasePath, "messages_analog");

                button5.Text = "Создать базу данных";
                button5.Enabled = true;

                LogInfo($"БД создана. Файл={databasePath}, размер={fileSize}, binary={binaryCount}, integer={integerCount}, analog={analogCount}");
                MessageBox.Show($"База данных успешно создана!\nФайл: {databasePath}\nРазмер: {fileSize} байт\nКоличество точек данных:\n - messages_binary: {binaryCount}\n - messages_integer: {integerCount}\n - messages_analog: {analogCount}\nОбщее количество точек данных: {totalPoints}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                LogWarning("Создание БД отменено пользователем.");
                MessageBox.Show("Операция создания БД была отменена.", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                button5.Text = "Создать базу данных";
                button5.Enabled = true;
                LogError($"Ошибка создания БД: {ex.Message}");
                MessageBox.Show($"Ошибка при создании базы данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                button5.Text = "Создать базу данных";
                button5.Enabled = true;
                EndOperation("Ожидание");
                SetOperationState(false);
            }
        }

        // Метод для получения количества записей в таблице
        private async Task<int> GetTableCountAsync(string dbPath, string tableName)
        {
            // Делегируем чтение количества строк в DatabaseWriter.
            return await DatabaseWriter.GetTableCountAsync(dbPath, tableName);
        }

        // Метод для создания и заполнения базы данных
        private async Task CreateAndFillDatabaseAsync(string dbPath, List<string> KKSIds, string cabinetName, string controllerIp)
        {
            // Метод оставлен для совместимости, но логика перенесена в DatabaseWriter.
            // Формируем строку в формате, совместимом с GetCabinetsInfo.
            var cabinetsInfo = new List<string>
            {
                $"{cabinetName},{controllerIp}" + (KKSIds.Count > 0 ? ", " + string.Join(", ", KKSIds) : string.Empty)
            };
            await DatabaseWriter.CreateAndFillDatabaseAsync(dbPath, cabinetsInfo);
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
            SaveSettings();
        }

        private void txtExcelFilePath_TextChanged(object sender, EventArgs e)
        {
            SaveSettings();
            AutoValidateExcel();
        }

        // Обработчик кнопки для создания второй базы данных
        private async void btnCreateSecondDatabase_Click(object sender, EventArgs e)
        {
            string secondDatabasePath = Path.Combine(txtOutputDirectory.Text, "SecondDatabase.db");

            var result = MessageBox.Show("Вы уверены, что хотите создать новую базу данных?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                LogInfo("Создание второй БД отменено пользователем.");
                return;
            }

            if (!ValidateOutputDirectory(txtOutputDirectory.Text, out var outputMessage))
            {
                LogError("Проверка папки вывода не пройдена. Создание второй БД отменено.");
                MessageBox.Show(outputMessage, "Ошибка проверки папки", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (File.Exists(secondDatabasePath))
            {
                var overwriteResult = MessageBox.Show("База данных уже существует. Перезаписать?", "Подтверждение", MessageBoxButtons.YesNo);
                if (overwriteResult == DialogResult.No)
                {
                    LogInfo("Перезапись второй БД отменена пользователем.");
                    return;
                }

                var backupPath = BackupFileIfExists(secondDatabasePath);
                LogWarning($"Существующая вторая БД сохранена в резерв: {backupPath}");
            }

            try
            {
                StartOperation("Создание второй базы данных", 0);
                SetOperationState(true);
                LogInfo($"Старт создания второй БД: {secondDatabasePath}");
                btnCreateSecondDatabase.Text = "Создание базы данных...";
                btnCreateSecondDatabase.Enabled = false;

                _cts?.Token.ThrowIfCancellationRequested();
                await CreateAndFillSecondDatabaseAsync(secondDatabasePath);

                btnCreateSecondDatabase.Text = "Создать вторую базу данных";
                btnCreateSecondDatabase.Enabled = true;

                LogInfo($"Вторая БД создана: {secondDatabasePath}");
                MessageBox.Show($"Вторая база данных успешно создана!\nФайл: {secondDatabasePath}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                LogWarning("Создание второй БД отменено пользователем.");
                MessageBox.Show("Операция создания второй БД была отменена.", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnCreateSecondDatabase.Text = "Создать вторую базу данных";
                btnCreateSecondDatabase.Enabled = true;
                LogError($"Ошибка создания второй БД: {ex.Message}");
                MessageBox.Show($"Ошибка при создании второй базы данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnCreateSecondDatabase.Text = "Создать вторую базу данных";
                btnCreateSecondDatabase.Enabled = true;
                EndOperation("Ожидание");
                SetOperationState(false);
            }
        }

        // Метод для создания и заполнения второй базы данных
        private async Task CreateAndFillSecondDatabaseAsync(string dbPath)
        {
            // Делегируем создание второй БД в DatabaseWriter.
            string sourceDbPath = Path.Combine(txtOutputDirectory.Text, "MyDatabase.db");
            await DatabaseWriter.CreateAndFillSecondDatabaseAsync(dbPath, sourceDbPath);
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
            // Папка для системных БД (System00_1..System10_2)
            string outputDirectory = txtOutputDirectory.Text;

            var result = MessageBox.Show("Вы уверены, что хотите создать новую базу данных с дублированными точками данных?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No)
            {
                LogInfo("Создание дублированных БД отменено пользователем.");
                return;
            }

            if (!ValidateExcelAndOutput(txtExcelFilePath.Text, txtOutputDirectory.Text, out var validationMessage))
            {
                LogError("Проверка входных данных не пройдена. Создание системных БД отменено.");
                MessageBox.Show(validationMessage, "Ошибка проверки данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Проверяем, есть ли уже системные БД.
            string[] systemFiles = { "System00_1.db", "System00_2.db", "System10_1.db", "System10_2.db" };
            bool anyExists = systemFiles.Any(f => File.Exists(Path.Combine(outputDirectory, f)));
            if (anyExists)
            {
                var overwriteResult = MessageBox.Show("Базы систем уже существуют. Перезаписать?", "Подтверждение", MessageBoxButtons.YesNo);
                if (overwriteResult == DialogResult.No)
                    return;

                foreach (var file in systemFiles)
                {
                    var fullPath = Path.Combine(outputDirectory, file);
                    if (File.Exists(fullPath))
                        BackupFileIfExists(fullPath);
                }
                LogWarning("Существующие системные БД сохранены в резерв перед пересозданием.");
            }

            try
            {
                StartOperation("Создание системных баз данных", 0);
                SetOperationState(true);
                LogInfo($"Старт создания системных БД в папке: {outputDirectory}");
                btnCreateDuplicatedDatabase.Text = "Создание базы данных...";
                btnCreateDuplicatedDatabase.Enabled = false;

                // ВАЖНО: метод ждёт папку, а не файл.
                _cts?.Token.ThrowIfCancellationRequested();
                await CreateAndFillSystemDatabasesAsync(outputDirectory);

                btnCreateDuplicatedDatabase.Text = "Создать базу данных";
                btnCreateDuplicatedDatabase.Enabled = true;

                LogInfo($"Системные БД созданы. Папка: {outputDirectory}");
                MessageBox.Show($"Базы данных успешно созданы!\nПапка: {outputDirectory}", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                LogWarning("Создание системных БД отменено пользователем.");
                MessageBox.Show("Операция создания системных БД была отменена.", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnCreateDuplicatedDatabase.Text = "Создать базу данных";
                btnCreateDuplicatedDatabase.Enabled = true;
                LogError($"Ошибка создания системных БД: {ex.Message}");
                MessageBox.Show($"Ошибка при создании базы данных: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnCreateDuplicatedDatabase.Text = "Создать базу данных";
                btnCreateDuplicatedDatabase.Enabled = true;
                EndOperation("Ожидание");
                SetOperationState(false);
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
            // Делегируем в DatabaseWriter: он работает с папкой, а не с файлом.
            await DatabaseWriter.CreateAndFillSystemDatabasesAsync(dirPath, txtExcelFilePath.Text);
            return;

#if false
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
#endif
        }

        // Обработчик кнопки для парсинга DPL и добавления данных в базу
        private void btnParseDpl_Click(object sender, EventArgs e)
        {
            string dplFilePath = GetDplFilePath();
            if (string.IsNullOrEmpty(dplFilePath) || !File.Exists(dplFilePath))
            {
                LogError($"Файл DPL не найден: {dplFilePath}");
                MessageBox.Show($"Ошибка: Входной файл '{dplFilePath}' не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string databasePath = GetDatabaseFilePath();
            if (string.IsNullOrEmpty(databasePath) || !File.Exists(databasePath))
            {
                LogError($"Файл базы данных не найден: {databasePath}");
                MessageBox.Show($"Ошибка: Файл базы данных '{databasePath}' не найден.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                LogInfo($"Старт парсинга DPL. Файл: {dplFilePath}. База: {databasePath}");
                List<string> matchingLines = GetMatchingLinesFromDplFile(dplFilePath);
                if (matchingLines.Count == 0)
                {
                    LogWarning("Совпадающие строки в DPL не найдены.");
                    MessageBox.Show("Не найдено строк, соответствующих заданным паттернам.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                List<(string dataPoint, string comment, string cabinetCode)> matchingData = matchingLines.Select(ExtractData).ToList();
                AddDataPointsToDatabase(databasePath, matchingData);
                LogInfo($"Данные из DPL добавлены в БД. Строк={matchingData.Count}");
                MessageBox.Show("Данные успешно добавлены в базу данных.", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                LogError($"Ошибка при добавлении данных из DPL: {ex.Message}");
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

        // Метод для получения строк из файла DPL, содержащих новые паттерны в LANG:10027
        // ВАЖНО: cabinetCode в комментарии может отсутствовать (как в примерах с "ППКиУП"),
        // поэтому мы НЕ фильтруем строку по IsCabinetCode.
        private List<string> GetMatchingLinesFromDplFile(string filePath)
        {
            // Делегируем разбор DPL в DplParser.
            return DplParser.GetMatchingLinesFromDplFile(filePath);
        }

        // Разбиваем строку в “нормализованные” токены: пробелы и '_' — разделители
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

        // Проверка “commentTokens начинается с patternTokens” (без чувствительности к регистру)
        private static bool StartsWithTokens(string[] commentTokens, string[] patternTokens)
        {
            if (commentTokens == null || patternTokens == null) return false;
            if (commentTokens.Length < patternTokens.Length) return false;

            for (int i = 0; i < patternTokens.Length; i++)
            {
                if (!string.Equals(commentTokens[i], patternTokens[i], StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            return true;
        }

        private void AddDataPointsToDatabase(string databasePath, List<(string dataPoint, string comment, string cabinetCode)> matchingData)
        {
            // Делегируем вставку в DatabaseWriter (с параметризацией).
            DatabaseWriter.AddDataPointsToDatabase(databasePath, matchingData);
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
            // Делегируем в DplParser.
            return DplParser.SplitLine(line);
        }

        // Метод для проверки, является ли строка кодом ККС
        static bool IsCabinetCode(string code)
        {
            // Делегируем в DplParser.
            return DplParser.IsCabinetCode(code);
        }

        // Метод для извлечения точки данных, комментария и кода ККС из строки
        static (string dataPoint, string comment, string cabinetCode) ExtractData(string line)
        {
            // Делегируем извлечение данных в DplParser.
            return DplParser.ExtractData(line);
        }

        // Метод для вставки точки данных в базу данных
        static void InsertDataPoint(SQLiteConnection connection, string sql)
        {
            using (var command = new SQLiteCommand(sql, connection))
            {
                command.ExecuteNonQuery();
            }
        }

        // Экспорт: 1 файл = 1 блок (00/10/20/30/40); внутри — несколько DEVICES (по шкафам) + POINTS (по их механизмам)
        private async void btnExportByBuilding_Click(object sender, EventArgs e)
        {
            string excelFilePath = txtExcelFilePath.Text;
            string outputDirectory = txtOutputDirectory.Text;

            if (string.IsNullOrWhiteSpace(excelFilePath) || !File.Exists(excelFilePath))
            {
                LogError("Экспорт по блокам: некорректный Excel-файл.");
                MessageBox.Show("Выбери корректный Excel-файл.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrWhiteSpace(outputDirectory) || !Directory.Exists(outputDirectory))
            {
                LogError("Экспорт по блокам: некорректная папка вывода.");
                MessageBox.Show("Выбери существующую папку для вывода.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var result = MessageBox.Show("Создать CSV-файлы по блокам (00/10/20/30/40)?", "Подтверждение", MessageBoxButtons.YesNo);
            if (result == DialogResult.No) return;

            try
            {
                if (!ValidateExcelAndOutput(excelFilePath, outputDirectory, out var validationMessage))
                {
                    LogError("Проверка входных данных не пройдена. Экспорт по блокам отменен.");
                    MessageBox.Show(validationMessage, "Ошибка проверки данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                StartOperation("Чтение Excel", 0);
                SetOperationState(true);
                LogInfo($"Старт экспорта по блокам. Файл: {excelFilePath}. Папка: {outputDirectory}");
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
                LogInfo($"Найдено блоков для экспорта: {blocksTotal}");
                UpdateProgress(0, progressBar1.Maximum, "Формирование CSV по блокам");

                foreach (var block in blocksOrdered)
                {
                    _cts?.Token.ThrowIfCancellationRequested();
                    List<string> cabinetKks = byBlock[block];
                    if (cabinetKks == null || cabinetKks.Count == 0)
                    {
                        progressBar1.Value++;
                        UpdateProgress(progressBar1.Value, progressBar1.Maximum, "Формирование CSV по блокам");
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

                        // Формат строки DEVICES совпадает с обычной конвертацией.
                        CsvExporter.AppendCabinetDeviceLine(sb, cabinetNm, ip, building, type, kkss);
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

                        // Формат POINTS совпадает с обычной конвертацией.
                        CsvExporter.AppendCabinetPoints(sb, cabinetNm, building, type, kkss);
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

                    if (File.Exists(path))
                    {
                        var backupPath = BackupFileIfExists(path);
                        LogWarning($"Файл блока сохранён в резерв: {backupPath}");
                    }

                    using (var writer = new StreamWriter(path, false, new UTF8Encoding(true)))
                        await writer.WriteAsync(sb.ToString());

                    filesCreated++;
                    progressBar1.Value++;
                    LogInfo($"Файл блока создан: {path}");
                    UpdateProgress(progressBar1.Value, progressBar1.Maximum, "Формирование CSV по блокам");
                    labelProgressDetail.Text = $"Блоки: {progressBar1.Value} / {progressBar1.Maximum}";
                }

                btnExportByBuilding.Enabled = true;
                btnExportByBuilding.Text = "CSV по блокам";

                LogInfo($"Экспорт по блокам завершен. Файлов создано: {filesCreated}");
                ShowSummaryReport("Итоговый отчёт: Экспорт по блокам", filesCreated, 0, 0, 0);
                MessageBox.Show($"Готово. Создано файлов: {filesCreated} (по {blocksTotal} блокам).",
                    "Экспорт по блокам", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (OperationCanceledException)
            {
                LogWarning("Экспорт по блокам отменён пользователем.");
                MessageBox.Show("Операция экспорта по блокам была отменена.", "Отмена", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                btnExportByBuilding.Enabled = true;
                btnExportByBuilding.Text = "CSV по блокам";
                LogError($"Ошибка экспорта по блокам: {ex.Message}");
                MessageBox.Show($"Ошибка экспорта: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnExportByBuilding.Enabled = true;
                btnExportByBuilding.Text = "CSV по блокам";
                EndOperation("Ожидание");
                SetOperationState(false);
            }
        }

        // Склеивает части с одинарным пробелом и без двойных пробелов
        private static string BuildDeviceDescription(string distinct, string cabinet, string type)
        {
            // Делегируем в CsvExporter.
            return CsvExporter.BuildDeviceDescription(distinct, cabinet, type);
        }

        private static (string mechKks, string mechNum, string mechBld) ParseMechTriplet(string raw, string fallbackBuilding = "")
        {
            // Делегируем в CsvExporter.
            return CsvExporter.ParseMechTriplet(raw, fallbackBuilding);
        }

        private static bool IsUint(string s) => CsvExporter.IsUint(s);

        // Возвращает "00","10","20","30","40" либо null, если не распознано
        // Возвращает "00","10","20","30","40" с нормализацией префиксов (0x→00, 1x→10, 2x→20, 3x→30, 4x→40)
        private static string GetBlockFromBuildingOrKks(string building, string cabinetKks)
        {
            // Делегируем в CsvExporter.
            return CsvExporter.GetBlockFromBuildingOrKks(building, cabinetKks);
        }

        private static string GetBlockLabel(string block)
        {
            // Делегируем в CsvExporter.
            return CsvExporter.GetBlockLabel(block);
        }

        /// <param name="csvContent"></param>
        // Общая «шапка» CSV (как в GenerateCsvContent, но один раз на файл)
        private static void AppendStandardHeader(StringBuilder csvContent)
        {
            // Делегируем в CsvExporter.
            CsvExporter.AppendStandardHeader(csvContent);
            return;

#if false
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
#endif
        }

        private static void AppendDevicesHeader(StringBuilder csvContent)
        {
            // Делегируем в CsvExporter.
            CsvExporter.AppendDevicesHeader(csvContent);
        }

        private static void AppendPointsHeader(StringBuilder csvContent)
        {
            // Делегируем в CsvExporter.
            CsvExporter.AppendPointsHeader(csvContent);
        }

        // Диагностический «блок шкафа» — как в твоём GenerateCsvContent
        private static void AppendCabinetDiagnostics(StringBuilder sb, string buildingName, string cabinetName, string type)
        {
            // Делегируем в CsvExporter.
            CsvExporter.AppendCabinetDiagnostics(sb, buildingName, cabinetName, type);
            return;

#if false
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
#endif
        }

        // безопасно достаёт список чужих зданий (кроме текущего шкафа)
        private static string GetDistinctBuildingsForBlockExport(List<string> kkss, string deviceBuilding)
        {
            // Делегируем в CsvExporter.
            return CsvExporter.GetDistinctBuildingsForBlockExport(kkss, deviceBuilding);
        }

        private void Converter_Load(object sender, EventArgs e)
        {

        }

        // Очистка журнала в интерфейсе
        private void btnClearLog_Click(object sender, EventArgs e)
        {
            if (rtbLog == null) return;

            rtbLog.Clear();
            LogInfo("Журнал очищен пользователем.");
        }

        private void btnShowLogFile_Click(object sender, EventArgs e)
        {
            try
            {
                string logPath = GetCurrentLogFilePath();
                if (!File.Exists(logPath))
                {
                    MessageBox.Show("Файл лога ещё не создан.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                Process.Start(logPath);
                LogInfo($"Открыт файл лога: {logPath}");
            }
            catch (Exception ex)
            {
                LogError($"Не удалось открыть лог: {ex.Message}");
                MessageBox.Show($"Не удалось открыть лог: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnExportLog_Click(object sender, EventArgs e)
        {
            try
            {
                using (var dialog = new SaveFileDialog())
                {
                    dialog.Filter = "Текстовый файл (*.txt)|*.txt|Все файлы (*.*)|*.*";
                    dialog.FileName = $"log_export_{DateTime.Now:yyyyMMdd_HHmmss}.txt";
                    dialog.OverwritePrompt = true;
                    if (dialog.ShowDialog() != DialogResult.OK)
                        return;

                    File.WriteAllText(dialog.FileName, rtbLog?.Text ?? string.Empty, Encoding.UTF8);
                    Settings.Default.LastLogDirectory = Path.GetDirectoryName(dialog.FileName);
                    Settings.Default.Save();
                    LogInfo($"Лог экспортирован: {dialog.FileName}");
                }
            }
            catch (Exception ex)
            {
                LogError($"Ошибка экспорта лога: {ex.Message}");
                MessageBox.Show($"Ошибка экспорта лога: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void clbLogFilter_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Пересоздаём вывод не будем, фильтр влияет только на новые сообщения.
            LogInfoDetailed("Фильтр лога изменён пользователем.");
        }

        private void btnOpenLastFolder_Click(object sender, EventArgs e)
        {
            string folder = Settings.Default.LastReportDirectory;
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                folder = Settings.Default.LastLogDirectory;

            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                MessageBox.Show("Последняя папка отчётов или логов не найдена.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                Process.Start("explorer.exe", folder);
                LogInfo($"Открыта папка: {folder}");
            }
            catch (Exception ex)
            {
                LogError($"Не удалось открыть папку: {ex.Message}");
                MessageBox.Show($"Не удалось открыть папку: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Открыть папку вывода в Проводнике
        private void btnOpenOutput_Click(object sender, EventArgs e)
        {
            string outputDir = txtOutputDirectory.Text;
            if (string.IsNullOrWhiteSpace(outputDir) || !Directory.Exists(outputDir))
            {
                LogError("Папка вывода не задана или не существует.");
                MessageBox.Show("Папка вывода не задана или не существует.", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                Process.Start("explorer.exe", outputDir);
                LogInfo($"Открыта папка вывода: {outputDir}");
            }
            catch (Exception ex)
            {
                LogError($"Не удалось открыть папку вывода: {ex.Message}");
                MessageBox.Show($"Не удалось открыть папку: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Отмена текущей операции
        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (_cts == null || _cts.IsCancellationRequested)
                return;

            _cts.Cancel();
            LogWarning("Запрошена отмена операции. Операция завершится при ближайшей проверке.");
            UpdateProgress(progressBar1.Value, progressBar1.Maximum, "Отмена операции...");
        }

        // Запускает операцию с поддержкой отмены
        private void StartOperation(string step, int total)
        {
            _cts?.Dispose();
            _cts = new CancellationTokenSource();
            _operationStopwatch.Restart();
            _currentStep = step;
            _lastProgressValue = 0;
            _lastProgressUpdate = DateTime.Now;
            UpdateProgress(0, total <= 0 ? 1 : total, step);
        }

        // Завершает операцию и сбрасывает состояние
        private void EndOperation(string step)
        {
            _cts?.Dispose();
            _cts = null;
            _operationStopwatch.Reset();
            _currentStep = step;
            UpdateProgress(progressBar1.Value, Math.Max(progressBar1.Maximum, 1), step);
        }

        // Управление доступностью кнопок при выполнении операции
        private void SetOperationState(bool isRunning)
        {
            if (_isBusy == isRunning)
                return;

            _isBusy = isRunning;

            foreach (Control control in panelButtons.Controls)
            {
                if (control == btnCancel) continue;
                control.Enabled = !isRunning;
            }

            btnCancel.Enabled = isRunning;
        }

        // Обновляет индикатор прогресса, проценты и шаг
        private void UpdateProgress(int current, int total, string step)
        {
            if (progressBar1.InvokeRequired || labelProgress.InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateProgress(current, total, step)));
                return;
            }

            int safeTotal = Math.Max(1, total);
            int safeCurrent = Math.Max(0, Math.Min(current, safeTotal));
            progressBar1.Maximum = safeTotal;
            progressBar1.Value = safeCurrent;

            int percent = (int)Math.Round((safeCurrent / (double)safeTotal) * 100);
            labelProgress.Text = $"Прогресс: {percent}% — {step}";

            _currentStep = step;
            _lastProgressValue = safeCurrent;
            _lastProgressUpdate = DateTime.Now;
        }

        private void UpdateStatusBar()
        {
            if (statusStrip1.InvokeRequired)
            {
                BeginInvoke(new Action(UpdateStatusBar));
                return;
            }

            statusLabelAction.Text = _currentStep;

            var elapsed = _operationStopwatch.Elapsed;
            statusLabelElapsed.Text = $"Время: {elapsed:mm\\:ss}";

            int total = Math.Max(1, progressBar1.Maximum);
            int current = Math.Max(0, progressBar1.Value);
            double seconds = Math.Max(1, elapsed.TotalSeconds);
            double speed = current / seconds;
            statusLabelSpeed.Text = $"Скорость: {speed:0.0}/с";

            if (speed > 0.1 && current <= total)
            {
                double remaining = (total - current) / speed;
                var eta = TimeSpan.FromSeconds(Math.Max(0, remaining));
                statusLabelEta.Text = $"ETA: {eta:mm\\:ss}";
            }
            else
            {
                statusLabelEta.Text = "ETA: --:--";
            }
        }

        private void AutoSaveProgress(string outputDir, int totalFiles, int totalMechanisms, int cabinetsWithoutIp, int cabinetsWithoutMechanisms)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(outputDir) || !Directory.Exists(outputDir))
                    return;

                string path = Path.Combine(outputDir, "autosave_progress.txt");
                var sb = new StringBuilder();
                sb.AppendLine($"Время: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"Шкафы обработано: {progressBar1.Value} / {progressBar1.Maximum}");
                sb.AppendLine($"Файлов создано: {totalFiles}");
                sb.AppendLine($"Механизмов: {totalMechanisms}");
                sb.AppendLine($"Шкафов без IP: {cabinetsWithoutIp}");
                sb.AppendLine($"Шкафов без механизмов: {cabinetsWithoutMechanisms}");
                File.WriteAllText(path, sb.ToString(), Encoding.UTF8);
                LogInfoDetailed($"Автосохранение прогресса: {path}");
            }
            catch
            {
                // Автосохранение не должно ломать основную работу.
            }
        }

        // Проверка Excel и папки вывода перед запуском операции
        private bool ValidateExcelAndOutput(string excelPath, string outputDir, out string message)
        {
            message = string.Empty;
            _lastValidationSummary.Clear();

            if (!ValidateExcelStructure(excelPath, _lastValidationSummary))
            {
                foreach (var err in _lastValidationSummary.Errors)
                    LogError(err);
                UpdateIssuesFromSummary(_lastValidationSummary);
                message = BuildValidationMessage(_lastValidationSummary, "Проверка Excel-файла не пройдена.");
                return false;
            }

            if (!ValidateOutputDirectory(outputDir, out var outputMessage))
            {
                message = outputMessage;
                return false;
            }

            if (_lastValidationSummary.HasWarnings)
            {
                LogWarning("Обнаружены предупреждения в структуре данных Excel.");
                foreach (var warn in _lastValidationSummary.Warnings)
                    LogWarning(warn);

                var warnMessage = BuildValidationMessage(_lastValidationSummary, "Данные Excel содержат предупреждения.");
                MessageBox.Show(warnMessage, "Предупреждения проверки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            UpdateIssuesFromSummary(_lastValidationSummary);

            message = BuildValidationMessage(_lastValidationSummary, "Проверка данных пройдена с предупреждениями.");
            return true;
        }

        private void AutoValidateExcel()
        {
            if (string.IsNullOrWhiteSpace(txtExcelFilePath.Text))
            {
                txtExcelFilePath.BackColor = _normalTextBackColor;
                toolTip1.SetToolTip(txtExcelFilePath, string.Empty);
                return;
            }

            var summary = new ValidationSummary();
            bool ok = ValidateExcelStructure(txtExcelFilePath.Text, summary);
            UpdateIssuesFromSummary(summary);

            if (!ok)
            {
                txtExcelFilePath.BackColor = Color.FromArgb(60, 20, 20);
                toolTip1.SetToolTip(txtExcelFilePath, BuildValidationMessage(summary, "Проверка Excel не пройдена."));
            }
            else
            {
                txtExcelFilePath.BackColor = _normalTextBackColor;
                toolTip1.SetToolTip(txtExcelFilePath, summary.HasWarnings
                    ? "Есть предупреждения по данным. Подробности в журнале."
                    : "Excel-файл корректен.");
            }
        }

        private void UpdateIssuesFromSummary(ValidationSummary summary)
        {
            _errorCount = summary.Errors.Count;
            _warningCount = summary.Warnings.Count +
                            summary.DuplicateCabinetKks.Count +
                            summary.DuplicateMechKks.Count +
                            summary.InvalidCabinetRows.Count +
                            summary.InvalidMechRows.Count;

            if (labelIssues.InvokeRequired)
            {
                BeginInvoke(new Action(() => UpdateIssuesFromSummary(summary)));
                return;
            }

            labelIssues.Text = $"Ошибки: {_errorCount}, Предупр.: {_warningCount}";
        }

        // Валидация папки вывода: доступ на запись + свободное место
        private bool ValidateOutputDirectory(string outputDir, out string message)
        {
            message = string.Empty;
            if (string.IsNullOrWhiteSpace(outputDir) || !Directory.Exists(outputDir))
            {
                message = "Папка вывода не задана или не существует.";
                return false;
            }

            try
            {
                string testFile = Path.Combine(outputDir, $"_write_test_{Guid.NewGuid():N}.tmp");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
            }
            catch (Exception ex)
            {
                message = $"Нет доступа на запись в папку вывода: {ex.Message}";
                return false;
            }

            try
            {
                var root = Path.GetPathRoot(outputDir);
                var drive = new DriveInfo(root);
                if (drive.AvailableFreeSpace < MinFreeSpaceBytes)
                {
                    message = $"Недостаточно свободного места на диске. Требуется минимум {MinFreeSpaceBytes / (1024 * 1024)} МБ.";
                    return false;
                }
            }
            catch (Exception ex)
            {
                message = $"Не удалось определить свободное место на диске: {ex.Message}";
                return false;
            }

            return true;
        }

        // Проверяет наличие обязательных листов и колонок, дублей и некорректных строк
        private bool ValidateExcelStructure(string excelPath, ValidationSummary summary)
        {
            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                summary.Errors.Add("Excel-файл не найден.");
                return false;
            }

            try
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    if (!workbook.Worksheets.Contains("Шкафы"))
                        summary.Errors.Add("Отсутствует лист «Шкафы».");
                    if (!workbook.Worksheets.Contains("Мех-мы"))
                        summary.Errors.Add("Отсутствует лист «Мех-мы».");

                    if (summary.Errors.Count > 0)
                        return false;

                    var cabinetsSheet = workbook.Worksheet("Шкафы");
                    var mechSheet = workbook.Worksheet("Мех-мы");

                    // Колонки листа "Шкафы"
                    var cabKks = FindColumn(cabinetsSheet, "KKS");
                    var cabIp = FindColumn(cabinetsSheet, "IP S7_A1");
                    var cabBuilding = FindColumn(cabinetsSheet, "Здание");
                    var cabType = FindColumn(cabinetsSheet, "Тип");
                    var cabVeksh = FindColumn(cabinetsSheet, "ВЕКШ");

                    if (!cabKks.HasValue || !cabIp.HasValue || !cabBuilding.HasValue || !cabType.HasValue || !cabVeksh.HasValue)
                    {
                        summary.Errors.Add("На листе «Шкафы» отсутствуют обязательные колонки: KKS, IP S7_A1, Здание, Тип, ВЕКШ.");
                    }

                    // Колонки листа "Мех-мы"
                    var mechCabKks = FindColumn(mechSheet, "KKS шкафа");
                    var mechKks = FindColumn(mechSheet, "KKS механизма");
                    var mechNum = FindColumn(mechSheet, "Desigo number");

                    if (!mechCabKks.HasValue || !mechKks.HasValue || !mechNum.HasValue)
                    {
                        summary.Errors.Add("На листе «Мех-мы» отсутствуют обязательные колонки: KKS шкафа, KKS механизма, Desigo number.");
                    }

                    if (summary.Errors.Count > 0)
                        return false;

                    // Поиск дублей и некорректных строк на листе "Шкафы"
                    var cabinetKksMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    foreach (var row in cabinetsSheet.RowsUsed().Skip(1))
                    {
                        string kks = SafeCellString(row, cabKks.Value);
                        if (string.IsNullOrWhiteSpace(kks))
                        {
                            summary.InvalidCabinetRows.Add($"Строка {row.RowNumber()}: пустой KKS.");
                            continue;
                        }

                        if (!IsValidKks(kks))
                        {
                            summary.InvalidCabinetRows.Add($"Строка {row.RowNumber()}: некорректный формат KKS.");
                        }

                        if (cabinetKksMap.ContainsKey(kks))
                        {
                            summary.DuplicateCabinetKks.Add(kks);
                        }
                        else
                        {
                            cabinetKksMap[kks] = row.RowNumber();
                        }
                    }

                    // Поиск некорректных строк на листе "Мех-мы"
                    var mechKksMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    foreach (var row in mechSheet.RowsUsed().Skip(1))
                    {
                        string kksCab = SafeCellString(row, mechCabKks.Value);
                        string kksMech = SafeCellString(row, mechKks.Value);
                        string dNum = SafeCellString(row, mechNum.Value);

                        if (string.IsNullOrWhiteSpace(kksCab) || string.IsNullOrWhiteSpace(kksMech) || string.IsNullOrWhiteSpace(dNum))
                        {
                            summary.InvalidMechRows.Add($"Строка {row.RowNumber()}: нет KKS шкафа/механизма или Desigo number.");
                            continue;
                        }

                        if (!int.TryParse(dNum, out _))
                        {
                            summary.InvalidMechRows.Add($"Строка {row.RowNumber()}: Desigo number не является числом.");
                        }

                        if (!IsValidKks(kksCab) || !IsValidKks(kksMech))
                        {
                            summary.InvalidMechRows.Add($"Строка {row.RowNumber()}: некорректный формат KKS.");
                        }

                        if (mechKksMap.ContainsKey(kksMech))
                        {
                            summary.DuplicateMechKks.Add(kksMech);
                        }
                        else
                        {
                            mechKksMap[kksMech] = row.RowNumber();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                summary.Errors.Add($"Ошибка чтения Excel: {ex.Message}");
            }

            if (summary.DuplicateCabinetKks.Count > 0)
                summary.Warnings.Add($"Дубли KKS на листе «Шкафы»: {summary.DuplicateCabinetKks.Count}.");
            if (summary.DuplicateMechKks.Count > 0)
                summary.Warnings.Add($"Дубли KKS механизма на листе «Мех-мы»: {summary.DuplicateMechKks.Count}.");
            if (summary.InvalidCabinetRows.Count > 0)
                summary.Warnings.Add($"Некорректные строки на «Шкафы»: {summary.InvalidCabinetRows.Count}.");
            if (summary.InvalidMechRows.Count > 0)
                summary.Warnings.Add($"Некорректные строки на «Мех-мы»: {summary.InvalidMechRows.Count}.");

            // Проверка IP-адресов по листу "Шкафы"
            try
            {
                using (var workbook = new XLWorkbook(excelPath))
                {
                    var worksheet = workbook.Worksheet("Шкафы");
                    var ipColumn = FindColumn(worksheet, "IP S7_A1");
                    if (ipColumn.HasValue)
                    {
                        int invalidIpCount = 0;
                        foreach (var row in worksheet.RowsUsed().Skip(1))
                        {
                            string ip = SafeCellString(row, ipColumn.Value);
                            if (string.IsNullOrWhiteSpace(ip) || string.Equals(ip.Trim(), "NO IP", StringComparison.OrdinalIgnoreCase))
                                continue;
                            if (!IsValidIp(ip))
                                invalidIpCount++;
                        }
                        if (invalidIpCount > 0)
                            summary.Warnings.Add($"Некорректные IP: {invalidIpCount}.");
                    }
                }
            }
            catch
            {
                // Не критично: проверка IP не должна ломать валидатор.
            }

            return summary.Errors.Count == 0;
        }

        private static int? FindColumn(IXLWorksheet worksheet, string header)
        {
            return worksheet.FirstRowUsed()
                ?.CellsUsed()
                .FirstOrDefault(c => c.Value.ToString().Contains(header))
                ?.Address.ColumnNumber;
        }

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

        private bool IsValidIp(string ip)
        {
            if (string.IsNullOrWhiteSpace(ip))
                return false;
            if (string.Equals(ip.Trim(), "NO IP", StringComparison.OrdinalIgnoreCase))
                return false;
            return System.Net.IPAddress.TryParse(ip.Trim(), out _);
        }

        private bool IsValidKks(string kks)
        {
            if (string.IsNullOrWhiteSpace(kks))
                return false;
            // Базовая проверка: 6-20 символов, латиница/цифры/подчёркивание/дефис
            return Regex.IsMatch(kks.Trim(), @"^[A-Za-z0-9_-]{6,20}$");
        }

        // Создаёт резервную копию файла при перезаписи
        private string BackupFileIfExists(string filePath)
        {
            if (!File.Exists(filePath))
                return null;

            string dir = Path.GetDirectoryName(filePath) ?? string.Empty;
            string name = Path.GetFileNameWithoutExtension(filePath);
            string ext = Path.GetExtension(filePath);
            string backupName = $"{name}_backup_{DateTime.Now:yyyyMMdd_HHmmss}{ext}";
            string backupPath = Path.Combine(dir, backupName);
            File.Move(filePath, backupPath);
            return backupPath;
        }

        // Формирует окно итогового отчёта
        private void ShowSummaryReport(string title, int totalFiles, int totalMechanisms, int cabinetsWithoutIp, int cabinetsWithoutMechanisms)
        {
            var report = new StringBuilder();
            report.AppendLine(title);
            report.AppendLine(new string('-', 60));
            report.AppendLine($"Всего создано файлов: {totalFiles}");
            report.AppendLine($"Всего механизмов: {totalMechanisms}");
            report.AppendLine($"Шкафов без IP: {cabinetsWithoutIp}");
            report.AppendLine($"Шкафов без механизмов: {cabinetsWithoutMechanisms}");

            if (_lastValidationSummary.HasWarnings)
            {
                report.AppendLine();
                report.AppendLine("Предупреждения по данным:");
                foreach (var warn in _lastValidationSummary.Warnings)
                    report.AppendLine($"- {warn}");
            }

            if (_lastValidationSummary.DuplicateCabinetKks.Count > 0)
            {
                report.AppendLine();
                report.AppendLine("Дубли KKS (Шкафы):");
                foreach (var kks in _lastValidationSummary.DuplicateCabinetKks.Distinct().Take(20))
                    report.AppendLine($"- {kks}");
            }

            if (_lastValidationSummary.DuplicateMechKks.Count > 0)
            {
                report.AppendLine();
                report.AppendLine("Дубли KKS (Мех-мы):");
                foreach (var kks in _lastValidationSummary.DuplicateMechKks.Distinct().Take(20))
                    report.AppendLine($"- {kks}");
            }

            if (_lastValidationSummary.DuplicateMechKks.Count > 0)
            {
                report.AppendLine();
                report.AppendLine("Дубли KKS (Мех-мы):");
                foreach (var kks in _lastValidationSummary.DuplicateMechKks.Distinct().Take(20))
                    report.AppendLine($"- {kks}");
            }

            if (_lastValidationSummary.InvalidCabinetRows.Count > 0 || _lastValidationSummary.InvalidMechRows.Count > 0)
            {
                report.AppendLine();
                report.AppendLine("Некорректные строки:");
                foreach (var row in _lastValidationSummary.InvalidCabinetRows.Take(20))
                    report.AppendLine($"- Шкафы: {row}");
                foreach (var row in _lastValidationSummary.InvalidMechRows.Take(20))
                    report.AppendLine($"- Мех-мы: {row}");
            }

            var reportForm = new ReportForm(title, report.ToString(), this.BackColor, this.ForeColor);
            reportForm.Show(this);
        }

        // Формирует текст ошибок/предупреждений
        private string BuildValidationMessage(ValidationSummary summary, string header)
        {
            var sb = new StringBuilder();
            sb.AppendLine(header);
            sb.AppendLine();

            if (summary.Errors.Count > 0)
            {
                sb.AppendLine("Ошибки:");
                foreach (var err in summary.Errors)
                    sb.AppendLine($"- {err}");
            }

            if (summary.Warnings.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Предупреждения:");
                foreach (var warn in summary.Warnings)
                    sb.AppendLine($"- {warn}");
            }

            if (summary.DuplicateCabinetKks.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Примеры дублей KKS (Шкафы):");
                foreach (var kks in summary.DuplicateCabinetKks.Distinct().Take(10))
                    sb.AppendLine($"- {kks}");
            }

            if (summary.DuplicateMechKks.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Примеры дублей KKS (Мех-мы):");
                foreach (var kks in summary.DuplicateMechKks.Distinct().Take(10))
                    sb.AppendLine($"- {kks}");
            }

            if (summary.InvalidCabinetRows.Count > 0 || summary.InvalidMechRows.Count > 0)
            {
                sb.AppendLine();
                sb.AppendLine("Примеры некорректных строк:");
                foreach (var row in summary.InvalidCabinetRows.Take(5))
                    sb.AppendLine($"- Шкафы: {row}");
                foreach (var row in summary.InvalidMechRows.Take(5))
                    sb.AppendLine($"- Мех-мы: {row}");
            }

            return sb.ToString();
        }

        private sealed class ValidationSummary
        {
            public List<string> Errors { get; } = new List<string>();
            public List<string> Warnings { get; } = new List<string>();
            public List<string> DuplicateCabinetKks { get; } = new List<string>();
            public List<string> DuplicateMechKks { get; } = new List<string>();
            public List<string> InvalidCabinetRows { get; } = new List<string>();
            public List<string> InvalidMechRows { get; } = new List<string>();

            public bool HasWarnings => Warnings.Count > 0 ||
                                       DuplicateCabinetKks.Count > 0 ||
                                       DuplicateMechKks.Count > 0 ||
                                       InvalidCabinetRows.Count > 0 ||
                                       InvalidMechRows.Count > 0;

            public void Clear()
            {
                Errors.Clear();
                Warnings.Clear();
                DuplicateCabinetKks.Clear();
                DuplicateMechKks.Clear();
                InvalidCabinetRows.Clear();
                InvalidMechRows.Clear();
            }
        }

        private sealed class Profile
        {
            public string ExcelPath { get; set; }
            public string OutputDirectory { get; set; }
            public string LogVerbosity { get; set; }
        }

        // Добавляет запись в лог с форматированием по уровню
        private void AppendLog(string level, string message, bool isDetailed)
        {
            if (rtbLog == null) return;

            string timeStamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string line = $"[{timeStamp}] [{level}] {message}";

            if (level == "INFO" && isDetailed && _logVerbosity == LogVerbosity.Short)
                return;
            if (!IsLogLevelEnabled(level))
                return;

            if (rtbLog.InvokeRequired)
            {
                rtbLog.BeginInvoke(new Action(() => AppendLog(level, message, isDetailed)));
                return;
            }

            rtbLog.SelectionStart = rtbLog.TextLength;
            rtbLog.SelectionColor = GetLogColor(level, isDetailed);
            rtbLog.AppendText(line + Environment.NewLine);
            rtbLog.SelectionColor = rtbLog.ForeColor;
            rtbLog.SelectionStart = rtbLog.TextLength;
            rtbLog.ScrollToCaret();

            WriteLogToFile(line);
        }

        private Color GetLogColor(string level, bool isDetailed)
        {
            switch (level)
            {
                case "ERROR":
                    return Color.LightCoral;
                case "WARN":
                    return Color.Khaki;
                case "INFO":
                default:
                    return isDetailed ? Color.Gray : Color.Gainsboro;
            }
        }

        private bool IsLogLevelEnabled(string level)
        {
            if (clbLogFilter == null || clbLogFilter.Items.Count == 0)
                return true;

            for (int i = 0; i < clbLogFilter.Items.Count; i++)
            {
                if (string.Equals(clbLogFilter.Items[i]?.ToString(), level, StringComparison.OrdinalIgnoreCase))
                    return clbLogFilter.GetItemChecked(i);
            }

            return true;
        }

        // Информационное сообщение в лог
        private void LogInfo(string message)
        {
            AppendLog("INFO", message, false);
        }

        // Подробное информационное сообщение
        private void LogInfoDetailed(string message)
        {
            AppendLog("INFO", message, true);
        }

        // Предупреждение в лог
        private void LogWarning(string message)
        {
            AppendLog("WARN", message, false);
        }

        // Ошибка в лог
        private void LogError(string message)
        {
            AppendLog("ERROR", message, false);
        }

        // Сохранение лога в файл
        private void WriteLogToFile(string line)
        {
            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string logsDir = Path.Combine(baseDir, "logs");
                if (!Directory.Exists(logsDir))
                    Directory.CreateDirectory(logsDir);

                string filePath = Path.Combine(logsDir, $"{DateTime.Now:yyyy-MM-dd}.log");
                File.AppendAllText(filePath, line + Environment.NewLine, Encoding.UTF8);
                if (!string.Equals(Settings.Default.LastLogDirectory, logsDir, StringComparison.OrdinalIgnoreCase))
                {
                    Settings.Default.LastLogDirectory = logsDir;
                    Settings.Default.Save();
                }
            }
            catch
            {
                // Лог в файл не должен ломать работу приложения.
            }
        }

        private string GetCurrentLogFilePath()
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string logsDir = Path.Combine(baseDir, "logs");
            return Path.Combine(logsDir, $"{DateTime.Now:yyyy-MM-dd}.log");
        }
    }
}
