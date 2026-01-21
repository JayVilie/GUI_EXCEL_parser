// Converter.Designer.cs
using System;
using System.Drawing;
using System.Windows.Forms;

namespace GUI_EXCEL_parser
{
    partial class Converter
    {
        /// <summary>Обязательная переменная конструктора.</summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>Освободить все используемые ресурсы.</summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        private Label label1;
        private Label label2;
        private TextBox txtExcelFilePath;
        private TextBox txtOutputDirectory;
        private Button button1;                    // Выбрать файл Excel
        private Button btnConvert;                 // Конвертировать
        private Button button3;                    // Выбрать папку вывода
        private Button button4;                    // О программе
        private Button button5;                    // Создать базу данных
        private Button btnCreateSecondDatabase;    // Создать вторую БД
        private ProgressBar progressBar1;
        private Button btnCreateDuplicatedDatabase;// Создать БД с дублированием
        private Button btnParseDpl;                // Парсить DPL
        private Button btnExportByBuilding;        // CSV по блокам (00/10/20/30/40)
        private FlowLayoutPanel panelButtons;      // Панель для ровных кнопок
        private Label labelLog;
        private RichTextBox rtbLog;
        private Label labelProgress;
        private Button btnCancel;
        private Button btnOpenOutput;
        private Label labelLogLevel;
        private ComboBox cmbLogLevel;
        private Button btnClearLog;
        private GroupBox groupInput;
        private GroupBox groupActions;
        private GroupBox groupLog;
        private Button btnShowLogFile;
        private Button btnOpenLastFolder;
        private Label labelIssues;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel statusLabelAction;
        private ToolStripStatusLabel statusLabelElapsed;
        private ToolStripStatusLabel statusLabelSpeed;
        private ToolStripStatusLabel statusLabelEta;
        private Label labelProfile;
        private ComboBox cmbProfile;
        private Button btnSaveProfile;
        private Button btnDeleteProfile;
        private ToolTip toolTip1;
        private Button btnExportLog;
        private Label labelLogFilter;
        private CheckedListBox clbLogFilter;
        private Label labelProgressDetail;

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();

            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtExcelFilePath = new System.Windows.Forms.TextBox();
            this.txtOutputDirectory = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btnConvert = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.btnCreateSecondDatabase = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnCreateDuplicatedDatabase = new System.Windows.Forms.Button();
            this.btnParseDpl = new System.Windows.Forms.Button();
            this.btnExportByBuilding = new System.Windows.Forms.Button();
            this.panelButtons = new System.Windows.Forms.FlowLayoutPanel();
            this.labelLog = new System.Windows.Forms.Label();
            this.rtbLog = new System.Windows.Forms.RichTextBox();
            this.labelProgress = new System.Windows.Forms.Label();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnOpenOutput = new System.Windows.Forms.Button();
            this.labelLogLevel = new System.Windows.Forms.Label();
            this.cmbLogLevel = new System.Windows.Forms.ComboBox();
            this.btnClearLog = new System.Windows.Forms.Button();
            this.groupInput = new System.Windows.Forms.GroupBox();
            this.groupActions = new System.Windows.Forms.GroupBox();
            this.groupLog = new System.Windows.Forms.GroupBox();
            this.btnShowLogFile = new System.Windows.Forms.Button();
            this.btnOpenLastFolder = new System.Windows.Forms.Button();
            this.labelIssues = new System.Windows.Forms.Label();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.statusLabelAction = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusLabelElapsed = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusLabelSpeed = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusLabelEta = new System.Windows.Forms.ToolStripStatusLabel();
            this.labelProfile = new System.Windows.Forms.Label();
            this.cmbProfile = new System.Windows.Forms.ComboBox();
            this.btnSaveProfile = new System.Windows.Forms.Button();
            this.btnDeleteProfile = new System.Windows.Forms.Button();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.btnExportLog = new System.Windows.Forms.Button();
            this.labelLogFilter = new System.Windows.Forms.Label();
            this.clbLogFilter = new System.Windows.Forms.CheckedListBox();
            this.labelProgressDetail = new System.Windows.Forms.Label();

            this.SuspendLayout();

            // ---------- Форма ----------
            // Базовые параметры масштабирования — возвращаем к 96 dpi,
            // чтобы не "раздувало" контролы при повторном открытии в дизайнере.
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.ClientSize = new System.Drawing.Size(920, 620);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Converter";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "EXCEL → CSV / DB — Converter";

            // ---------- Группа: Входные данные ----------
            this.groupInput.Location = new System.Drawing.Point(12, 12);
            this.groupInput.Name = "groupInput";
            this.groupInput.Size = new System.Drawing.Size(896, 150);
            this.groupInput.TabIndex = 0;
            this.groupInput.TabStop = false;
            this.groupInput.Text = "Входные данные";
            this.groupInput.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // ---------- Метки ----------
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 15);
            this.label1.TabIndex = 100;
            this.label1.Text = "Excel-файл:";

            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 15);
            this.label2.TabIndex = 101;
            this.label2.Text = "Папка вывода:";

            // ---------- Поля ввода ----------
            this.txtExcelFilePath.Location = new System.Drawing.Point(116, 20);
            this.txtExcelFilePath.Name = "txtExcelFilePath";
            this.txtExcelFilePath.Size = new System.Drawing.Size(640, 23);
            this.txtExcelFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtExcelFilePath.TabIndex = 0;
            this.txtExcelFilePath.TextChanged += new System.EventHandler(this.txtExcelFilePath_TextChanged);

            this.txtOutputDirectory.Location = new System.Drawing.Point(116, 52);
            this.txtOutputDirectory.Name = "txtOutputDirectory";
            this.txtOutputDirectory.Size = new System.Drawing.Size(640, 23);
            this.txtOutputDirectory.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtOutputDirectory.TabIndex = 2;
            this.txtOutputDirectory.TextChanged += new System.EventHandler(this.txtOutputDirectory_TextChanged);

            // ---------- Кнопки рядом с полями ----------
            this.button1.Location = new System.Drawing.Point(764, 18);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(118, 27);
            this.button1.TabIndex = 1;
            this.button1.Text = "Выбрать…";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.button1.Click += new System.EventHandler(this.btnSelectExcelFile_Click);

            this.button3.Location = new System.Drawing.Point(764, 50);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(118, 27);
            this.button3.TabIndex = 3;
            this.button3.Text = "Папка…";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.button3.Click += new System.EventHandler(this.btnSelectOutputDirectory_Click);

            // ---------- Прогрессбар ----------
            this.progressBar1.Location = new System.Drawing.Point(12, 84);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(870, 18);
            this.progressBar1.TabIndex = 4;
            this.progressBar1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // ---------- Подпись прогресса ----------
            this.labelProgress.AutoSize = true;
            this.labelProgress.Location = new System.Drawing.Point(12, 104);
            this.labelProgress.Name = "labelProgress";
            this.labelProgress.Size = new System.Drawing.Size(170, 15);
            this.labelProgress.TabIndex = 150;
            this.labelProgress.Text = "Прогресс: 0% — Ожидание";

            // ---------- Счетчик ошибок/предупреждений ----------
            this.labelIssues.AutoSize = true;
            this.labelIssues.Location = new System.Drawing.Point(560, 104);
            this.labelIssues.Name = "labelIssues";
            this.labelIssues.Size = new System.Drawing.Size(130, 15);
            this.labelIssues.TabIndex = 151;
            this.labelIssues.Text = "Ошибки: 0, Предупр.: 0";

            // ---------- Подробный прогресс ----------
            this.labelProgressDetail.AutoSize = true;
            this.labelProgressDetail.Location = new System.Drawing.Point(12, 124);
            this.labelProgressDetail.Name = "labelProgressDetail";
            this.labelProgressDetail.Size = new System.Drawing.Size(170, 15);
            this.labelProgressDetail.TabIndex = 152;
            this.labelProgressDetail.Text = "Шкафы: 0 / 0";

            // ---------- Группа: Действия ----------
            this.groupActions.Location = new System.Drawing.Point(12, 170);
            this.groupActions.Name = "groupActions";
            this.groupActions.Size = new System.Drawing.Size(896, 180);
            this.groupActions.TabIndex = 1;
            this.groupActions.TabStop = false;
            this.groupActions.Text = "Операции";
            this.groupActions.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // ---------- Панель кнопок (ровная сетка) ----------
            this.panelButtons.Location = new System.Drawing.Point(12, 24);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Size = new System.Drawing.Size(872, 144);
            this.panelButtons.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.panelButtons.AutoSize = false;
            this.panelButtons.WrapContents = true;
            this.panelButtons.FlowDirection = FlowDirection.LeftToRight;
            this.panelButtons.Padding = new Padding(0);
            this.panelButtons.Margin = new Padding(0);
            this.panelButtons.TabIndex = 5;

            // ---------- Группа: Журнал ----------
            this.groupLog.Location = new System.Drawing.Point(12, 360);
            this.groupLog.Name = "groupLog";
            this.groupLog.Size = new System.Drawing.Size(896, 236);
            this.groupLog.TabIndex = 2;
            this.groupLog.TabStop = false;
            this.groupLog.Text = "Журнал";
            this.groupLog.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

            // ---------- Журнал действий ----------
            this.labelLog.AutoSize = true;
            this.labelLog.Location = new System.Drawing.Point(12, 24);
            this.labelLog.Name = "labelLog";
            this.labelLog.Size = new System.Drawing.Size(115, 15);
            this.labelLog.TabIndex = 200;
            this.labelLog.Text = "Журнал операций:";

            this.rtbLog.Location = new System.Drawing.Point(12, 112);
            this.rtbLog.Name = "rtbLog";
            this.rtbLog.Size = new System.Drawing.Size(872, 102);
            this.rtbLog.ReadOnly = true;
            this.rtbLog.HideSelection = false;
            this.rtbLog.WordWrap = false;
            this.rtbLog.ScrollBars = RichTextBoxScrollBars.Vertical;
            this.rtbLog.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            this.rtbLog.TabIndex = 6;

            // ---------- Уровень логирования ----------
            this.labelLogLevel.AutoSize = true;
            this.labelLogLevel.Location = new System.Drawing.Point(330, 24);
            this.labelLogLevel.Name = "labelLogLevel";
            this.labelLogLevel.Size = new System.Drawing.Size(134, 15);
            this.labelLogLevel.TabIndex = 201;
            this.labelLogLevel.Text = "Уровень логирования:";

            this.cmbLogLevel.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbLogLevel.FormattingEnabled = true;
            this.cmbLogLevel.Items.AddRange(new object[] { "Кратко", "Подробно" });
            this.cmbLogLevel.Location = new System.Drawing.Point(480, 20);
            this.cmbLogLevel.Name = "cmbLogLevel";
            this.cmbLogLevel.Size = new System.Drawing.Size(100, 23);
            this.cmbLogLevel.TabIndex = 7;
            this.cmbLogLevel.SelectedIndexChanged += new System.EventHandler(this.cmbLogLevel_SelectedIndexChanged);

            // ---------- Очистка лога ----------
            this.btnClearLog.Name = "btnClearLog";
            this.btnClearLog.Text = "Очистить лог";
            this.btnClearLog.Size = new System.Drawing.Size(120, 23);
            this.btnClearLog.Location = new System.Drawing.Point(588, 20);
            this.btnClearLog.TabIndex = 8;
            this.btnClearLog.UseVisualStyleBackColor = true;
            this.btnClearLog.Click += new System.EventHandler(this.btnClearLog_Click);

            // ---------- Экспорт лога ----------
            this.btnExportLog.Name = "btnExportLog";
            this.btnExportLog.Text = "Экспорт лога";
            this.btnExportLog.Size = new System.Drawing.Size(110, 23);
            this.btnExportLog.Location = new System.Drawing.Point(712, 20);
            this.btnExportLog.TabIndex = 11;
            this.btnExportLog.UseVisualStyleBackColor = true;
            this.btnExportLog.Click += new System.EventHandler(this.btnExportLog_Click);

            // ---------- Фильтр лога ----------
            this.labelLogFilter.AutoSize = true;
            this.labelLogFilter.Location = new System.Drawing.Point(12, 52);
            this.labelLogFilter.Name = "labelLogFilter";
            this.labelLogFilter.Size = new System.Drawing.Size(86, 15);
            this.labelLogFilter.TabIndex = 202;
            this.labelLogFilter.Text = "Фильтр лога:";

            this.clbLogFilter.CheckOnClick = true;
            this.clbLogFilter.FormattingEnabled = true;
            this.clbLogFilter.Items.AddRange(new object[] { "INFO", "WARN", "ERROR" });
            this.clbLogFilter.Location = new System.Drawing.Point(104, 48);
            this.clbLogFilter.Name = "clbLogFilter";
            this.clbLogFilter.Size = new System.Drawing.Size(210, 58);
            this.clbLogFilter.TabIndex = 12;
            this.clbLogFilter.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.clbLogFilter_ItemCheck);
            // ---------- Статус‑бар ----------
            this.statusStrip1.Dock = DockStyle.Bottom;
            this.statusStrip1.Items.AddRange(new ToolStripItem[] {
                this.statusLabelAction,
                this.statusLabelElapsed,
                this.statusLabelSpeed,
                this.statusLabelEta
            });
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.TabIndex = 300;

            this.statusLabelAction.Name = "statusLabelAction";
            this.statusLabelAction.Text = "Готово";
            this.statusLabelAction.Spring = true;

            this.statusLabelElapsed.Name = "statusLabelElapsed";
            this.statusLabelElapsed.Text = "Время: 00:00";

            this.statusLabelSpeed.Name = "statusLabelSpeed";
            this.statusLabelSpeed.Text = "Скорость: 0/с";

            this.statusLabelEta.Name = "statusLabelEta";
            this.statusLabelEta.Text = "ETA: --:--";

            // ---------- Профили ----------
            this.labelProfile.AutoSize = true;
            this.labelProfile.Location = new System.Drawing.Point(12, 126);
            this.labelProfile.Name = "labelProfile";
            this.labelProfile.Size = new System.Drawing.Size(59, 15);
            this.labelProfile.TabIndex = 160;
            this.labelProfile.Text = "Профиль:";

            this.cmbProfile.DropDownStyle = ComboBoxStyle.DropDownList;
            this.cmbProfile.FormattingEnabled = true;
            this.cmbProfile.Location = new System.Drawing.Point(116, 122);
            this.cmbProfile.Name = "cmbProfile";
            this.cmbProfile.Size = new System.Drawing.Size(280, 23);
            this.cmbProfile.TabIndex = 4;
            this.cmbProfile.SelectedIndexChanged += new System.EventHandler(this.cmbProfile_SelectedIndexChanged);

            this.btnSaveProfile.Name = "btnSaveProfile";
            this.btnSaveProfile.Text = "Сохранить профиль";
            this.btnSaveProfile.Size = new System.Drawing.Size(150, 23);
            this.btnSaveProfile.Location = new System.Drawing.Point(406, 122);
            this.btnSaveProfile.TabIndex = 5;
            this.btnSaveProfile.UseVisualStyleBackColor = true;
            this.btnSaveProfile.Click += new System.EventHandler(this.btnSaveProfile_Click);

            this.btnDeleteProfile.Name = "btnDeleteProfile";
            this.btnDeleteProfile.Text = "Удалить профиль";
            this.btnDeleteProfile.Size = new System.Drawing.Size(140, 23);
            this.btnDeleteProfile.Location = new System.Drawing.Point(562, 122);
            this.btnDeleteProfile.TabIndex = 6;
            this.btnDeleteProfile.UseVisualStyleBackColor = true;
            this.btnDeleteProfile.Click += new System.EventHandler(this.btnDeleteProfile_Click);

            // ---------- Показать лог как файл ----------
            this.btnShowLogFile.Name = "btnShowLogFile";
            this.btnShowLogFile.Text = "Показать лог";
            this.btnShowLogFile.Size = new System.Drawing.Size(110, 23);
            this.btnShowLogFile.Location = new System.Drawing.Point(150, 20);
            this.btnShowLogFile.TabIndex = 9;
            this.btnShowLogFile.UseVisualStyleBackColor = true;
            this.btnShowLogFile.Click += new System.EventHandler(this.btnShowLogFile_Click);

            // ---------- Открыть последнюю папку отчёта/лога ----------
            this.btnOpenLastFolder.Name = "btnOpenLastFolder";
            this.btnOpenLastFolder.Text = "Открыть папку отчёта/лога";
            this.btnOpenLastFolder.Size = new System.Drawing.Size(210, 23);
            this.btnOpenLastFolder.Location = new System.Drawing.Point(266, 20);
            this.btnOpenLastFolder.TabIndex = 10;
            this.btnOpenLastFolder.UseVisualStyleBackColor = true;
            this.btnOpenLastFolder.Click += new System.EventHandler(this.btnOpenLastFolder_Click);

            // Единые размеры и отступы для «больших» кнопок
            Size btnSize = new Size(200, 34);
            Padding btnMargin = new Padding(6, 6, 6, 6);

            // btnConvert
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Text = "Конвертировать";
            this.btnConvert.Size = btnSize;
            this.btnConvert.Margin = btnMargin;
            this.btnConvert.TabIndex = 10;
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);

            // button5 — Создать базу данных
            this.button5.Name = "button5";
            this.button5.Text = "Создать базу данных";
            this.button5.Size = btnSize;
            this.button5.Margin = btnMargin;
            this.button5.TabIndex = 11;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);

            // btnCreateSecondDatabase
            this.btnCreateSecondDatabase.Name = "btnCreateSecondDatabase";
            this.btnCreateSecondDatabase.Text = "Создать вторую БД";
            this.btnCreateSecondDatabase.Size = btnSize;
            this.btnCreateSecondDatabase.Margin = btnMargin;
            this.btnCreateSecondDatabase.TabIndex = 12;
            this.btnCreateSecondDatabase.UseVisualStyleBackColor = true;
            this.btnCreateSecondDatabase.Click += new System.EventHandler(this.btnCreateSecondDatabase_Click);

            // btnCreateDuplicatedDatabase
            this.btnCreateDuplicatedDatabase.Name = "btnCreateDuplicatedDatabase";
            this.btnCreateDuplicatedDatabase.Text = "Создать БД (дублирование)";
            this.btnCreateDuplicatedDatabase.Size = btnSize;
            this.btnCreateDuplicatedDatabase.Margin = btnMargin;
            this.btnCreateDuplicatedDatabase.TabIndex = 13;
            this.btnCreateDuplicatedDatabase.UseVisualStyleBackColor = true;
            this.btnCreateDuplicatedDatabase.Click += new System.EventHandler(this.btnCreateDuplicatedDatabase_Click);

            // btnParseDpl
            this.btnParseDpl.Name = "btnParseDpl";
            this.btnParseDpl.Text = "Парсить DPL → DB";
            this.btnParseDpl.Size = btnSize;
            this.btnParseDpl.Margin = btnMargin;
            this.btnParseDpl.TabIndex = 14;
            this.btnParseDpl.UseVisualStyleBackColor = true;
            this.btnParseDpl.Click += new System.EventHandler(this.btnParseDpl_Click);

            // btnExportByBuilding — для экспорта CSV по блокам (00/10/20/30/40)
            this.btnExportByBuilding.Name = "btnExportByBuilding";
            this.btnExportByBuilding.Text = "CSV по блокам (00/10/20/30/40)";
            this.btnExportByBuilding.Size = btnSize;
            this.btnExportByBuilding.Margin = btnMargin;
            this.btnExportByBuilding.TabIndex = 15;
            this.btnExportByBuilding.UseVisualStyleBackColor = true;
            this.btnExportByBuilding.Click += new System.EventHandler(this.btnExportByBuilding_Click);

            // btnOpenOutput — открыть папку вывода
            this.btnOpenOutput.Name = "btnOpenOutput";
            this.btnOpenOutput.Text = "Открыть папку вывода";
            this.btnOpenOutput.Size = btnSize;
            this.btnOpenOutput.Margin = btnMargin;
            this.btnOpenOutput.TabIndex = 16;
            this.btnOpenOutput.UseVisualStyleBackColor = true;
            this.btnOpenOutput.Click += new System.EventHandler(this.btnOpenOutput_Click);

            // button4 — О программе
            this.button4.Name = "button4";
            this.button4.Text = "О программе";
            this.button4.Size = btnSize;
            this.button4.Margin = btnMargin;
            this.button4.TabIndex = 17;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.btnInfo_Click);

            // btnCancel — отмена операции
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Text = "Отмена";
            this.btnCancel.Size = btnSize;
            this.btnCancel.Margin = btnMargin;
            this.btnCancel.TabIndex = 18;
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Enabled = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);

            // Добавляем кнопки в панель
            this.panelButtons.Controls.Add(this.btnConvert);
            this.panelButtons.Controls.Add(this.button5);
            this.panelButtons.Controls.Add(this.btnCreateSecondDatabase);
            this.panelButtons.Controls.Add(this.btnCreateDuplicatedDatabase);
            this.panelButtons.Controls.Add(this.btnParseDpl);
            this.panelButtons.Controls.Add(this.btnExportByBuilding);
            this.panelButtons.Controls.Add(this.btnOpenOutput);
            this.panelButtons.Controls.Add(this.button4);
            this.panelButtons.Controls.Add(this.btnCancel);

            // ---------- Добавляем на форму ----------
            this.groupInput.Controls.Add(this.label1);
            this.groupInput.Controls.Add(this.label2);
            this.groupInput.Controls.Add(this.txtExcelFilePath);
            this.groupInput.Controls.Add(this.txtOutputDirectory);
            this.groupInput.Controls.Add(this.button1);
            this.groupInput.Controls.Add(this.button3);
            this.groupInput.Controls.Add(this.progressBar1);
            this.groupInput.Controls.Add(this.labelProgress);
            this.groupInput.Controls.Add(this.labelIssues);
            this.groupInput.Controls.Add(this.labelProgressDetail);
            this.groupInput.Controls.Add(this.labelProfile);
            this.groupInput.Controls.Add(this.cmbProfile);
            this.groupInput.Controls.Add(this.btnSaveProfile);
            this.groupInput.Controls.Add(this.btnDeleteProfile);

            this.groupActions.Controls.Add(this.panelButtons);

            this.groupLog.Controls.Add(this.labelLog);
            this.groupLog.Controls.Add(this.labelLogLevel);
            this.groupLog.Controls.Add(this.cmbLogLevel);
            this.groupLog.Controls.Add(this.btnClearLog);
            this.groupLog.Controls.Add(this.rtbLog);
            this.groupLog.Controls.Add(this.btnShowLogFile);
            this.groupLog.Controls.Add(this.btnOpenLastFolder);
            this.groupLog.Controls.Add(this.btnExportLog);
            this.groupLog.Controls.Add(this.labelLogFilter);
            this.groupLog.Controls.Add(this.clbLogFilter);

            this.Controls.Add(this.groupInput);
            this.Controls.Add(this.groupActions);
            this.Controls.Add(this.groupLog);
            this.Controls.Add(this.statusStrip1);

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
