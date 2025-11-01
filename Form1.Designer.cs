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

            this.SuspendLayout();

            // ---------- Форма ----------
            // Базовые параметры масштабирования — возвращаем к 96 dpi,
            // чтобы не "раздувало" контролы при повторном открытии в дизайнере.
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point);
            this.ClientSize = new System.Drawing.Size(920, 420);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Converter";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "EXCEL → CSV / DB — Converter";

            // ---------- Метки ----------
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 15);
            this.label1.TabIndex = 100;
            this.label1.Text = "Excel-файл:";

            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(16, 58);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 15);
            this.label2.TabIndex = 101;
            this.label2.Text = "Папка вывода:";

            // ---------- Поля ввода ----------
            this.txtExcelFilePath.Location = new System.Drawing.Point(120, 14);
            this.txtExcelFilePath.Name = "txtExcelFilePath";
            this.txtExcelFilePath.Size = new System.Drawing.Size(650, 23);
            this.txtExcelFilePath.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtExcelFilePath.TabIndex = 0;

            this.txtOutputDirectory.Location = new System.Drawing.Point(120, 54);
            this.txtOutputDirectory.Name = "txtOutputDirectory";
            this.txtOutputDirectory.Size = new System.Drawing.Size(650, 23);
            this.txtOutputDirectory.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            this.txtOutputDirectory.TabIndex = 2;
            this.txtOutputDirectory.TextChanged += new System.EventHandler(this.txtOutputDirectory_TextChanged);

            // ---------- Кнопки рядом с полями ----------
            this.button1.Location = new System.Drawing.Point(780, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(120, 27);
            this.button1.TabIndex = 1;
            this.button1.Text = "Выбрать…";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.button1.Click += new System.EventHandler(this.btnSelectExcelFile_Click);

            this.button3.Location = new System.Drawing.Point(780, 52);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(120, 27);
            this.button3.TabIndex = 3;
            this.button3.Text = "Папка…";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            this.button3.Click += new System.EventHandler(this.btnSelectOutputDirectory_Click);

            // ---------- Прогрессбар ----------
            this.progressBar1.Location = new System.Drawing.Point(16, 96);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(884, 18);
            this.progressBar1.TabIndex = 4;
            this.progressBar1.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // ---------- Панель кнопок (ровная сетка) ----------
            this.panelButtons.Location = new System.Drawing.Point(16, 126);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Size = new System.Drawing.Size(884, 270);
            this.panelButtons.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
            this.panelButtons.AutoSize = false;
            this.panelButtons.WrapContents = true;
            this.panelButtons.FlowDirection = FlowDirection.LeftToRight;
            this.panelButtons.Padding = new Padding(0);
            this.panelButtons.Margin = new Padding(0);
            this.panelButtons.TabIndex = 5;

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

            // button4 — О программе
            this.button4.Name = "button4";
            this.button4.Text = "О программе";
            this.button4.Size = btnSize;
            this.button4.Margin = btnMargin;
            this.button4.TabIndex = 16;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.btnInfo_Click);

            // Добавляем кнопки в панель
            this.panelButtons.Controls.Add(this.btnConvert);
            this.panelButtons.Controls.Add(this.button5);
            this.panelButtons.Controls.Add(this.btnCreateSecondDatabase);
            this.panelButtons.Controls.Add(this.btnCreateDuplicatedDatabase);
            this.panelButtons.Controls.Add(this.btnParseDpl);
            this.panelButtons.Controls.Add(this.btnExportByBuilding);
            this.panelButtons.Controls.Add(this.button4);

            // ---------- Добавляем на форму ----------
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtExcelFilePath);
            this.Controls.Add(this.txtOutputDirectory);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.panelButtons);

            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion
    }
}
