// Converter.Designer.cs
using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace GUI_EXCEL_parser
{
    partial class Converter
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtExcelFilePath;
        private System.Windows.Forms.TextBox txtOutputDirectory;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button btnCreateSecondDatabase;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnCreateDuplicatedDatabase;
        private System.Windows.Forms.Button btnParseDpl;

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Converter));
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
            this.SuspendLayout();
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // txtExcelFilePath
            // 
            resources.ApplyResources(this.txtExcelFilePath, "txtExcelFilePath");
            this.txtExcelFilePath.Name = "txtExcelFilePath";
            // 
            // txtOutputDirectory
            // 
            resources.ApplyResources(this.txtOutputDirectory, "txtOutputDirectory");
            this.txtOutputDirectory.Name = "txtOutputDirectory";
            this.txtOutputDirectory.TextChanged += new System.EventHandler(this.txtOutputDirectory_TextChanged);
            // 
            // button1
            // 
            resources.ApplyResources(this.button1, "button1");
            this.button1.Name = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnSelectExcelFile_Click);
            // 
            // btnConvert
            // 
            resources.ApplyResources(this.btnConvert, "btnConvert");
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // button3
            // 
            resources.ApplyResources(this.button3, "button3");
            this.button3.Name = "button3";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.btnSelectOutputDirectory_Click);
            // 
            // button4
            // 
            resources.ApplyResources(this.button4, "button4");
            this.button4.Name = "button4";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.btnInfo_Click);
            // 
            // button5
            // 
            resources.ApplyResources(this.button5, "button5");
            this.button5.Name = "button5";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // btnCreateSecondDatabase
            // 
            resources.ApplyResources(this.btnCreateSecondDatabase, "btnCreateSecondDatabase");
            this.btnCreateSecondDatabase.Name = "btnCreateSecondDatabase";
            this.btnCreateSecondDatabase.UseVisualStyleBackColor = true;
            this.btnCreateSecondDatabase.Click += new System.EventHandler(this.btnCreateSecondDatabase_Click);
            // 
            // progressBar1
            // 
            resources.ApplyResources(this.progressBar1, "progressBar1");
            this.progressBar1.Name = "progressBar1";
            // 
            // btnCreateDuplicatedDatabase
            // 
            resources.ApplyResources(this.btnCreateDuplicatedDatabase, "btnCreateDuplicatedDatabase");
            this.btnCreateDuplicatedDatabase.Name = "btnCreateDuplicatedDatabase";
            this.btnCreateDuplicatedDatabase.UseVisualStyleBackColor = true;
            this.btnCreateDuplicatedDatabase.Click += new System.EventHandler(this.btnCreateDuplicatedDatabase_Click);
            // 
            // btnParseDpl
            // 
            resources.ApplyResources(this.btnParseDpl, "btnParseDpl");
            this.btnParseDpl.Name = "btnParseDpl";
            this.btnParseDpl.UseVisualStyleBackColor = true;
            this.btnParseDpl.Click += new System.EventHandler(this.btnParseDpl_Click);
            // 
            // Converter
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.Controls.Add(this.btnParseDpl);
            this.Controls.Add(this.btnCreateDuplicatedDatabase);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.btnCreateSecondDatabase);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.btnConvert);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtOutputDirectory);
            this.Controls.Add(this.txtExcelFilePath);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Converter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        // ─────────────────────────────────────────────────────────────────
        // ТЕМА ОФОРМЛЕНИЯ: современная серо‑жёлтая палитра, плоские кнопки,
        // жёлтый прогрессбар, повышенная читаемость.
        // НИКАКИХ изменений в размерах/позициях — только цвета/стили.
        // ─────────────────────────────────────────────────────────────────

        // WinAPI для окраски ProgressBar в жёлтый (PBST_PAUSED)
        private const int WM_USER = 0x0400;
        private const int PBM_SETSTATE = WM_USER + 16;
        private const int PBST_NORMAL = 0x0001; // зелёный (по умолчанию)
        private const int PBST_ERROR = 0x0002; // красный
        private const int PBST_PAUSED = 0x0003; // жёлтый

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private void ApplyModernGrayYellowTheme()
        {
            // Палитра (подобрана для контраста и «кинематографического» ощущения)
            var colBgPrimary = Color.FromArgb(32, 32, 36);   // тёмно‑серый фон формы
            var colBgControl = Color.FromArgb(40, 40, 46);   // фон контролов
            var colBgHover = Color.FromArgb(55, 55, 62);   // наведение
            var colBgActive = Color.FromArgb(65, 65, 72);   // нажатие/актив
            var colBorder = Color.FromArgb(120, 110, 60); // приглушённый жёлтый контур
            var colAccent = Color.FromArgb(227, 179, 65); // тёплый жёлтый (акцент)
            var colTextPrimary = Color.FromArgb(230, 220, 180); // мягкий жёлтый текст
            var colTextMuted = Color.FromArgb(190, 180, 140); // вторичный текст

            // Фон формы и антифликер
            this.BackColor = colBgPrimary;
            this.ForeColor = colTextPrimary;
            this.DoubleBuffered = true;

            // Лейблы — только цвет (не меняем размеры/шрифты)
            StyleLabel(label1, colTextPrimary);
            StyleLabel(label2, colTextPrimary);

            // Текстбоксы — тёмный фон, жёлтый текст, тонкая рамка
            StyleTextBox(txtExcelFilePath, colBgControl, colTextPrimary, colBorder);
            StyleTextBox(txtOutputDirectory, colBgControl, colTextPrimary, colBorder);

            // Кнопки — плоские, без изменения размеров
            StyleButton(button1, colBgControl, colTextPrimary, colBorder, colBgHover, colBgActive);
            StyleButton(btnConvert, colBgControl, colTextPrimary, colBorder, colBgHover, colBgActive);
            StyleButton(button3, colBgControl, colTextPrimary, colBorder, colBgHover, colBgActive);
            StyleButton(button4, colBgControl, colTextPrimary, colBorder, colBgHover, colBgActive);
            StyleButton(button5, colBgControl, colTextPrimary, colBorder, colBgHover, colBgActive);
            StyleButton(btnCreateSecondDatabase, colBgControl, colTextPrimary, colBorder, colBgHover, colBgActive);
            StyleButton(btnCreateDuplicatedDatabase, colBgControl, colTextPrimary, colBorder, colBgHover, colBgActive);
            StyleButton(btnParseDpl, colBgControl, colTextPrimary, colBorder, colBgHover, colBgActive);

            // Прогрессбар — режим Continuous + жёлтый state (PBST_PAUSED)
            try
            {
                progressBar1.Style = ProgressBarStyle.Continuous;
                // под Windows 10/11 это переключает цвет в жёлтый
                SendMessage(progressBar1.Handle, PBM_SETSTATE, (IntPtr)PBST_PAUSED, IntPtr.Zero);
                progressBar1.ForeColor = colAccent; // на старых темах может влиять
                progressBar1.BackColor = colBgControl;
            }
            catch { /* игнор если ОС не поддерживает PBM_SETSTATE */ }

            // Мелкие штрихи на форме
            this.FormBorderStyle = FormBorderStyle.FixedSingle; // как было
            this.MaximizeBox = false;                       // как было
        }

        private static void StyleLabel(Label lbl, Color fore)
        {
            if (lbl == null) return;
            lbl.ForeColor = fore;
            lbl.BackColor = Color.Transparent; // чтобы над серым фоном выглядело аккуратно
            // ВАЖНО: шрифт и AutoSize не трогаем, чтобы не менять размер!
        }

        private static void StyleTextBox(TextBox tb, Color back, Color fore, Color border)
        {
            if (tb == null) return;
            tb.BackColor = back;
            tb.ForeColor = fore;
            tb.BorderStyle = BorderStyle.FixedSingle; // тонкая рамка
            // Подсказка: запрет авто‑вставки системного цвета
            tb.HideSelection = false;
        }

        private static void StyleButton(Button btn, Color back, Color fore, Color border, Color hover, Color active)
        {
            if (btn == null) return;
            btn.UseVisualStyleBackColor = false; // позволяем задавать BackColor
            btn.BackColor = back;
            btn.ForeColor = fore;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = border;
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = hover;
            btn.FlatAppearance.MouseDownBackColor = active;

            // Тонкая настройка рендера текста для читаемости
            btn.TextImageRelation = TextImageRelation.ImageBeforeText;
        }


    }
}
