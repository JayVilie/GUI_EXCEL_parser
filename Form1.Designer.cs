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
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 10);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Введите путь к Excel файлу:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 55);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(296, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Введите путь к директории для сохранения CSV файлов:";
            // 
            // txtExcelFilePath
            // 
            this.txtExcelFilePath.Location = new System.Drawing.Point(12, 26);
            this.txtExcelFilePath.Name = "txtExcelFilePath";
            this.txtExcelFilePath.Size = new System.Drawing.Size(531, 20);
            this.txtExcelFilePath.TabIndex = 2;
            // 
            // txtOutputDirectory
            // 
            this.txtOutputDirectory.Location = new System.Drawing.Point(12, 71);
            this.txtOutputDirectory.Name = "txtOutputDirectory";
            this.txtOutputDirectory.Size = new System.Drawing.Size(531, 20);
            this.txtOutputDirectory.TabIndex = 3;
            this.txtOutputDirectory.TextChanged += new System.EventHandler(this.txtOutputDirectory_TextChanged);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(549, 21);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(73, 29);
            this.button1.TabIndex = 4;
            this.button1.Text = "Выбрать";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.btnSelectExcelFile_Click);
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(12, 108);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(114, 30);
            this.btnConvert.TabIndex = 5;
            this.btnConvert.Text = "Конвертировать";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(549, 66);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(73, 29);
            this.button3.TabIndex = 7;
            this.button3.Text = "Выбрать";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.btnSelectOutputDirectory_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(132, 108);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(73, 29);
            this.button4.TabIndex = 8;
            this.button4.Text = "Инфо";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.btnInfo_Click);
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(235, 108);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(114, 30);
            this.button5.TabIndex = 9;
            this.button5.Text = "SQLite";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // btnCreateSecondDatabase
            // 
            this.btnCreateSecondDatabase.Location = new System.Drawing.Point(367, 107);
            this.btnCreateSecondDatabase.Name = "btnCreateSecondDatabase";
            this.btnCreateSecondDatabase.Size = new System.Drawing.Size(114, 30);
            this.btnCreateSecondDatabase.TabIndex = 10;
            this.btnCreateSecondDatabase.Text = "SQLite 2";
            this.btnCreateSecondDatabase.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.btnCreateSecondDatabase.UseVisualStyleBackColor = true;
            this.btnCreateSecondDatabase.Click += new System.EventHandler(this.btnCreateSecondDatabase_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 144);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(610, 23);
            this.progressBar1.TabIndex = 11;
            // 
            // btnCreateDuplicatedDatabase
            // 
            this.btnCreateDuplicatedDatabase.Location = new System.Drawing.Point(487, 107);
            this.btnCreateDuplicatedDatabase.Name = "btnCreateDuplicatedDatabase";
            this.btnCreateDuplicatedDatabase.Size = new System.Drawing.Size(135, 30);
            this.btnCreateDuplicatedDatabase.TabIndex = 12;
            this.btnCreateDuplicatedDatabase.Text = "CreateDuplicatedDatabase";
            this.btnCreateDuplicatedDatabase.UseVisualStyleBackColor = true;
            this.btnCreateDuplicatedDatabase.Click += new System.EventHandler(this.btnCreateDuplicatedDatabase_Click);
            // 
            // Converter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(639, 180);
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
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Converter";
            this.Text = "Converter";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

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
    }
}
