using System.Drawing;
using System.Windows.Forms;
using GUI_EXCEL_parser.Properties;

namespace GUI_EXCEL_parser
{
    internal sealed class ReportForm : Form
    {
        private readonly RichTextBox _reportBox;
        private readonly Button _btnSave;
        private readonly Button _btnCopy;

        public ReportForm(string title, string content, Color backColor, Color foreColor)
        {
            Text = title;
            StartPosition = FormStartPosition.CenterParent;
            Width = 720;
            Height = 520;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            BackColor = backColor;
            ForeColor = foreColor;

            _btnSave = new Button
            {
                Text = "Сохранить отчёт в файл",
                Dock = DockStyle.Top,
                Height = 30,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat
            };
            _btnSave.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
            _btnSave.Click += (s, e) => SaveReportToFile(title, content);

            _btnCopy = new Button
            {
                Text = "Копировать отчёт",
                Dock = DockStyle.Top,
                Height = 30,
                BackColor = Color.FromArgb(45, 45, 48),
                ForeColor = foreColor,
                FlatStyle = FlatStyle.Flat
            };
            _btnCopy.FlatAppearance.BorderColor = Color.FromArgb(70, 70, 70);
            _btnCopy.Click += (s, e) => CopyReportToClipboard(content);

            _reportBox = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                WordWrap = false,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                BackColor = Color.FromArgb(24, 24, 24),
                ForeColor = foreColor,
                BorderStyle = BorderStyle.FixedSingle,
                Text = content
            };

            Controls.Add(_reportBox);
            Controls.Add(_btnCopy);
            Controls.Add(_btnSave);
        }

        private void SaveReportToFile(string title, string content)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Filter = "Текстовый файл (*.txt)|*.txt|Все файлы (*.*)|*.*";
                dialog.FileName = $"{title}_{System.DateTime.Now:yyyyMMdd_HHmmss}.txt";
                dialog.OverwritePrompt = true;
                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                System.IO.File.WriteAllText(dialog.FileName, content, System.Text.Encoding.UTF8);
                Settings.Default.LastReportDirectory = System.IO.Path.GetDirectoryName(dialog.FileName);
                Settings.Default.Save();
            }
        }

        private void CopyReportToClipboard(string content)
        {
            try
            {
                Clipboard.SetText(content);
            }
            catch
            {
                // Копирование в буфер не должно ломать окно отчёта.
            }
        }
    }
}
