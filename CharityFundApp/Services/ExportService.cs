using System;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace CharityFundApp.Services
{
    /// <summary>
    /// Сервис экспорта табличных данных в форматы CSV (для Excel) и HTML (для печати/просмотра)
    /// </summary>
    public static class ExportService
    {
        /// <summary>
        /// Экспорт данных DataGridView в CSV с кодировкой UTF-8 BOM для безупречного открытия в Excel
        /// </summary>
        public static void ExportToCsv(DataGridView dgv, string defaultFileName)
        {
            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта!", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Файл CSV (*.csv)|*.csv|Все файлы (*.*)|*.*",
                FileName = $"{defaultFileName}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                var sb = new StringBuilder();

                // Заголовки видимых колонок
                var headers = new System.Collections.Generic.List<string>();
                var visibleCols = new System.Collections.Generic.List<int>();

                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (dgv.Columns[i].Visible)
                    {
                        headers.Add(EscapeCsv(dgv.Columns[i].HeaderText));
                        visibleCols.Add(i);
                    }
                }
                sb.AppendLine(string.Join(";", headers));

                // Строки данных
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    var fields = new System.Collections.Generic.List<string>();
                    foreach (int colIdx in visibleCols)
                    {
                        var val = row.Cells[colIdx].Value;
                        string text = val?.ToString() ?? "";
                        if (val is DateTime dt)
                        {
                            text = dt.ToString("dd.MM.yyyy HH:mm");
                        }
                        else if (val is decimal dec)
                        {
                            text = dec.ToString("N2");
                        }
                        fields.Add(EscapeCsv(text));
                    }
                    sb.AppendLine(string.Join(";", fields));
                }

                // Запись с UTF-8 BOM
                File.WriteAllText(sfd.FileName, sb.ToString(), new UTF8Encoding(true));
                MessageBox.Show("Данные успешно экспортированы в CSV!", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Экспорт DataGridView в красивый печатный HTML-документ
        /// </summary>
        public static void ExportToHtml(DataGridView dgv, string title, string defaultFileName)
        {
            if (dgv.Rows.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта!", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using var sfd = new SaveFileDialog
            {
                Filter = "Веб-страница HTML (*.html)|*.html|Все файлы (*.*)|*.*",
                FileName = $"{defaultFileName}_{DateTime.Now:yyyyMMdd_HHmmss}.html"
            };

            if (sfd.ShowDialog() != DialogResult.OK) return;

            try
            {
                var sb = new StringBuilder();
                sb.AppendLine("<!DOCTYPE html>");
                sb.AppendLine("<html lang=\"ru\">");
                sb.AppendLine("<head>");
                sb.AppendLine("  <meta charset=\"UTF-8\">");
                sb.AppendLine($"  <title>{title}</title>");
                sb.AppendLine("  <style>");
                sb.AppendLine("    body { font-family: 'Segoe UI', Tahoma, sans-serif; margin: 30px; color: #222; }");
                sb.AppendLine("    h2 { color: #1e3a8a; border-bottom: 2px solid #3b82f6; padding-bottom: 8px; }");
                sb.AppendLine("    .meta { font-size: 13px; color: #64748b; margin-bottom: 20px; }");
                sb.AppendLine("    table { border-collapse: collapse; width: 100%; margin-top: 15px; }");
                sb.AppendLine("    th, td { border: 1px solid #cbd5e1; padding: 10px 12px; text-align: left; }");
                sb.AppendLine("    th { background-color: #f1f5f9; color: #1e293b; font-weight: 600; }");
                sb.AppendLine("    tr:nth-child(even) { background-color: #f8fafc; }");
                sb.AppendLine("    tr:hover { background-color: #e2e8f0; }");
                sb.AppendLine("    .footer { margin-top: 30px; font-size: 12px; color: #94a3b8; text-align: right; }");
                sb.AppendLine("  </style>");
                sb.AppendLine("</head>");
                sb.AppendLine("<body>");
                sb.AppendLine($"  <h2>{title}</h2>");
                sb.AppendLine($"  <div class=\"meta\">Дата формирования: {DateTime.Now:dd.MM.yyyy HH:mm:ss} | ИС «Благотворительный фонд»</div>");
                sb.AppendLine("  <table>");
                sb.AppendLine("    <thead><tr>");

                var visibleCols = new System.Collections.Generic.List<int>();
                for (int i = 0; i < dgv.Columns.Count; i++)
                {
                    if (dgv.Columns[i].Visible)
                    {
                        sb.AppendLine($"      <th>{System.Net.WebUtility.HtmlEncode(dgv.Columns[i].HeaderText)}</th>");
                        visibleCols.Add(i);
                    }
                }
                sb.AppendLine("    </tr></thead>");
                sb.AppendLine("    <tbody>");

                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.IsNewRow) continue;
                    sb.AppendLine("    <tr>");
                    foreach (int colIdx in visibleCols)
                    {
                        var val = row.Cells[colIdx].Value;
                        string text = val?.ToString() ?? "";
                        if (val is DateTime dt) text = dt.ToString("dd.MM.yyyy HH:mm");
                        else if (val is decimal dec) text = dec.ToString("N2");
                        sb.AppendLine($"      <td>{System.Net.WebUtility.HtmlEncode(text)}</td>");
                    }
                    sb.AppendLine("    </tr>");
                }

                sb.AppendLine("    </tbody>");
                sb.AppendLine("  </table>");
                sb.AppendLine("  <div class=\"footer\">Разработчик: Афанасьев М.А., ПГУ им. Т.Г. Шевченко, 2026 г.</div>");
                sb.AppendLine("</body>");
                sb.AppendLine("</html>");

                File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                MessageBox.Show("HTML-отчет успешно сформирован!", "Экспорт", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте HTML: {ex.Message}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string EscapeCsv(string text)
        {
            if (string.IsNullOrEmpty(text)) return "\"\"";
            if (text.Contains(";") || text.Contains("\"") || text.Contains("\n") || text.Contains("\r"))
            {
                text = text.Replace("\"", "\"\"");
                return $"\"{text}\"";
            }
            return text;
        }
    }
}
