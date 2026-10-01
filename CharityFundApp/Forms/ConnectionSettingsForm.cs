using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Data.SqlClient;
using CharityFundApp.DataAccess;

namespace CharityFundApp.Forms
{
    public class ConnectionSettingsForm : Form
    {
        private readonly ComboBox _cmbPresets;
        private readonly TextBox _txtConnString;
        private readonly Button _btnTest;
        private readonly Button _btnInitDb;
        private readonly Label _lblStatus;
        private readonly Button _btnSave;
        private readonly Button _btnCancel;

        public ConnectionSettingsForm()
        {
            Text = "Параметры подключения к MS SQL Server";
            Size = new Size(620, 420);
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9.5f);

            int y = 20;

            var lblPreset = new Label { Text = "Шаблоны подключения (пресеты):", Location = new Point(20, y), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };
            y += 25;
            _cmbPresets = new ComboBox { Location = new Point(20, y), Width = 560, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbPresets.Items.AddRange(new object[] {
                "1. Локальный сервер по умолчанию (localhost)",
                "2. Сервер SQL Server Express (.\\SQLEXPRESS)",
                "3. Локальная база Visual Studio ((localdb)\\MSSQLLocalDB)",
                "4. Автономный режим (Встроенная база данных SQLite — без установки сервера)",
                "5. Пользовательская строка подключения"
            });
            _cmbPresets.SelectedIndex = DatabaseHelper.CurrentProvider == DatabaseProvider.Sqlite ? 3 : 0;
            _cmbPresets.SelectedIndexChanged += CmbPresets_SelectedIndexChanged;
            y += 35;

            var lblConn = new Label { Text = "Строка подключения (Connection String):", Location = new Point(20, y), AutoSize = true };
            y += 25;
            _txtConnString = new TextBox { Location = new Point(20, y), Width = 560, Height = 65, Multiline = true, Text = DatabaseHelper.ConnectionString };
            y += 75;

            _btnTest = new Button { Text = "🔌 Проверить подключение", Location = new Point(20, y), Width = 230, Height = 35, BackColor = Color.FromArgb(241, 245, 249), FlatStyle = FlatStyle.Flat };
            _btnTest.Click += BtnTest_Click;

            _btnInitDb = new Button { Text = "⚡ Развернуть БД из скрипта", Location = new Point(260, y), Width = 230, Height = 35, BackColor = Color.FromArgb(254, 243, 199), FlatStyle = FlatStyle.Flat };
            _btnInitDb.Click += BtnInitDb_Click;
            y += 45;

            _lblStatus = new Label { Text = "Нажмите «Проверить подключение» для проверки соединения с сервером", Location = new Point(20, y), Width = 560, Height = 45, ForeColor = Color.FromArgb(71, 85, 105) };
            y += 55;

            _btnSave = new Button { Text = "Применить", Location = new Point(360, y), Width = 105, Height = 34, DialogResult = DialogResult.OK, BackColor = Color.FromArgb(37, 99, 235), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            _btnCancel = new Button { Text = "Закрыть", Location = new Point(475, y), Width = 105, Height = 34, DialogResult = DialogResult.Cancel };

            _btnSave.Click += (s, e) =>
            {
                if (_cmbPresets.SelectedIndex == 3 || _txtConnString.Text.Contains(".sqlite") || _txtConnString.Text.Contains("Data Source="))
                {
                    DatabaseHelper.CurrentProvider = DatabaseProvider.Sqlite;
                    DatabaseHelper.EnsureSqliteDatabaseCreated();
                }
                else
                {
                    DatabaseHelper.CurrentProvider = DatabaseProvider.SqlServer;
                    DatabaseHelper.ConnectionString = _txtConnString.Text.Trim();
                }
            };

            Controls.AddRange(new Control[] {
                lblPreset, _cmbPresets, lblConn, _txtConnString,
                _btnTest, _btnInitDb, _lblStatus, _btnSave, _btnCancel
            });

            AcceptButton = _btnSave;
            CancelButton = _btnCancel;
        }

        private void CmbPresets_SelectedIndexChanged(object? sender, EventArgs e)
        {
            switch (_cmbPresets.SelectedIndex)
            {
                case 0:
                    DatabaseHelper.CurrentProvider = DatabaseProvider.SqlServer;
                    _txtConnString.Text = "Server=localhost;Database=CharityFundDB;Trusted_Connection=True;TrustServerCertificate=True;";
                    break;
                case 1:
                    DatabaseHelper.CurrentProvider = DatabaseProvider.SqlServer;
                    _txtConnString.Text = "Server=.\\SQLEXPRESS;Database=CharityFundDB;Integrated Security=True;TrustServerCertificate=True;";
                    break;
                case 2:
                    DatabaseHelper.CurrentProvider = DatabaseProvider.SqlServer;
                    _txtConnString.Text = "Server=(localdb)\\MSSQLLocalDB;Database=CharityFundDB;Integrated Security=True;TrustServerCertificate=True;";
                    break;
                case 3:
                    DatabaseHelper.CurrentProvider = DatabaseProvider.Sqlite;
                    _txtConnString.Text = $"Data Source={DatabaseHelper.SqliteDbPath}";
                    break;
            }
        }

        private void BtnTest_Click(object? sender, EventArgs e)
        {
            _lblStatus.Text = "Выполняется проверка соединения...";
            _lblStatus.ForeColor = Color.DarkGoldenrod;
            Application.DoEvents();

            var prevConn = DatabaseHelper.ConnectionString;
            DatabaseHelper.ConnectionString = _txtConnString.Text.Trim();

            var (success, message) = DatabaseHelper.TestConnection();
            if (success)
            {
                _lblStatus.Text = $"✅ {message}";
                _lblStatus.ForeColor = Color.ForestGreen;
            }
            else
            {
                _lblStatus.Text = $"❌ {message}";
                _lblStatus.ForeColor = Color.Crimson;
                DatabaseHelper.ConnectionString = prevConn;
            }
        }

        private void BtnInitDb_Click(object? sender, EventArgs e)
        {
            string scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sql", "CharityFund_CreateDB.sql");
            if (!File.Exists(scriptPath))
            {
                // Поиск в каталоге проекта
                scriptPath = Path.Combine(Directory.GetCurrentDirectory(), "Sql", "CharityFund_CreateDB.sql");
            }

            if (!File.Exists(scriptPath))
            {
                using var ofd = new OpenFileDialog
                {
                    Filter = "SQL Files (*.sql)|*.sql|All files (*.*)|*.*",
                    Title = "Выберите файл CharityFund_CreateDB.sql"
                };
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    scriptPath = ofd.FileName;
                }
                else return;
            }

            try
            {
                _lblStatus.Text = "Выполняется развертывание схемы и тестовых данных...";
                _lblStatus.ForeColor = Color.DarkGoldenrod;
                Application.DoEvents();

                string fullScript = File.ReadAllText(scriptPath);

                // Базовое подключение к master для создания БД если ее еще нет
                var builder = new SqlConnectionStringBuilder(_txtConnString.Text.Trim())
                {
                    InitialCatalog = "master"
                };

                using var conn = new SqlConnection(builder.ConnectionString);
                conn.Open();

                // Разделение по GO
                var commands = fullScript.Split(new[] { "\r\nGO", "\nGO", "\rGO", "\ngo", "\r\ngo" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var cmdText in commands)
                {
                    if (string.IsNullOrWhiteSpace(cmdText)) continue;
                    using var cmd = new SqlCommand(cmdText, conn);
                    cmd.ExecuteNonQuery();
                }

                _lblStatus.Text = "✅ База данных CharityFundDB успешно инициализирована!";
                _lblStatus.ForeColor = Color.ForestGreen;
                MessageBox.Show("База данных успешно создана и наполнена демонстрационными данными!", "Инициализация БД", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                _lblStatus.Text = $"❌ Ошибка развертывания: {ex.Message}";
                _lblStatus.ForeColor = Color.Crimson;
                MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка инициализации", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
