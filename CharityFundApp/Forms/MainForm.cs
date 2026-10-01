using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using CharityFundApp.DataAccess;
using CharityFundApp.Models;
using CharityFundApp.Services;

namespace CharityFundApp.Forms
{
    public class MainForm : Form
    {
        // Репозитории ADO.NET
        private readonly CategoryRepository _categoryRepo = new();
        private readonly DonorRepository _donorRepo = new();
        private readonly ProjectRepository _projectRepo = new();
        private readonly DonationRepository _donationRepo = new();
        private readonly RecipientRepository _recipientRepo = new();
        private readonly ExpenseRepository _expenseRepo = new();
        private readonly ReportRepository _reportRepo = new();

        // Главные элементы управления
        private TabControl _mainTabControl = null!;
        private StatusStrip _statusStrip = null!;
        private ToolStripStatusLabel _lblDbStatus = null!;
        private ToolStripStatusLabel _lblInfo = null!;

        // Таблицы и контролы вкладок
        // Вкладка Доноры
        private DataGridView _dgvDonors = null!;
        private TextBox _txtDonorSearch = null!;

        // Вкладка Проекты
        private DataGridView _dgvProjects = null!;
        private ComboBox _cmbProjectStatus = null!;
        private TextBox _txtProjectSearch = null!;

        // Вкладка Категории
        private DataGridView _dgvCategories = null!;

        // Вкладка Пожертвования
        private DataGridView _dgvDonations = null!;
        private ComboBox _cmbDonationProjectFilter = null!;
        private DateTimePicker _dtpDonationFrom = null!;
        private DateTimePicker _dtpDonationTo = null!;
        private CheckBox _chkDonationDateFilter = null!;
        private Label _lblDonationsTotal = null!;

        // Вкладка Получатели
        private DataGridView _dgvRecipients = null!;
        private ComboBox _cmbRecipientStatus = null!;
        private TextBox _txtRecipientSearch = null!;

        // Вкладка Расходы
        private DataGridView _dgvExpenses = null!;
        private ComboBox _cmbExpenseProjectFilter = null!;
        private ComboBox _cmbExpenseCategoryFilter = null!;
        private Label _lblExpensesTotal = null!;

        // Вкладка Аналитика
        private ComboBox _cmbReportType = null!;
        private DateTimePicker _dtpReportFrom = null!;
        private DateTimePicker _dtpReportTo = null!;
        private CheckBox _chkReportDateFilter = null!;
        private DataGridView _dgvReports = null!;
        private Label _lblReportSummary = null!;

        public MainForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            Text = "Информационная система «Благотворительный фонд» — Афанасьев М.А., гр. ФТ24ДР62ПИ";
            Size = new Size(1180, 760);
            MinimumSize = new Size(1000, 650);
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9.5f);
            Icon = SystemIcons.Application;

            // 1. Верхняя панель заголовка
            var topPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 65,
                BackColor = Color.FromArgb(30, 41, 59), // Slate 800
                Padding = new Padding(15, 8, 15, 8)
            };

            var lblAppTitle = new Label
            {
                Text = "ИНФОРМАЦИОННАЯ СИСТЕМА «БЛАГОТВОРИТЕЛЬНЫЙ ФОНД»",
                Font = new Font("Segoe UI", 12f, FontStyle.Bold),
                ForeColor = Color.White,
                Location = new Point(15, 10),
                AutoSize = true
            };

            var lblAppSubtitle = new Label
            {
                Text = "Студент: Афанасьев М.А. | Кафедра ПОВТ, ФТИ, ГОУ «ПГУ им. Т.Г. Шевченко», 2026 г.",
                Font = new Font("Segoe UI", 8.5f),
                ForeColor = Color.FromArgb(148, 163, 184),
                Location = new Point(16, 36),
                AutoSize = true
            };

            var btnRefreshAll = new Button
            {
                Text = "🔄 Обновить всё",
                Size = new Size(125, 34),
                Location = new Point(880, 16),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(51, 65, 85),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnRefreshAll.Click += (s, e) => RefreshAllTabs();

            var btnDbSettings = new Button
            {
                Text = "⚙ Параметры БД",
                Size = new Size(130, 34),
                Location = new Point(1015, 16),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.FromArgb(37, 99, 235),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnDbSettings.Click += (s, e) =>
            {
                using var dlg = new ConnectionSettingsForm();
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    CheckDbConnection();
                    RefreshAllTabs();
                }
            };

            topPanel.Controls.AddRange(new Control[] { lblAppTitle, lblAppSubtitle, btnRefreshAll, btnDbSettings });

            // 2. Строка состояния
            _statusStrip = new StatusStrip();
            _lblDbStatus = new ToolStripStatusLabel { Text = "Проверка БД..." };
            _lblInfo = new ToolStripStatusLabel { Text = "Готово", Spring = true, TextAlign = ContentAlignment.MiddleRight };
            _statusStrip.Items.AddRange(new ToolStripItem[] { _lblDbStatus, _lblInfo });

            // 3. TabControl
            _mainTabControl = new TabControl
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f),
                Padding = new Point(12, 6)
            };

            _mainTabControl.TabPages.Add(CreateDonorsTab());
            _mainTabControl.TabPages.Add(CreateProjectsTab());
            _mainTabControl.TabPages.Add(CreateCategoriesTab());
            _mainTabControl.TabPages.Add(CreateDonationsTab());
            _mainTabControl.TabPages.Add(CreateRecipientsTab());
            _mainTabControl.TabPages.Add(CreateExpensesTab());
            _mainTabControl.TabPages.Add(CreateReportsTab());

            _mainTabControl.SelectedIndexChanged += (s, e) => RefreshCurrentTab();

            Controls.Add(_mainTabControl);
            Controls.Add(topPanel);
            Controls.Add(_statusStrip);

            Load += (s, e) =>
            {
                CheckDbConnection();
                RefreshAllTabs();
            };
        }

        private void CheckDbConnection()
        {
            var (success, msg) = DatabaseHelper.TestConnection();
            if (success)
            {
                if (DatabaseHelper.CurrentProvider == DatabaseProvider.Sqlite)
                {
                    _lblDbStatus.Text = "● БД: Автономный режим (SQLite)";
                    _lblDbStatus.ForeColor = Color.SteelBlue;
                }
                else
                {
                    _lblDbStatus.Text = "● БД подключена (MS SQL Server)";
                    _lblDbStatus.ForeColor = Color.DarkGreen;
                }
            }
            else
            {
                _lblDbStatus.Text = "● Сервер MS SQL недоступен";
                _lblDbStatus.ForeColor = Color.DarkOrange;

                var res = MessageBox.Show(
                    "Не удалось подключиться к серверу Microsoft SQL Server!\n" +
                    "(Сервер не найден или не запущен на данном компьютере).\n\n" +
                    "Желаете переключиться на Автономный режим (встроенная локальная база данных SQLite)?\n\n" +
                    "✔ Все функции программы (доноры, проекты, взносы, расходы)\n" +
                    "✔ Все 5 аналитических отчетов\n" +
                    "✔ Экспорт в Excel / CSV\n" +
                    "будут полноценно работать прямо сейчас без необходимости установки службы SQL Server!",
                    "Выбор режима базы данных",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (res == DialogResult.Yes)
                {
                    DatabaseHelper.CurrentProvider = DatabaseProvider.Sqlite;
                    DatabaseHelper.EnsureSqliteDatabaseCreated();
                    _lblDbStatus.Text = "● БД: Автономный режим (SQLite)";
                    _lblDbStatus.ForeColor = Color.SteelBlue;
                    RefreshAllTabs();
                }
            }
        }

        #region Вкладка 1: Доноры
        private TabPage CreateDonorsTab()
        {
            var tab = new TabPage("👥 Доноры");
            var pnlTool = new Panel { Dock = DockStyle.Top, Height = 48, BackColor = Color.FromArgb(248, 250, 252) };

            var btnAdd = CreateToolButton("➕ Добавить", Color.FromArgb(22, 163, 74), 10, 8);
            btnAdd.Click += (s, e) =>
            {
                using var dlg = new DonorEditForm();
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        int newId = _donorRepo.Insert(dlg.Donor);
                        if (!string.IsNullOrEmpty(_txtDonorSearch.Text))
                        {
                            _txtDonorSearch.Text = string.Empty;
                        }
                        LoadDonors(newId);
                        _lblInfo.Text = $"✔ Донор «{dlg.Donor.FullName}» успешно добавлен!";
                    }
                    catch (Exception ex) { ShowError(ex, "Добавление донора"); }
                }
            };

            var btnEdit = CreateToolButton("✏ Редактировать", Color.FromArgb(37, 99, 235), 120, 8);
            btnEdit.Click += (s, e) =>
            {
                if (_dgvDonors.SelectedRows.Count == 0) return;
                var donor = _dgvDonors.SelectedRows[0].DataBoundItem as Donor;
                if (donor == null) return;

                using var dlg = new DonorEditForm(donor);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _donorRepo.Update(dlg.Donor);
                        LoadDonors(dlg.Donor.Id);
                        _lblInfo.Text = $"✔ Данные донора «{dlg.Donor.FullName}» обновлены";
                    }
                    catch (Exception ex) { ShowError(ex, "Редактирование донора"); }
                }
            };

            var btnDel = CreateToolButton("🗑 Удалить", Color.FromArgb(220, 38, 38), 255, 8);
            btnDel.Click += (s, e) =>
            {
                if (_dgvDonors.SelectedRows.Count == 0) return;
                var donor = _dgvDonors.SelectedRows[0].DataBoundItem as Donor;
                if (donor == null) return;

                if (MessageBox.Show($"Удалить донора «{donor.FullName}»?\nВсе связанные пожертвования также будут удалены!", "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        _donorRepo.Delete(donor.Id);
                        LoadDonors();
                        _lblInfo.Text = "✔ Донор успешно удален";
                    }
                    catch (Exception ex) { ShowError(ex, "Удаление донора"); }
                }
            };

            var lblSearch = new Label { Text = "Поиск:", Location = new Point(380, 14), AutoSize = true };
            _txtDonorSearch = new TextBox { Location = new Point(435, 11), Width = 220 };
            _txtDonorSearch.TextChanged += (s, e) => LoadDonors();

            var btnCsv = CreateToolButton("📥 CSV", Color.FromArgb(71, 85, 105), 670, 8);
            btnCsv.Click += (s, e) => ExportService.ExportToCsv(_dgvDonors, "Donors");

            var btnHtml = CreateToolButton("🌐 HTML", Color.FromArgb(71, 85, 105), 760, 8);
            btnHtml.Click += (s, e) => ExportService.ExportToHtml(_dgvDonors, "Список благотворителей (доноров)", "Donors_Report");

            pnlTool.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDel, lblSearch, _txtDonorSearch, btnCsv, btnHtml });

            _dgvDonors = CreateStyledGrid();
            tab.Controls.Add(_dgvDonors);
            tab.Controls.Add(pnlTool);

            return tab;
        }

        private void LoadDonors(int? selectId = null)
        {
            try
            {
                var list = _donorRepo.GetAll(_txtDonorSearch.Text);
                _dgvDonors.DataSource = null;
                _dgvDonors.DataSource = list;
                ConfigureGridHeaders(_dgvDonors, new[] {
                    ("Id", "№", 50),
                    ("FullName", "ФИО / Название организации", 240),
                    ("DonorType", "Тип донора", 130),
                    ("Phone", "Телефон", 140),
                    ("Email", "Email", 170),
                    ("Address", "Адрес", 200),
                    ("RegistrationDate", "Дата регистрации", 130)
                });
                _lblInfo.Text = $"Доноров в базе: {list.Count}";
                if (selectId.HasValue) SelectRowById(_dgvDonors, selectId.Value);
            }
            catch (Exception ex) { ShowError(ex, "Загрузка списка доноров"); }
        }
        #endregion

        #region Вкладка 2: Проекты
        private TabPage CreateProjectsTab()
        {
            var tab = new TabPage("📁 Проекты фонда");
            var pnlTool = new Panel { Dock = DockStyle.Top, Height = 48, BackColor = Color.FromArgb(248, 250, 252) };

            var btnAdd = CreateToolButton("➕ Добавить", Color.FromArgb(22, 163, 74), 10, 8);
            btnAdd.Click += (s, e) =>
            {
                var cats = _categoryRepo.GetAll();
                if (cats.Count == 0)
                {
                    MessageBox.Show("Сначала создайте хотя бы одну категорию!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                using var dlg = new ProjectEditForm(null, cats);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        int newId = _projectRepo.Insert(dlg.Project);
                        if (_cmbProjectStatus.SelectedIndex != 0) _cmbProjectStatus.SelectedIndex = 0;
                        if (!string.IsNullOrEmpty(_txtProjectSearch.Text)) _txtProjectSearch.Text = string.Empty;
                        RefreshProjectDropdowns();
                        LoadProjects(newId);
                        _lblInfo.Text = $"✔ Проект «{dlg.Project.Name}» успешно создан!";
                    }
                    catch (Exception ex) { ShowError(ex, "Создание проекта"); }
                }
            };

            var btnEdit = CreateToolButton("✏ Редактировать", Color.FromArgb(37, 99, 235), 120, 8);
            btnEdit.Click += (s, e) =>
            {
                if (_dgvProjects.SelectedRows.Count == 0) return;
                var proj = _dgvProjects.SelectedRows[0].DataBoundItem as Project;
                if (proj == null) return;

                var cats = _categoryRepo.GetAll();
                using var dlg = new ProjectEditForm(proj, cats);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _projectRepo.Update(dlg.Project);
                        RefreshProjectDropdowns();
                        LoadProjects(dlg.Project.Id);
                        _lblInfo.Text = $"✔ Данные проекта «{dlg.Project.Name}» обновлены";
                    }
                    catch (Exception ex) { ShowError(ex, "Редактирование проекта"); }
                }
            };

            var btnDel = CreateToolButton("🗑 Удалить", Color.FromArgb(220, 38, 38), 255, 8);
            btnDel.Click += (s, e) =>
            {
                if (_dgvProjects.SelectedRows.Count == 0) return;
                var proj = _dgvProjects.SelectedRows[0].DataBoundItem as Project;
                if (proj == null) return;

                if (MessageBox.Show($"Удалить проект «{proj.Name}»?\nВсе пожертвования и расходы по проекту будут также удалены!", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        _projectRepo.Delete(proj.Id);
                        RefreshProjectDropdowns();
                        LoadProjects();
                        _lblInfo.Text = "✔ Проект успешно удален";
                    }
                    catch (Exception ex) { ShowError(ex, "Удаление проекта"); }
                }
            };

            var lblStatus = new Label { Text = "Статус:", Location = new Point(370, 14), AutoSize = true };
            _cmbProjectStatus = new ComboBox { Location = new Point(425, 11), Width = 130, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbProjectStatus.Items.AddRange(new object[] { "Все", "Активен", "Завершен", "Приостановлен" });
            _cmbProjectStatus.SelectedIndex = 0;
            _cmbProjectStatus.SelectedIndexChanged += (s, e) => LoadProjects();

            var lblSearch = new Label { Text = "Поиск:", Location = new Point(570, 14), AutoSize = true };
            _txtProjectSearch = new TextBox { Location = new Point(620, 11), Width = 170 };
            _txtProjectSearch.TextChanged += (s, e) => LoadProjects();

            var btnCsv = CreateToolButton("📥 CSV", Color.FromArgb(71, 85, 105), 805, 8);
            btnCsv.Click += (s, e) => ExportService.ExportToCsv(_dgvProjects, "Projects");

            pnlTool.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDel, lblStatus, _cmbProjectStatus, lblSearch, _txtProjectSearch, btnCsv });

            _dgvProjects = CreateStyledGrid();
            tab.Controls.Add(_dgvProjects);
            tab.Controls.Add(pnlTool);

            return tab;
        }

        private void LoadProjects(int? selectId = null)
        {
            try
            {
                var list = _projectRepo.GetAll(_cmbProjectStatus.SelectedItem?.ToString(), _txtProjectSearch.Text);
                _dgvProjects.DataSource = null;
                _dgvProjects.DataSource = list;
                ConfigureGridHeaders(_dgvProjects, new[] {
                    ("Id", "№", 45),
                    ("Name", "Название проекта", 220),
                    ("CategoryName", "Категория", 150),
                    ("TargetAmount", "Цель (руб.)", 110),
                    ("CurrentAmount", "Собрано (руб.)", 110),
                    ("CompletionPercentage", "% сбора", 80),
                    ("RemainingAmount", "Остаток (руб.)", 110),
                    ("StartDate", "Дата начала", 100),
                    ("EndDate", "Дата оконч.", 100),
                    ("Status", "Статус", 90),
                    ("Description", "Описание", 200)
                });
                _lblInfo.Text = $"Проектов в базе: {list.Count}";
                if (selectId.HasValue) SelectRowById(_dgvProjects, selectId.Value);
            }
            catch (Exception ex) { ShowError(ex, "Загрузка списка проектов"); }
        }
        #endregion

        #region Вкладка 3: Категории
        private TabPage CreateCategoriesTab()
        {
            var tab = new TabPage("🏷 Категории");
            var pnlTool = new Panel { Dock = DockStyle.Top, Height = 48, BackColor = Color.FromArgb(248, 250, 252) };

            var btnAdd = CreateToolButton("➕ Добавить", Color.FromArgb(22, 163, 74), 10, 8);
            btnAdd.Click += (s, e) =>
            {
                using var dlg = new CategoryEditForm();
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        int newId = _categoryRepo.Insert(dlg.Category);
                        LoadCategories(newId);
                        _lblInfo.Text = $"✔ Категория «{dlg.Category.Name}» успешно добавлена!";
                    }
                    catch (Exception ex) { ShowError(ex, "Создание категории"); }
                }
            };

            var btnEdit = CreateToolButton("✏ Редактировать", Color.FromArgb(37, 99, 235), 120, 8);
            btnEdit.Click += (s, e) =>
            {
                if (_dgvCategories.SelectedRows.Count == 0) return;
                var cat = _dgvCategories.SelectedRows[0].DataBoundItem as Category;
                if (cat == null) return;

                using var dlg = new CategoryEditForm(cat);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _categoryRepo.Update(dlg.Category);
                        LoadCategories(dlg.Category.Id);
                        _lblInfo.Text = $"✔ Категория «{dlg.Category.Name}» обновлена";
                    }
                    catch (Exception ex) { ShowError(ex, "Редактирование категории"); }
                }
            };

            var btnDel = CreateToolButton("🗑 Удалить", Color.FromArgb(220, 38, 38), 255, 8);
            btnDel.Click += (s, e) =>
            {
                if (_dgvCategories.SelectedRows.Count == 0) return;
                var cat = _dgvCategories.SelectedRows[0].DataBoundItem as Category;
                if (cat == null) return;

                if (_categoryRepo.IsUsed(cat.Id))
                {
                    MessageBox.Show("Нельзя удалить категорию, так как с ней связаны проекты или получатели помощи!\nСначала удалите или перенесите связанные записи.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (MessageBox.Show($"Удалить категорию «{cat.Name}»?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        _categoryRepo.Delete(cat.Id);
                        LoadCategories();
                        _lblInfo.Text = "✔ Категория успешно удалена";
                    }
                    catch (Exception ex) { ShowError(ex, "Удаление категории"); }
                }
            };

            var btnCsv = CreateToolButton("📥 CSV", Color.FromArgb(71, 85, 105), 380, 8);
            btnCsv.Click += (s, e) => ExportService.ExportToCsv(_dgvCategories, "Categories");

            pnlTool.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDel, btnCsv });

            _dgvCategories = CreateStyledGrid();
            tab.Controls.Add(_dgvCategories);
            tab.Controls.Add(pnlTool);

            return tab;
        }

        private void LoadCategories(int? selectId = null)
        {
            try
            {
                var list = _categoryRepo.GetAll();
                _dgvCategories.DataSource = null;
                _dgvCategories.DataSource = list;
                ConfigureGridHeaders(_dgvCategories, new[] {
                    ("Id", "№", 50),
                    ("Name", "Наименование категории", 250),
                    ("Description", "Описание направления помощи", 450)
                });
                _lblInfo.Text = $"Категорий в базе: {list.Count}";
                if (selectId.HasValue) SelectRowById(_dgvCategories, selectId.Value);
            }
            catch (Exception ex) { ShowError(ex, "Загрузка списка категорий"); }
        }
        #endregion

        #region Вкладка 4: Пожертвования
        private TabPage CreateDonationsTab()
        {
            var tab = new TabPage("💰 Пожертвования");
            var pnlTool = new Panel { Dock = DockStyle.Top, Height = 80, BackColor = Color.FromArgb(248, 250, 252) };

            var btnAdd = CreateToolButton("➕ Внести взнос", Color.FromArgb(22, 163, 74), 10, 8);
            btnAdd.Click += (s, e) =>
            {
                var donors = _donorRepo.GetAll();
                var projects = _projectRepo.GetAll("Активен");
                if (donors.Count == 0 || projects.Count == 0)
                {
                    MessageBox.Show("Для оформления пожертвования требуется хотя бы один донор и один активный проект!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var dlg = new DonationEditForm(null, donors, projects);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        int newId = _donationRepo.InsertWithTransaction(dlg.Donation);
                        if (_cmbDonationProjectFilter.SelectedIndex != 0) _cmbDonationProjectFilter.SelectedIndex = 0;
                        if (_chkDonationDateFilter.Checked) _chkDonationDateFilter.Checked = false;
                        LoadDonations(newId);
                        LoadProjects(); // Баланс обновился в транзакции
                        _lblInfo.Text = $"✔ Пожертвование на сумму {dlg.Donation.Amount:N2} руб. успешно зарегистрировано!";
                    }
                    catch (Exception ex) { ShowError(ex, "Внесение пожертвования"); }
                }
            };

            var btnDel = CreateToolButton("🗑 Удалить (с откатом)", Color.FromArgb(220, 38, 38), 140, 8);
            btnDel.Click += (s, e) =>
            {
                if (_dgvDonations.SelectedRows.Count == 0) return;
                var don = _dgvDonations.SelectedRows[0].DataBoundItem as Donation;
                if (don == null) return;

                if (MessageBox.Show($"Удалить пожертвование №{don.Id} на сумму {don.Amount:N2} руб.?\nСумма сбора проекта будет автоматически уменьшена в транзакции!", "Откат пожертвования", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        _donationRepo.DeleteWithTransaction(don.Id);
                        LoadDonations();
                        LoadProjects();
                        _lblInfo.Text = "✔ Пожертвование удалено, баланс проекта скорректирован";
                    }
                    catch (Exception ex) { ShowError(ex, "Удаление пожертвования"); }
                }
            };

            var btnCsv = CreateToolButton("📥 CSV", Color.FromArgb(71, 85, 105), 310, 8);
            btnCsv.Click += (s, e) => ExportService.ExportToCsv(_dgvDonations, "Donations");

            var btnHtml = CreateToolButton("🌐 Ведомость (HTML)", Color.FromArgb(71, 85, 105), 400, 8);
            btnHtml.Click += (s, e) => ExportService.ExportToHtml(_dgvDonations, "Реестр пожертвований благотворительного фонда", "Donations_Registry");

            // Фильтры
            var lblProj = new Label { Text = "Проект:", Location = new Point(10, 48), AutoSize = true };
            _cmbDonationProjectFilter = new ComboBox { Location = new Point(65, 45), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbDonationProjectFilter.SelectedIndexChanged += (s, e) => LoadDonations();

            _chkDonationDateFilter = new CheckBox { Text = "Период с:", Location = new Point(300, 48), AutoSize = true };
            _chkDonationDateFilter.CheckedChanged += (s, e) =>
            {
                _dtpDonationFrom.Enabled = _chkDonationDateFilter.Checked;
                _dtpDonationTo.Enabled = _chkDonationDateFilter.Checked;
                LoadDonations();
            };

            _dtpDonationFrom = new DateTimePicker { Location = new Point(390, 45), Width = 110, Format = DateTimePickerFormat.Short, Enabled = false, Value = DateTime.Today.AddMonths(-1) };
            _dtpDonationFrom.ValueChanged += (s, e) => { if (_chkDonationDateFilter.Checked) LoadDonations(); };

            var lblTo = new Label { Text = "по:", Location = new Point(505, 48), AutoSize = true };
            _dtpDonationTo = new DateTimePicker { Location = new Point(530, 45), Width = 110, Format = DateTimePickerFormat.Short, Enabled = false, Value = DateTime.Today };
            _dtpDonationTo.ValueChanged += (s, e) => { if (_chkDonationDateFilter.Checked) LoadDonations(); };

            _lblDonationsTotal = new Label
            {
                Location = new Point(660, 48),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(22, 101, 52)
            };

            pnlTool.Controls.AddRange(new Control[] {
                btnAdd, btnDel, btnCsv, btnHtml,
                lblProj, _cmbDonationProjectFilter,
                _chkDonationDateFilter, _dtpDonationFrom, lblTo, _dtpDonationTo,
                _lblDonationsTotal
            });

            _dgvDonations = CreateStyledGrid();
            tab.Controls.Add(_dgvDonations);
            tab.Controls.Add(pnlTool);

            return tab;
        }

        private void LoadDonations(int? selectId = null)
        {
            try
            {
                // Заполнение выпадающего списка проектов
                if (_cmbDonationProjectFilter.Items.Count == 0)
                {
                    _cmbDonationProjectFilter.Items.Add(new Project { Id = 0, Name = "— Все проекты —" });
                    foreach (var p in _projectRepo.GetAll()) _cmbDonationProjectFilter.Items.Add(p);
                    _cmbDonationProjectFilter.SelectedIndex = 0;
                }

                int? projId = (_cmbDonationProjectFilter.SelectedItem as Project)?.Id;
                DateTime? dateFrom = _chkDonationDateFilter.Checked ? _dtpDonationFrom.Value : null;
                DateTime? dateTo = _chkDonationDateFilter.Checked ? _dtpDonationTo.Value : null;

                var list = _donationRepo.GetAll(projId, null, dateFrom, dateTo);
                _dgvDonations.DataSource = null;
                _dgvDonations.DataSource = list;

                ConfigureGridHeaders(_dgvDonations, new[] {
                    ("Id", "№", 50),
                    ("DonorName", "Благотворитель (донор)", 210),
                    ("ProjectName", "Проект", 210),
                    ("Amount", "Сумма (руб.)", 120),
                    ("DonationDate", "Дата и время", 130),
                    ("PaymentMethod", "Способ оплаты", 130),
                    ("Notes", "Примечание", 230)
                });

                decimal total = list.Sum(d => d.Amount);
                _lblDonationsTotal.Text = $"Итого: {total:N2} руб. ({list.Count} взносов)";
                _lblInfo.Text = $"Пожертвований: {list.Count} на сумму {total:N2} руб.";
                if (selectId.HasValue) SelectRowById(_dgvDonations, selectId.Value);
            }
            catch (Exception ex) { ShowError(ex, "Загрузка списка пожертвований"); }
        }
        #endregion

        #region Вкладка 5: Получатели помощи
        private TabPage CreateRecipientsTab()
        {
            var tab = new TabPage("🤝 Получатели помощи");
            var pnlTool = new Panel { Dock = DockStyle.Top, Height = 48, BackColor = Color.FromArgb(248, 250, 252) };

            var btnAdd = CreateToolButton("➕ Добавить", Color.FromArgb(22, 163, 74), 10, 8);
            btnAdd.Click += (s, e) =>
            {
                var cats = _categoryRepo.GetAll();
                if (cats.Count == 0)
                {
                    MessageBox.Show("Сначала создайте категории!", "Предупреждение", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                using var dlg = new RecipientEditForm(null, cats);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        int newId = _recipientRepo.Insert(dlg.Recipient);
                        if (_cmbRecipientStatus.SelectedIndex != 0) _cmbRecipientStatus.SelectedIndex = 0;
                        if (!string.IsNullOrEmpty(_txtRecipientSearch.Text)) _txtRecipientSearch.Text = string.Empty;
                        LoadRecipients(newId);
                        _lblInfo.Text = $"✔ Получатель «{dlg.Recipient.FullName}» успешно добавлен!";
                    }
                    catch (Exception ex) { ShowError(ex, "Добавление получателя помощи"); }
                }
            };

            var btnEdit = CreateToolButton("✏ Редактировать", Color.FromArgb(37, 99, 235), 120, 8);
            btnEdit.Click += (s, e) =>
            {
                if (_dgvRecipients.SelectedRows.Count == 0) return;
                var r = _dgvRecipients.SelectedRows[0].DataBoundItem as Recipient;
                if (r == null) return;

                var cats = _categoryRepo.GetAll();
                using var dlg = new RecipientEditForm(r, cats);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _recipientRepo.Update(dlg.Recipient);
                        LoadRecipients(dlg.Recipient.Id);
                        _lblInfo.Text = $"✔ Данные получателя «{dlg.Recipient.FullName}» обновлены";
                    }
                    catch (Exception ex) { ShowError(ex, "Редактирование получателя помощи"); }
                }
            };

            var btnDel = CreateToolButton("🗑 Удалить", Color.FromArgb(220, 38, 38), 255, 8);
            btnDel.Click += (s, e) =>
            {
                if (_dgvRecipients.SelectedRows.Count == 0) return;
                var r = _dgvRecipients.SelectedRows[0].DataBoundItem as Recipient;
                if (r == null) return;

                if (MessageBox.Show($"Удалить получателя «{r.FullName}»?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        _recipientRepo.Delete(r.Id);
                        LoadRecipients();
                        _lblInfo.Text = "✔ Получатель успешно удален";
                    }
                    catch (Exception ex) { ShowError(ex, "Удаление получателя помощи"); }
                }
            };

            var lblStatus = new Label { Text = "Статус:", Location = new Point(370, 14), AutoSize = true };
            _cmbRecipientStatus = new ComboBox { Location = new Point(425, 11), Width = 140, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbRecipientStatus.Items.AddRange(new object[] { "Все", "На рассмотрении", "Одобрен", "Помощь оказана", "Отклонен" });
            _cmbRecipientStatus.SelectedIndex = 0;
            _cmbRecipientStatus.SelectedIndexChanged += (s, e) => LoadRecipients();

            var lblSearch = new Label { Text = "Поиск:", Location = new Point(580, 14), AutoSize = true };
            _txtRecipientSearch = new TextBox { Location = new Point(630, 11), Width = 170 };
            _txtRecipientSearch.TextChanged += (s, e) => LoadRecipients();

            var btnCsv = CreateToolButton("📥 CSV", Color.FromArgb(71, 85, 105), 815, 8);
            btnCsv.Click += (s, e) => ExportService.ExportToCsv(_dgvRecipients, "Recipients");

            pnlTool.Controls.AddRange(new Control[] { btnAdd, btnEdit, btnDel, lblStatus, _cmbRecipientStatus, lblSearch, _txtRecipientSearch, btnCsv });

            _dgvRecipients = CreateStyledGrid();
            tab.Controls.Add(_dgvRecipients);
            tab.Controls.Add(pnlTool);

            return tab;
        }

        private void LoadRecipients(int? selectId = null)
        {
            try
            {
                var list = _recipientRepo.GetAll(_cmbRecipientStatus.SelectedItem?.ToString(), _txtRecipientSearch.Text);
                _dgvRecipients.DataSource = null;
                _dgvRecipients.DataSource = list;
                ConfigureGridHeaders(_dgvRecipients, new[] {
                    ("Id", "№", 45),
                    ("FullName", "ФИО / Наименование получателя", 210),
                    ("CategoryName", "Категория помощи", 140),
                    ("Status", "Статус", 120),
                    ("NeedDescription", "Потребность в помощи", 260),
                    ("Phone", "Телефон", 120),
                    ("Address", "Адрес", 170),
                    ("RegistrationDate", "Дата заявления", 100)
                });
                _lblInfo.Text = $"Получателей помощи в базе: {list.Count}";
                if (selectId.HasValue) SelectRowById(_dgvRecipients, selectId.Value);
            }
            catch (Exception ex) { ShowError(ex, "Загрузка списка получателей помощи"); }
        }
        #endregion

        #region Вкладка 6: Расходы фонда
        private TabPage CreateExpensesTab()
        {
            var tab = new TabPage("📉 Расходы фонда");
            var pnlTool = new Panel { Dock = DockStyle.Top, Height = 80, BackColor = Color.FromArgb(248, 250, 252) };

            var btnAdd = CreateToolButton("➕ Зарегистрировать расход", Color.FromArgb(22, 163, 74), 10, 8);
            btnAdd.Click += (s, e) =>
            {
                var projs = _projectRepo.GetAll();
                var recs = _recipientRepo.GetAll();
                if (projs.Count == 0)
                {
                    MessageBox.Show("Требуется хотя бы один проект для списания средств!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                using var dlg = new ExpenseEditForm(null, projs, recs);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        int newId = _expenseRepo.Insert(dlg.Expense);
                        if (_cmbExpenseProjectFilter.SelectedIndex != 0) _cmbExpenseProjectFilter.SelectedIndex = 0;
                        if (_cmbExpenseCategoryFilter.SelectedIndex != 0) _cmbExpenseCategoryFilter.SelectedIndex = 0;
                        LoadExpenses(newId);
                        LoadProjects();
                        _lblInfo.Text = $"✔ Расход на сумму {dlg.Expense.Amount:N2} руб. успешно зафиксирован!";
                    }
                    catch (Exception ex) { ShowError(ex, "Регистрация расхода фонда"); }
                }
            };

            var btnEdit = CreateToolButton("✏ Редактировать", Color.FromArgb(37, 99, 235), 215, 8);
            btnEdit.Click += (s, e) =>
            {
                if (_dgvExpenses.SelectedRows.Count == 0) return;
                var exp = _dgvExpenses.SelectedRows[0].DataBoundItem as Expense;
                if (exp == null) return;

                var projs = _projectRepo.GetAll();
                var recs = _recipientRepo.GetAll();
                using var dlg = new ExpenseEditForm(exp, projs, recs);
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _expenseRepo.Update(dlg.Expense);
                        LoadExpenses(dlg.Expense.Id);
                        LoadProjects();
                        _lblInfo.Text = $"✔ Данные расхода №{dlg.Expense.Id} обновлены";
                    }
                    catch (Exception ex) { ShowError(ex, "Редактирование расхода"); }
                }
            };

            var btnDel = CreateToolButton("🗑 Удалить", Color.FromArgb(220, 38, 38), 350, 8);
            btnDel.Click += (s, e) =>
            {
                if (_dgvExpenses.SelectedRows.Count == 0) return;
                var exp = _dgvExpenses.SelectedRows[0].DataBoundItem as Expense;
                if (exp == null) return;

                if (MessageBox.Show($"Удалить запись о расходе №{exp.Id} на сумму {exp.Amount:N2} руб.?", "Подтверждение удаления", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    try
                    {
                        _expenseRepo.Delete(exp.Id);
                        LoadExpenses();
                        LoadProjects();
                        _lblInfo.Text = "✔ Запись о расходе успешно удалена";
                    }
                    catch (Exception ex) { ShowError(ex, "Удаление расхода"); }
                }
            };

            var btnCsv = CreateToolButton("📥 CSV", Color.FromArgb(71, 85, 105), 455, 8);
            btnCsv.Click += (s, e) => ExportService.ExportToCsv(_dgvExpenses, "Expenses");

            // Фильтры
            var lblProj = new Label { Text = "Проект:", Location = new Point(10, 48), AutoSize = true };
            _cmbExpenseProjectFilter = new ComboBox { Location = new Point(65, 45), Width = 220, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbExpenseProjectFilter.SelectedIndexChanged += (s, e) => LoadExpenses();

            var lblCat = new Label { Text = "Статья:", Location = new Point(300, 48), AutoSize = true };
            _cmbExpenseCategoryFilter = new ComboBox { Location = new Point(355, 45), Width = 180, DropDownStyle = ComboBoxStyle.DropDownList };
            _cmbExpenseCategoryFilter.Items.AddRange(new object[] { "Все", "Адресная помощь", "Закупка медикаментов", "Оборудование", "Транспорт", "Административные расходы" });
            _cmbExpenseCategoryFilter.SelectedIndex = 0;
            _cmbExpenseCategoryFilter.SelectedIndexChanged += (s, e) => LoadExpenses();

            _lblExpensesTotal = new Label
            {
                Location = new Point(555, 48),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(185, 28, 28)
            };

            pnlTool.Controls.AddRange(new Control[] {
                btnAdd, btnEdit, btnDel, btnCsv,
                lblProj, _cmbExpenseProjectFilter,
                lblCat, _cmbExpenseCategoryFilter,
                _lblExpensesTotal
            });

            _dgvExpenses = CreateStyledGrid();
            tab.Controls.Add(_dgvExpenses);
            tab.Controls.Add(pnlTool);

            return tab;
        }

        private void LoadExpenses(int? selectId = null)
        {
            try
            {
                if (_cmbExpenseProjectFilter.Items.Count == 0)
                {
                    _cmbExpenseProjectFilter.Items.Add(new Project { Id = 0, Name = "— Все проекты —" });
                    foreach (var p in _projectRepo.GetAll()) _cmbExpenseProjectFilter.Items.Add(p);
                    _cmbExpenseProjectFilter.SelectedIndex = 0;
                }

                int? projId = (_cmbExpenseProjectFilter.SelectedItem as Project)?.Id;
                string? cat = _cmbExpenseCategoryFilter.SelectedItem?.ToString();

                var list = _expenseRepo.GetAll(projId, cat);
                _dgvExpenses.DataSource = null;
                _dgvExpenses.DataSource = list;

                ConfigureGridHeaders(_dgvExpenses, new[] {
                    ("Id", "№", 50),
                    ("ProjectName", "Проект списания", 200),
                    ("ExpenseCategory", "Статья расхода", 150),
                    ("Amount", "Сумма (руб.)", 110),
                    ("ExpenseDate", "Дата расхода", 100),
                    ("DocumentNumber", "Первичный документ", 140),
                    ("RecipientName", "Адресный получатель", 190),
                    ("Description", "Назначение / Детализация", 200)
                });

                decimal total = list.Sum(e => e.Amount);
                _lblExpensesTotal.Text = $"Всего израсходовано: {total:N2} руб.";
                _lblInfo.Text = $"Расходов: {list.Count} на сумму {total:N2} руб.";
                if (selectId.HasValue) SelectRowById(_dgvExpenses, selectId.Value);
            }
            catch (Exception ex) { ShowError(ex, "Загрузка списка расходов"); }
        }

        private void RefreshProjectDropdowns()
        {
            try
            {
                if (_cmbDonationProjectFilter != null)
                {
                    int cur = (_cmbDonationProjectFilter.SelectedItem as Project)?.Id ?? 0;
                    _cmbDonationProjectFilter.Items.Clear();
                    _cmbDonationProjectFilter.Items.Add(new Project { Id = 0, Name = "— Все проекты —" });
                    var projs = _projectRepo.GetAll();
                    int sel = 0;
                    for (int i = 0; i < projs.Count; i++)
                    {
                        _cmbDonationProjectFilter.Items.Add(projs[i]);
                        if (projs[i].Id == cur) sel = i + 1;
                    }
                    _cmbDonationProjectFilter.SelectedIndex = sel;
                }

                if (_cmbExpenseProjectFilter != null)
                {
                    int cur = (_cmbExpenseProjectFilter.SelectedItem as Project)?.Id ?? 0;
                    _cmbExpenseProjectFilter.Items.Clear();
                    _cmbExpenseProjectFilter.Items.Add(new Project { Id = 0, Name = "— Все проекты —" });
                    var projs = _projectRepo.GetAll();
                    int sel = 0;
                    for (int i = 0; i < projs.Count; i++)
                    {
                        _cmbExpenseProjectFilter.Items.Add(projs[i]);
                        if (projs[i].Id == cur) sel = i + 1;
                    }
                    _cmbExpenseProjectFilter.SelectedIndex = sel;
                }
            }
            catch
            {
                // Игнорируем ошибки фонового обновления комбобоксов
            }
        }
        #endregion

        #region Вкладка 7: Аналитические отчеты (5 обязательных запросов)
        private TabPage CreateReportsTab()
        {
            var tab = new TabPage("📊 Аналитические запросы и отчеты");
            var pnlTool = new Panel { Dock = DockStyle.Top, Height = 95, BackColor = Color.FromArgb(241, 245, 249) };

            var lblReport = new Label { Text = "Выберите аналитический отчет (согласно заданию практики):", Location = new Point(12, 10), AutoSize = true, Font = new Font("Segoe UI", 9.5f, FontStyle.Bold) };

            _cmbReportType = new ComboBox
            {
                Location = new Point(12, 33),
                Width = 530,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular)
            };
            _cmbReportType.Items.AddRange(new object[] {
                "1. Сумма пожертвований (сводка по проектам и фильтрация по периоду)",
                "2. Активные проекты (цель, собрано, процент выполнения, остаток к сбору)",
                "3. Крупнейшие доноры (рейтинг благотворителей по общей сумме взносов)",
                "4. Расходы фонда (сравнение собранных и израсходованных средств, баланс)",
                "5. Помощь по категориям (дети, медицина, малоимущие, ветераны, животные)"
            });
            _cmbReportType.SelectedIndex = 0;
            _cmbReportType.SelectedIndexChanged += (s, e) => GenerateSelectedReport();

            _chkReportDateFilter = new CheckBox { Text = "Фильтр по дате с:", Location = new Point(12, 67), AutoSize = true };
            _chkReportDateFilter.CheckedChanged += (s, e) =>
            {
                _dtpReportFrom.Enabled = _chkReportDateFilter.Checked;
                _dtpReportTo.Enabled = _chkReportDateFilter.Checked;
                GenerateSelectedReport();
            };

            _dtpReportFrom = new DateTimePicker { Location = new Point(140, 64), Width = 110, Format = DateTimePickerFormat.Short, Enabled = false, Value = DateTime.Today.AddMonths(-3) };
            _dtpReportFrom.ValueChanged += (s, e) => { if (_chkReportDateFilter.Checked) GenerateSelectedReport(); };

            var lblTo = new Label { Text = "по:", Location = new Point(255, 67), AutoSize = true };
            _dtpReportTo = new DateTimePicker { Location = new Point(280, 64), Width = 110, Format = DateTimePickerFormat.Short, Enabled = false, Value = DateTime.Today };
            _dtpReportTo.ValueChanged += (s, e) => { if (_chkReportDateFilter.Checked) GenerateSelectedReport(); };

            var btnRun = CreateToolButton("⚡ Сформировать", Color.FromArgb(37, 99, 235), 555, 30);
            btnRun.Height = 32;
            btnRun.Click += (s, e) => GenerateSelectedReport();

            var btnCsv = CreateToolButton("📥 В Excel / CSV", Color.FromArgb(22, 163, 74), 690, 30);
            btnCsv.Height = 32;
            btnCsv.Click += (s, e) => ExportService.ExportToCsv(_dgvReports, $"Report_{_cmbReportType.SelectedIndex + 1}");

            var btnHtml = CreateToolButton("🌐 Печать / HTML", Color.FromArgb(71, 85, 105), 825, 30);
            btnHtml.Height = 32;
            btnHtml.Click += (s, e) => ExportService.ExportToHtml(_dgvReports, _cmbReportType.SelectedItem?.ToString() ?? "Отчет", $"Report_{_cmbReportType.SelectedIndex + 1}");

            _lblReportSummary = new Label
            {
                Location = new Point(410, 67),
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(30, 41, 59)
            };

            pnlTool.Controls.AddRange(new Control[] {
                lblReport, _cmbReportType, _chkReportDateFilter,
                _dtpReportFrom, lblTo, _dtpReportTo,
                btnRun, btnCsv, btnHtml, _lblReportSummary
            });

            _dgvReports = CreateStyledGrid();
            tab.Controls.Add(_dgvReports);
            tab.Controls.Add(pnlTool);

            return tab;
        }

        private void GenerateSelectedReport()
        {
            try
            {
                DateTime? dateFrom = _chkReportDateFilter.Checked ? _dtpReportFrom.Value : null;
                DateTime? dateTo = _chkReportDateFilter.Checked ? _dtpReportTo.Value : null;

                _dgvReports.DataSource = null;

                switch (_cmbReportType.SelectedIndex)
                {
                    case 0: // 1. Сумма пожертвований
                        {
                            var list = _reportRepo.GetDonationsSummary(dateFrom, dateTo);
                            _dgvReports.DataSource = list;
                            ConfigureGridHeaders(_dgvReports, new[] {
                                ("ProjectId", "ID", 45),
                                ("ProjectName", "Проект", 230),
                                ("CategoryName", "Категория", 150),
                                ("DonationsCount", "Кол-во взносов", 120),
                                ("TotalDonated", "Сумма пожертвований (руб.)", 160),
                                ("TargetAmount", "Цель сбора (руб.)", 130),
                                ("FirstDonationDate", "Первый взнос", 110),
                                ("LastDonationDate", "Последний взнос", 110)
                            });
                            decimal sum = list.Sum(x => x.TotalDonated);
                            int totalCount = list.Sum(x => x.DonationsCount);
                            _lblReportSummary.Text = $"Всего собрано: {sum:N2} руб. | Взносов: {totalCount}";
                        }
                        break;

                    case 1: // 2. Активные проекты
                        {
                            var list = _reportRepo.GetActiveProjects();
                            _dgvReports.DataSource = list;
                            ConfigureGridHeaders(_dgvReports, new[] {
                                ("ProjectId", "ID", 45),
                                ("ProjectName", "Название активного проекта", 240),
                                ("CategoryName", "Категория", 150),
                                ("TargetAmount", "Цель (руб.)", 120),
                                ("CurrentAmount", "Собрано (руб.)", 120),
                                ("CompletionPercentage", "Выполнение (%)", 110),
                                ("RemainingAmount", "Осталось собрать (руб.)", 140),
                                ("StartDate", "Дата начала", 100),
                                ("EndDate", "Плановое окончание", 110),
                                ("Status", "Статус", 80)
                            });
                            decimal totalTarget = list.Sum(x => x.TargetAmount);
                            decimal totalCollected = list.Sum(x => x.CurrentAmount);
                            _lblReportSummary.Text = $"Активных проектов: {list.Count} | Собрано: {totalCollected:N2} из {totalTarget:N2} руб.";
                        }
                        break;

                    case 2: // 3. Крупнейшие доноры
                        {
                            var list = _reportRepo.GetTopDonors(15);
                            _dgvReports.DataSource = list;
                            ConfigureGridHeaders(_dgvReports, new[] {
                                ("DonorId", "ID", 45),
                                ("DonorName", "ФИО / Наименование благотворителя", 250),
                                ("DonorType", "Тип", 130),
                                ("DonationsCount", "Кол-во взносов", 110),
                                ("TotalDonated", "Общая сумма взносов (руб.)", 160),
                                ("MaxSingleDonation", "Макс. разовый взнос (руб.)", 150),
                                ("Phone", "Телефон", 130),
                                ("Email", "Email", 160),
                                ("LastDonationDate", "Дата посл. взноса", 120)
                            });
                            decimal topSum = list.Sum(x => x.TotalDonated);
                            _lblReportSummary.Text = $"Топ-{list.Count} доноров внесли суммарно: {topSum:N2} руб.";
                        }
                        break;

                    case 3: // 4. Расходы фонда
                        {
                            var list = _reportRepo.GetFundExpensesBalance();
                            _dgvReports.DataSource = list;
                            ConfigureGridHeaders(_dgvReports, new[] {
                                ("ProjectId", "ID", 45),
                                ("ProjectName", "Проект", 260),
                                ("TotalCollected", "Собрано по проекту (руб.)", 160),
                                ("TotalSpent", "Израсходовано (руб.)", 150),
                                ("Balance", "Остаток баланса (руб.)", 150),
                                ("ExpensesCount", "Кол-во актов расхода", 130)
                            });
                            decimal totalColl = list.Sum(x => x.TotalCollected);
                            decimal totalSpent = list.Sum(x => x.TotalSpent);
                            decimal bal = totalColl - totalSpent;
                            _lblReportSummary.Text = $"Собрано: {totalColl:N2} | Израсходовано: {totalSpent:N2} | Текущий баланс: {bal:N2} руб.";
                        }
                        break;

                    case 4: // 5. Помощь по категориям
                        {
                            var list = _reportRepo.GetCategoryAssistance();
                            _dgvReports.DataSource = list;
                            ConfigureGridHeaders(_dgvReports, new[] {
                                ("CategoryId", "ID", 45),
                                ("CategoryName", "Категория благотворительности", 240),
                                ("ProjectsCount", "Кол-во проектов", 120),
                                ("RecipientsCount", "Получателей помощи", 130),
                                ("TotalDonated", "Собрано пожертвований (руб.)", 170),
                                ("TotalExpenses", "Оказано помощи (руб.)", 160)
                            });
                            decimal donSum = list.Sum(x => x.TotalDonated);
                            decimal expSum = list.Sum(x => x.TotalExpenses);
                            _lblReportSummary.Text = $"Всего направлений: {list.Count} | Сборы: {donSum:N2} руб. | Выплаты: {expSum:N2} руб.";
                        }
                        break;
                }
            }
            catch (Exception ex) { ShowError(ex, "Формирование аналитического отчета"); }
        }
        #endregion

        #region Вспомогательные методы UI
        private void RefreshCurrentTab()
        {
            switch (_mainTabControl.SelectedIndex)
            {
                case 0: LoadDonors(); break;
                case 1: LoadProjects(); break;
                case 2: LoadCategories(); break;
                case 3: LoadDonations(); break;
                case 4: LoadRecipients(); break;
                case 5: LoadExpenses(); break;
                case 6: GenerateSelectedReport(); break;
            }
        }

        private void RefreshAllTabs()
        {
            RefreshProjectDropdowns();
            LoadDonors();
            LoadProjects();
            LoadCategories();
            LoadDonations();
            LoadRecipients();
            LoadExpenses();
            GenerateSelectedReport();
            _lblInfo.Text = "Все данные успешно обновлены";
        }

        private static Button CreateToolButton(string text, Color backColor, int x, int y)
        {
            return new Button
            {
                Text = text,
                Location = new Point(x, y),
                Size = new Size(105, 32),
                BackColor = backColor,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                Cursor = Cursors.Hand
            };
        }

        private static DataGridView CreateStyledGrid()
        {
            var dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                ReadOnly = true,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false,
                AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None,
                EnableHeadersVisualStyles = false,
                RowTemplate = { Height = 28 }
            };

            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(241, 245, 249);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(15, 23, 42);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
            dgv.ColumnHeadersHeight = 34;

            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            dgv.DefaultCellStyle.SelectionForeColor = Color.FromArgb(15, 23, 42);
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);

            return dgv;
        }

        private static void ConfigureGridHeaders(DataGridView dgv, (string Prop, string Header, int Width)[] columns)
        {
            foreach (DataGridViewColumn col in dgv.Columns)
            {
                col.Visible = false;
            }

            for (int i = 0; i < columns.Length; i++)
            {
                var (prop, header, width) = columns[i];
                if (dgv.Columns.Contains(prop))
                {
                    var col = dgv.Columns[prop];
                    col.Visible = true;
                    col.HeaderText = header;
                    col.Width = width;
                    try
                    {
                        if (i < dgv.Columns.Count) col.DisplayIndex = i;
                    }
                    catch
                    {
                        // Игнорируем ошибки DisplayIndex при динамической смене источников
                    }

                    if (prop.Contains("Amount") || prop.Contains("Donated") || prop.Contains("Spent") || prop.Contains("Balance"))
                    {
                        col.DefaultCellStyle.Format = "N2";
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                    }
                    else if (prop.Contains("Percentage"))
                    {
                        col.DefaultCellStyle.Format = "N2";
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                    else if (prop.Contains("Date"))
                    {
                        col.DefaultCellStyle.Format = "dd.MM.yyyy";
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                    else if (prop == "Id" || prop.EndsWith("Id") || prop.Contains("Count"))
                    {
                        col.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                }
            }
        }

        private static void SelectRowById(DataGridView dgv, int id)
        {
            if (dgv.Rows.Count == 0) return;
            try
            {
                dgv.ClearSelection();
                foreach (DataGridViewRow row in dgv.Rows)
                {
                    if (row.DataBoundItem != null)
                    {
                        var prop = row.DataBoundItem.GetType().GetProperty("Id");
                        if (prop != null && prop.GetValue(row.DataBoundItem) is int val && val == id)
                        {
                            row.Selected = true;
                            if (row.Index >= 0 && row.Index < dgv.Rows.Count)
                            {
                                dgv.FirstDisplayedScrollingRowIndex = row.Index;
                            }
                            break;
                        }
                    }
                }
            }
            catch
            {
                // Игнорируем исключения прокрутки скрытых элементов
            }
        }

        private static void ShowError(Exception ex, string? actionDescription = null)
        {
            string message = ex.Message ?? string.Empty;
            string title = "Внимание";
            MessageBoxIcon icon = MessageBoxIcon.Warning;
            string friendlyText;

            if (message.Contains("FOREIGN KEY", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("REFERENCE constraint", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("FK_", StringComparison.OrdinalIgnoreCase))
            {
                friendlyText = "Невозможно удалить или изменить выбранную запись, так как с ней связаны другие данные фонда.\n\n" +
                               "Например, по данному элементу уже зафиксированы пожертвования, проекты или выплаты.\n" +
                               "Для выполнения операции сначала удалите связанные записи.";
            }
            else if (message.Contains("UNIQUE constraint", StringComparison.OrdinalIgnoreCase) ||
                     message.Contains("duplicate key", StringComparison.OrdinalIgnoreCase) ||
                     message.Contains("PRIMARY KEY", StringComparison.OrdinalIgnoreCase))
            {
                friendlyText = "Запись с такими данными уже существует в базе данных.\n\n" +
                               "Пожалуйста, укажите уникальное значение (например, другое наименование, телефон или номер документа).";
            }
            else if (message.Contains("NOT NULL constraint", StringComparison.OrdinalIgnoreCase) ||
                     message.Contains("Cannot insert the value NULL", StringComparison.OrdinalIgnoreCase))
            {
                friendlyText = "Заполнены не все обязательные поля формы.\n\n" +
                               "Пожалуйста, вернитесь к форме и проверьте правильность заполнения всех обязательных полей.";
            }
            else if (message.Contains("database is locked", StringComparison.OrdinalIgnoreCase) ||
                     message.Contains("busy", StringComparison.OrdinalIgnoreCase) ||
                     message.Contains("being used by another process", StringComparison.OrdinalIgnoreCase))
            {
                friendlyText = "Файл базы данных временно занят другой программой или процессом.\n\n" +
                               "Пожалуйста, закройте сторонние приложения (например, SQLite Studio или проводник) и повторите попытку.";
            }
            else if (ex is FormatException || ex is OverflowException || ex is InvalidCastException)
            {
                friendlyText = "Введены некорректные данные.\n\n" +
                               "Пожалуйста, проверьте правильность числовых значений, сумм и дат.";
            }
            else if (message.Contains("cannot open database", StringComparison.OrdinalIgnoreCase) ||
                     message.Contains("Could not find file", StringComparison.OrdinalIgnoreCase) ||
                     message.Contains("SqlException", StringComparison.OrdinalIgnoreCase) ||
                     ex is FileNotFoundException)
            {
                friendlyText = "Не удалось подключиться к базе данных.\n\n" +
                               "Убедитесь, что файл базы данных CharityFundDB.db существует и доступен для записи.";
                icon = MessageBoxIcon.Error;
                title = "Ошибка доступа к базе данных";
            }
            else
            {
                friendlyText = "Не удалось выполнить операцию.\n\n" +
                               "Пожалуйста, проверьте правильность введённых данных и повторите попытку.";
                icon = MessageBoxIcon.Error;
                title = "Ошибка";
            }

            if (!string.IsNullOrWhiteSpace(actionDescription))
            {
                friendlyText = $"Операция: {actionDescription}\n\n" + friendlyText;
            }

            MessageBox.Show(friendlyText, title, MessageBoxButtons.OK, icon);
        }
        #endregion
    }
}
