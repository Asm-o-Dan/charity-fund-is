using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Common;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Data.SqlClient;
using Microsoft.Data.Sqlite;

namespace CharityFundApp.DataAccess
{
    public enum DatabaseProvider
    {
        Sqlite,
        SqlServer
    }

    /// <summary>
    /// Вспомогательный класс доступа к данным через чистый ADO.NET.
    /// Основной режим по умолчанию — надежная локальная база данных SQLite в виде файла CharityFundDB.db,
    /// которую можно легко переносить на любой компьютер без установки сторонних серверов СУБД.
    /// Также сохранена поддержка корпоративного MS SQL Server.
    /// </summary>
    public static class DatabaseHelper
    {
        private static string? _customConnectionString;
        private static DatabaseProvider _currentProvider = DatabaseProvider.Sqlite;

        public static DatabaseProvider CurrentProvider
        {
            get => _currentProvider;
            set => _currentProvider = value;
        }

        public static string SqliteDbFileName => "CharityFundDB.db";
        public static string SqliteDbPath => Path.Combine(AppDomain.CurrentDomain.BaseDirectory, SqliteDbFileName);

        /// <summary>
        /// Текущая строка подключения к базе данных.
        /// По умолчанию используется локальный файл SQLite CharityFundDB.db.
        /// </summary>
        public static string ConnectionString
        {
            get
            {
                if (_currentProvider == DatabaseProvider.Sqlite)
                {
                    return $"Data Source={SqliteDbPath};";
                }

                if (!string.IsNullOrWhiteSpace(_customConnectionString))
                {
                    return _customConnectionString;
                }

                try
                {
                    var configConn = ConfigurationManager.ConnectionStrings["CharityFundDB"]?.ConnectionString;
                    if (!string.IsNullOrWhiteSpace(configConn))
                    {
                        return configConn;
                    }
                }
                catch
                {
                    // Игнорируем ошибку чтения конфигуратора
                }

                return "Server=localhost;Database=CharityFundDB;Trusted_Connection=True;TrustServerCertificate=True;";
            }
            set
            {
                _customConnectionString = value;
            }
        }

        /// <summary>
        /// Создает и возвращает новое соединение DbConnection
        /// </summary>
        public static DbConnection CreateConnection()
        {
            if (_currentProvider == DatabaseProvider.Sqlite)
            {
                EnsureSqliteDatabaseCreated();
                return new SqliteConnection(ConnectionString);
            }
            return new SqlConnection(ConnectionString);
        }

        /// <summary>
        /// Проверка доступности подключения к базе данных
        /// </summary>
        public static (bool Success, string Message) TestConnection()
        {
            try
            {
                if (_currentProvider == DatabaseProvider.Sqlite)
                {
                    EnsureSqliteDatabaseCreated();
                    using var conn = new SqliteConnection(ConnectionString);
                    conn.Open();
                    using var cmd = new SqliteCommand("SELECT COUNT(*) FROM Projects;", conn);
                    long count = Convert.ToInt64(cmd.ExecuteScalar() ?? 0);
                    return (true, $"База данных SQLite подключена ({SqliteDbFileName})! Проектов: {count}");
                }
                else
                {
                    using var conn = new SqlConnection(ConnectionString);
                    conn.Open();

                    using var cmd = new SqlCommand("SELECT DB_NAME() AS CurrentDb, @@VERSION AS SqlVersion;", conn);
                    using var reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        string dbName = reader["CurrentDb"]?.ToString() ?? "Unknown";
                        string version = reader["SqlVersion"]?.ToString() ?? "";
                        string shortVer = version.Split('\n')[0];
                        return (true, $"Подключение к MS SQL Server «{dbName}» ({shortVer})");
                    }

                    return (true, "Подключение к MS SQL Server успешно установлено!");
                }
            }
            catch (Exception ex)
            {
                return (false, $"Ошибка подключения: {ex.Message}");
            }
        }

        /// <summary>
        /// Адаптирует SQL запрос под активный провайдер (для 100% совместимости SQLite и SQL Server)
        /// </summary>
        private static string AdaptSql(string sql)
        {
            if (_currentProvider == DatabaseProvider.Sqlite)
            {
                string adapted = sql;
                adapted = Regex.Replace(adapted, @"\bISNULL\(", "COALESCE(", RegexOptions.IgnoreCase);
                adapted = Regex.Replace(adapted, @"\bSCOPE_IDENTITY\(\)", "last_insert_rowid()", RegexOptions.IgnoreCase);
                adapted = Regex.Replace(adapted, @"(?<!\w)N'([^']*)'", "'$1'");

                // Замена SELECT TOP (@Limit) ... на SELECT ... LIMIT @Limit
                var topMatch = Regex.Match(adapted, @"SELECT\s+TOP\s*\((@?\w+)\)\s*(.*?)(;?\s*$)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                if (topMatch.Success)
                {
                    string limitVar = topMatch.Groups[1].Value;
                    string rest = topMatch.Groups[2].Value.TrimEnd(';', ' ', '\r', '\n');
                    adapted = $"SELECT {rest} LIMIT {limitVar};";
                }

                return adapted;
            }
            return sql;
        }

        /// <summary>
        /// Выполняет параметризованный SELECT-запрос и возвращает результат в виде DataTable
        /// </summary>
        public static DataTable ExecuteQuery(string query, params DbParameter[] parameters)
        {
            query = AdaptSql(query);

            using var conn = CreateConnection();
            conn.Open();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = query;

            if (parameters != null && parameters.Length > 0)
            {
                foreach (var p in parameters)
                {
                    var param = cmd.CreateParameter();
                    param.ParameterName = p.ParameterName;
                    param.Value = p.Value ?? DBNull.Value;
                    cmd.Parameters.Add(param);
                }
            }

            var table = new DataTable();
            using var reader = cmd.ExecuteReader();
            table.Load(reader);
            return table;
        }

        /// <summary>
        /// Выполняет INSERT, UPDATE или DELETE запрос с параметрами
        /// </summary>
        public static int ExecuteNonQuery(string commandText, DbParameter[]? parameters = null, DbTransaction? transaction = null)
        {
            commandText = AdaptSql(commandText);

            if (transaction != null)
            {
                using var cmd = transaction.Connection!.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Length > 0)
                {
                    foreach (var p in parameters)
                    {
                        var param = cmd.CreateParameter();
                        param.ParameterName = p.ParameterName;
                        param.Value = p.Value ?? DBNull.Value;
                        cmd.Parameters.Add(param);
                    }
                }
                int affected = cmd.ExecuteNonQuery();
                if (_currentProvider == DatabaseProvider.Sqlite) SyncSqliteProjectAmounts();
                return affected;
            }
            else
            {
                using var conn = CreateConnection();
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Length > 0)
                {
                    foreach (var p in parameters)
                    {
                        var param = cmd.CreateParameter();
                        param.ParameterName = p.ParameterName;
                        param.Value = p.Value ?? DBNull.Value;
                        cmd.Parameters.Add(param);
                    }
                }
                int affected = cmd.ExecuteNonQuery();
                if (_currentProvider == DatabaseProvider.Sqlite) SyncSqliteProjectAmounts();
                return affected;
            }
        }

        /// <summary>
        /// Выполняет запрос и возвращает скалярный результат
        /// </summary>
        public static object? ExecuteScalar(string commandText, DbParameter[]? parameters = null, DbTransaction? transaction = null)
        {
            commandText = AdaptSql(commandText);

            if (transaction != null)
            {
                using var cmd = transaction.Connection!.CreateCommand();
                cmd.Transaction = transaction;
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Length > 0)
                {
                    foreach (var p in parameters)
                    {
                        var param = cmd.CreateParameter();
                        param.ParameterName = p.ParameterName;
                        param.Value = p.Value ?? DBNull.Value;
                        cmd.Parameters.Add(param);
                    }
                }
                var res = cmd.ExecuteScalar();
                if (_currentProvider == DatabaseProvider.Sqlite) SyncSqliteProjectAmounts();
                return res == DBNull.Value ? null : res;
            }
            else
            {
                using var conn = CreateConnection();
                conn.Open();
                using var cmd = conn.CreateCommand();
                cmd.CommandText = commandText;
                if (parameters != null && parameters.Length > 0)
                {
                    foreach (var p in parameters)
                    {
                        var param = cmd.CreateParameter();
                        param.ParameterName = p.ParameterName;
                        param.Value = p.Value ?? DBNull.Value;
                        cmd.Parameters.Add(param);
                    }
                }
                var res = cmd.ExecuteScalar();
                if (_currentProvider == DatabaseProvider.Sqlite) SyncSqliteProjectAmounts();
                return res == DBNull.Value ? null : res;
            }
        }

        public static object ToDbValue(object? value)
        {
            return value ?? DBNull.Value;
        }

        /// <summary>
        /// Автоматическая синхронизация сумм проектов по пожертвованиям в SQLite
        /// </summary>
        private static void SyncSqliteProjectAmounts()
        {
            try
            {
                using var conn = new SqliteConnection($"Data Source={SqliteDbPath}");
                conn.Open();
                using var cmd = new SqliteCommand(@"
                    UPDATE Projects 
                    SET CurrentAmount = COALESCE((SELECT SUM(Amount) FROM Donations WHERE Donations.ProjectId = Projects.Id), 0.0);
                ", conn);
                cmd.ExecuteNonQuery();
            }
            catch
            {
                // Игнорируем
            }
        }

        /// <summary>
        /// Создание базы данных SQLite и наполнение начальными данными при первом запуске
        /// </summary>
        public static void EnsureSqliteDatabaseCreated()
        {
            string dbFile = SqliteDbPath;
            bool isNew = !File.Exists(dbFile) || new FileInfo(dbFile).Length == 0;

            if (!isNew)
            {
                try
                {
                    using var connCheck = new SqliteConnection($"Data Source={dbFile}");
                    connCheck.Open();
                    using var cmdCheck = new SqliteCommand("PRAGMA table_info(Recipients);", connCheck);
                    using var reader = cmdCheck.ExecuteReader();
                    bool hasReg = false;
                    while (reader.Read())
                    {
                        if (string.Equals(reader["name"]?.ToString(), "RegistrationDate", StringComparison.OrdinalIgnoreCase))
                        {
                            hasReg = true;
                            break;
                        }
                    }
                    reader.Close();
                    if (!hasReg)
                    {
                        using var cmdAlter = new SqliteCommand("ALTER TABLE Recipients ADD COLUMN RegistrationDate TEXT NOT NULL DEFAULT '2026-01-16';", connCheck);
                        cmdAlter.ExecuteNonQuery();
                    }

                    using var cmdExpCheck = new SqliteCommand("PRAGMA table_info(Expenses);", connCheck);
                    using var expReader = cmdExpCheck.ExecuteReader();
                    bool hasExpCat = false;
                    while (expReader.Read())
                    {
                        if (string.Equals(expReader["name"]?.ToString(), "ExpenseCategory", StringComparison.OrdinalIgnoreCase))
                        {
                            hasExpCat = true;
                            break;
                        }
                    }
                    expReader.Close();
                    if (!hasExpCat)
                    {
                        using var cmdAddCat = new SqliteCommand("ALTER TABLE Expenses ADD COLUMN ExpenseCategory TEXT NOT NULL DEFAULT 'Адресная помощь';", connCheck);
                        cmdAddCat.ExecuteNonQuery();
                        using var cmdAddDoc = new SqliteCommand("ALTER TABLE Expenses ADD COLUMN DocumentNumber TEXT NOT NULL DEFAULT 'ПП-001';", connCheck);
                        cmdAddDoc.ExecuteNonQuery();
                        using var cmdAddDesc = new SqliteCommand("ALTER TABLE Expenses ADD COLUMN Description TEXT NULL;", connCheck);
                        cmdAddDesc.ExecuteNonQuery();
                    }
                }
                catch
                {
                    // Игнорируем
                }
                return;
            }

            // Если в корне проекта или рядом есть файл CharityFundDB.db или CharityFundDB.sqlite, скопируем его
            string rootFile1 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "CharityFundDB.db");
            string rootFile2 = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "CharityFundDB.sqlite");
            if (File.Exists(rootFile1) && new FileInfo(rootFile1).Length > 0)
            {
                File.Copy(rootFile1, dbFile, true);
                return;
            }
            if (File.Exists(rootFile2) && new FileInfo(rootFile2).Length > 0)
            {
                File.Copy(rootFile2, dbFile, true);
                return;
            }

            using var conn = new SqliteConnection($"Data Source={dbFile}");
            conn.Open();

            string ddl = @"
                CREATE TABLE IF NOT EXISTS Categories (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL UNIQUE,
                    Description TEXT NULL
                );

                CREATE TABLE IF NOT EXISTS Donors (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    FullName TEXT NOT NULL,
                    DonorType TEXT NOT NULL DEFAULT 'Физическое лицо',
                    Phone TEXT NULL,
                    Email TEXT NULL,
                    Address TEXT NULL,
                    RegistrationDate TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS Projects (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    Name TEXT NOT NULL,
                    CategoryId INTEGER NOT NULL,
                    TargetAmount REAL NOT NULL,
                    CurrentAmount REAL NOT NULL DEFAULT 0.0,
                    StartDate TEXT NOT NULL,
                    EndDate TEXT NULL,
                    Status TEXT NOT NULL DEFAULT 'Активен',
                    Description TEXT NULL,
                    FOREIGN KEY (CategoryId) REFERENCES Categories(Id)
                );

                CREATE TABLE IF NOT EXISTS Recipients (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    FullName TEXT NOT NULL,
                    CategoryId INTEGER NOT NULL,
                    Phone TEXT NULL,
                    Address TEXT NULL,
                    NeedDescription TEXT NOT NULL,
                    Status TEXT NOT NULL DEFAULT 'Одобрено',
                    RegistrationDate TEXT NOT NULL DEFAULT '2026-01-16',
                    CreatedAt TEXT NOT NULL DEFAULT '2026-01-16',
                    FOREIGN KEY (CategoryId) REFERENCES Categories(Id)
                );

                CREATE TABLE IF NOT EXISTS Donations (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    DonorId INTEGER NOT NULL,
                    ProjectId INTEGER NOT NULL,
                    Amount REAL NOT NULL,
                    DonationDate TEXT NOT NULL,
                    PaymentMethod TEXT NOT NULL,
                    Notes TEXT NULL,
                    FOREIGN KEY (DonorId) REFERENCES Donors(Id),
                    FOREIGN KEY (ProjectId) REFERENCES Projects(Id)
                );

                CREATE TABLE IF NOT EXISTS Expenses (
                    Id INTEGER PRIMARY KEY AUTOINCREMENT,
                    ProjectId INTEGER NOT NULL,
                    RecipientId INTEGER NULL,
                    ExpenseCategory TEXT NOT NULL DEFAULT 'Адресная помощь',
                    Amount REAL NOT NULL,
                    ExpenseDate TEXT NOT NULL,
                    DocumentNumber TEXT NOT NULL DEFAULT 'ПП-001',
                    Description TEXT NULL,
                    Purpose TEXT NULL,
                    ResponsiblePerson TEXT NULL,
                    FOREIGN KEY (ProjectId) REFERENCES Projects(Id),
                    FOREIGN KEY (RecipientId) REFERENCES Recipients(Id)
                );

                -- Демо-данные
                INSERT INTO Categories (Name, Description) VALUES
                ('Детское здоровье и лечение', 'Сбор средств на проведение срочных операций, закупку жизненно важных медикаментов и реабилитацию детей.'),
                ('Помощь ветеранам и пожилым людям', 'Социальный патронаж, доставка продуктовых наборов, медикаментов и обустройство быта одиноких пенсионеров.'),
                ('Поддержка детских домов и интернатов', 'Помощь воспитанникам сиротских учреждений, закупка учебных пособий, развивающего оборудования и одежды.'),
                ('Экологические инициативы и защита природы', 'Очистка берегов рек, посадка зеленых насаждений, экологическое просвещение молодежи.'),
                ('Срочная помощь при ЧС и малоимущим', 'Адресная поддержка семей, пострадавших от пожаров, стихийных бедствий или оказавшихся в кризисной ситуации.');

                INSERT INTO Donors (FullName, DonorType, Phone, Email, Address, RegistrationDate) VALUES
                ('ООО «Интерднестрком»', 'Юридическое лицо', '+373-533-91111', 'csr@idknet.com', 'г. Тирасполь, ул. Шелковая, 1', '2026-01-10'),
                ('ЗАО «Тиротекс»', 'Юридическое лицо', '+373-533-72000', 'charity@tirotex.com', 'г. Тирасполь, проезд Магистральный, 10', '2026-01-12'),
                ('Иванов Сергей Викторович', 'Физическое лицо', '+373-777-12345', 's.ivanov@mail.ru', 'г. Бендеры, ул. Ленина, 45, кв. 12', '2026-01-15'),
                ('Петрова Елена Николаевна', 'Физическое лицо', '+373-778-54321', 'elena.petrova@gmail.com', 'г. Тирасполь, ул. 25 Октября, 104, кв. 3', '2026-01-18'),
                ('Анонимный меценат', 'Анонимный благотворитель', NULL, NULL, NULL, '2026-01-20'),
                ('ЗАО «Квинт»', 'Юридическое лицо', '+373-533-85555', 'info@kvint.md', 'г. Тирасполь, ул. Ленина, 38', '2026-01-25'),
                ('Кузнецов Дмитрий Анатольевич', 'Физическое лицо', '+373-779-99887', 'd.kuznetsov@yandex.ru', 'г. Рыбница, ул. Кирова, 19, кв. 5', '2026-02-01');

                INSERT INTO Projects (Name, CategoryId, TargetAmount, CurrentAmount, StartDate, EndDate, Status, Description) VALUES
                ('Здоровое сердце — детям', 1, 150000.0, 95000.0, '2026-01-15', '2026-12-31', 'Активен', 'Оплата кардиохирургических операций и восстановительного лечения для детей со сложными пороками сердца.'),
                ('Теплый дом для ветеранов', 2, 80000.0, 28500.0, '2026-02-01', '2026-11-30', 'Активен', 'Обеспечение одиноких ветеранов войны и труда дровами, теплыми вещами, обогревателями и продуктами.'),
                ('Светлое будущее воспитанников', 3, 60000.0, 26000.0, '2026-03-01', '2026-08-31', 'Активен', 'Оснащение компьютерного класса и творческой мастерской в Республиканском интернате.'),
                ('Чистые берега Днестра 2026', 4, 30000.0, 20000.0, '2026-04-10', '2026-09-30', 'Активен', 'Организация волонтерских эко-субботников, установка мусорных контейнеров и вывоз отходов из прибрежной зоны.'),
                ('Новогодняя сказка для малышей', 3, 45000.0, 45000.0, '2025-11-01', '2026-01-10', 'Завершен', 'Праздничные подарки, представления и зимняя обувь для детей из малообеспеченных семей.');

                INSERT INTO Recipients (FullName, CategoryId, Phone, Address, NeedDescription, Status, CreatedAt) VALUES
                ('Волков Максим (7 лет)', 1, '+373-777-44112', 'г. Бендеры, ул. Суворова, 15', 'Требуется дорогостоящая эндоваскулярная операция по устранению дефекта межпредсердной перегородки.', 'Одобрено', '2026-01-16'),
                ('Григорьева Анна Степановна (83 года)', 2, '+373-533-31245', 'г. Тирасполь, пер. Западный, 4', 'Одинокая пенсионерка, инвалид II группы. Необходимы слуховой аппарат и лекарства для суставов.', 'Одобрено', '2026-02-03'),
                ('МОУ «Школа-интернат для детей-сирот»', 3, '+373-555-41200', 'г. Рыбница, ул. Гвардейская, 30', 'Требуется замена кухонного технологического оборудования в детской столовой.', 'Одобрено', '2026-03-02'),
                ('Семья Морозовых (многодетная, 5 детей)', 5, '+373-778-90123', 'г. Слободзея, ул. Фрунзе, 67', 'Сгорела крыша жилого дома. Необходимы строительные материалы и теплая одежда.', 'Одобрено', '2026-04-05'),
                ('Ильин Роман (12 лет)', 1, '+373-779-11223', 'г. Дубоссары, ул. Ломоносова, 8', 'Реабилитационный курс после тяжелой травмы позвоночника.', 'Помощь оказана', '2026-01-20');

                INSERT INTO Donations (DonorId, ProjectId, Amount, DonationDate, PaymentMethod, Notes) VALUES
                (1, 1, 45000.0, '2026-01-20 10:15:00', 'Банковский перевод', 'Целевой корпоративный благотворительный взнос на программу кардиохирургии'),
                (2, 1, 30000.0, '2026-02-05 14:30:00', 'Банковский перевод', 'Пожертвование от коллектива предприятия'),
                (3, 1, 5000.0, '2026-02-12 18:20:00', 'Банковская карта', 'Для лечения маленького Максима Волкова'),
                (4, 2, 3500.0, '2026-02-15 11:45:00', 'Банковская карта', 'На покупку медикаментов для ветеранов'),
                (1, 2, 25000.0, '2026-02-28 09:10:00', 'Банковский перевод', 'Поддержка ветеранской программы фонда'),
                (6, 3, 20000.0, '2026-03-10 16:00:00', 'Банковский перевод', 'На оборудование компьютерного класса в интернате'),
                (5, 1, 15000.0, '2026-03-15 12:00:00', 'Наличные', 'Анонимное пожертвование на детское отделение'),
                (7, 4, 8000.0, '2026-04-15 13:20:00', 'Банковская карта', 'На закупку инвентаря для эко-волонтеров'),
                (2, 4, 12000.0, '2026-04-20 15:10:00', 'Банковский перевод', 'Экологический грант компании'),
                (3, 3, 6000.0, '2026-05-02 17:35:00', 'Банковская карта', 'На книги и спортивный инвентарь для сирот');

                INSERT INTO Expenses (ProjectId, RecipientId, ExpenseCategory, Amount, ExpenseDate, DocumentNumber, Description, Purpose, ResponsiblePerson) VALUES
                (1, 1, 'Адресная помощь', 40000.0, '2026-02-10 11:00:00', 'ПП-101', 'Частичная предоплата высокотехнологичной кардиохирургической операции', 'Частичная предоплата операции', 'Завьялова Т.П.'),
                (2, 2, 'Закупка медикаментов', 12000.0, '2026-03-05 15:30:00', 'ПП-102', 'Приобретение слухового аппарата и годового запаса медикаментов', 'Приобретение слухового аппарата', 'Николаев К.В.'),
                (3, 3, 'Оборудование', 18500.0, '2026-03-25 14:00:00', 'ПП-103', 'Закупка 5 современных компьютеров для учебного класса интерната', 'Закупка компьютеров', 'Завьялова Т.П.'),
                (4, NULL, 'Транспорт', 6500.0, '2026-04-25 10:20:00', 'ПП-104', 'Закупка перчаток, мешков для мусора и аренда спецтранспорта', 'Закупка инвентаря и транспорт', 'Григорьев А.С.'),
                (1, 1, 'Адресная помощь', 35000.0, '2026-05-15 12:45:00', 'ПП-105', 'Окончательный расчет за операцию и реабилитацию', 'Расчет за операцию', 'Завьялова Т.П.');
            ";

            using var cmdInit = new SqliteCommand(ddl, conn);
            cmdInit.ExecuteNonQuery();
        }
    }
}
