-- ============================================================================
-- Скрипт создания базы данных «CharityFundDB»
-- Тема: «Информационная система Благотворительный фонд»
-- Разработчик: Афанасьев М.А., студент 2 курса, гр. ФТ24ДР62ПИ
-- Профиль: «Разработка программно-информационных систем», кафедра ПОВТ
-- ГОУ «Приднестровский государственный университет им. Т.Г. Шевченко», 2026 г.
-- ============================================================================

USE master;
GO

IF DB_ID('CharityFundDB') IS NULL
BEGIN
    CREATE DATABASE CharityFundDB COLLATE Cyrillic_General_100_CI_AS;
    PRINT N'База данных CharityFundDB успешно создана.';
END
ELSE
BEGIN
    PRINT N'База данных CharityFundDB уже существует.';
END
GO

USE CharityFundDB;
GO

-- 1. Удаление существующих объектов (хранимые процедуры, представления, триггеры, таблицы)
IF OBJECT_ID('sp_GetCategoryHelpSummary', 'P') IS NOT NULL DROP PROCEDURE sp_GetCategoryHelpSummary;
IF OBJECT_ID('sp_GetFundBalanceSummary', 'P') IS NOT NULL DROP PROCEDURE sp_GetFundBalanceSummary;
IF OBJECT_ID('sp_GetTopDonors', 'P') IS NOT NULL DROP PROCEDURE sp_GetTopDonors;
IF OBJECT_ID('sp_GetActiveProjects', 'P') IS NOT NULL DROP PROCEDURE sp_GetActiveProjects;
IF OBJECT_ID('sp_GetDonationsSummary', 'P') IS NOT NULL DROP PROCEDURE sp_GetDonationsSummary;

IF OBJECT_ID('vw_CategoryAssistance', 'V') IS NOT NULL DROP VIEW vw_CategoryAssistance;
IF OBJECT_ID('vw_FundExpenses', 'V') IS NOT NULL DROP VIEW vw_FundExpenses;
IF OBJECT_ID('vw_TopDonors', 'V') IS NOT NULL DROP VIEW vw_TopDonors;
IF OBJECT_ID('vw_ActiveProjects', 'V') IS NOT NULL DROP VIEW vw_ActiveProjects;
IF OBJECT_ID('vw_DonationsSummary', 'V') IS NOT NULL DROP VIEW vw_DonationsSummary;
IF OBJECT_ID('FundExpenses', 'V') IS NOT NULL DROP VIEW FundExpenses;

IF OBJECT_ID('TR_Donations_UpdateCurrentAmount', 'TR') IS NOT NULL DROP TRIGGER TR_Donations_UpdateCurrentAmount;

IF OBJECT_ID('Expenses', 'U') IS NOT NULL DROP TABLE Expenses;
IF OBJECT_ID('Donations', 'U') IS NOT NULL DROP TABLE Donations;
IF OBJECT_ID('Recipients', 'U') IS NOT NULL DROP TABLE Recipients;
IF OBJECT_ID('Projects', 'U') IS NOT NULL DROP TABLE Projects;
IF OBJECT_ID('Donors', 'U') IS NOT NULL DROP TABLE Donors;
IF OBJECT_ID('Categories', 'U') IS NOT NULL DROP TABLE Categories;
GO

-- ============================================================================
-- 2. СОЗДАНИЕ ТАБЛИЦ И ОГРАНИЧЕНИЙ (DDL)
-- ============================================================================

-- 2.1. Таблица: Categories (Категории программ и помощи)
CREATE TABLE Categories (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Name NVARCHAR(150) NOT NULL UNIQUE,
    Description NVARCHAR(500) NULL
);
GO

-- 2.2. Таблица: Donors (Доноры / Благотворители)
CREATE TABLE Donors (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    FullName NVARCHAR(200) NOT NULL,
    DonorType NVARCHAR(50) NOT NULL DEFAULT N'Физическое лицо', -- 'Физическое лицо' или 'Юридическое лицо'
    Phone NVARCHAR(30) NULL,
    Email NVARCHAR(100) NULL,
    Address NVARCHAR(250) NULL,
    RegistrationDate DATETIME2 NOT NULL DEFAULT GETDATE(),
    CreatedAt AS RegistrationDate -- Совместимость со спецификацией
);
GO

-- 2.3. Таблица: Projects (Благотворительные программы и проекты)
CREATE TABLE Projects (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    Name NVARCHAR(200) NOT NULL,
    CategoryId INT NOT NULL,
    TargetAmount DECIMAL(18,2) NOT NULL CHECK (TargetAmount > 0),
    CurrentAmount DECIMAL(18,2) NOT NULL DEFAULT 0.00 CHECK (CurrentAmount >= 0),
    StartDate DATE NOT NULL,
    EndDate DATE NULL,
    Status NVARCHAR(50) NOT NULL DEFAULT N'Активен', -- 'Активен', 'Завершен', 'Приостановлен'
    Description NVARCHAR(1000) NULL,
    Title AS Name, -- Совместимость со спецификацией
    CONSTRAINT FK_Projects_Categories FOREIGN KEY (CategoryId) 
        REFERENCES Categories(Id) ON DELETE NO ACTION ON UPDATE CASCADE,
    CONSTRAINT CK_Projects_Status CHECK (Status IN (N'Активен', N'Завершен', N'Приостановлен')),
    CONSTRAINT CK_Projects_Dates CHECK (EndDate IS NULL OR EndDate >= StartDate)
);
GO

-- 2.4. Таблица: Recipients (Получатели благотворительной помощи)
CREATE TABLE Recipients (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    FullName NVARCHAR(200) NOT NULL,
    CategoryId INT NOT NULL,
    Phone NVARCHAR(30) NULL,
    Address NVARCHAR(250) NULL,
    NeedDescription NVARCHAR(1000) NOT NULL,
    Status NVARCHAR(50) NOT NULL DEFAULT N'На рассмотрении', -- 'На рассмотрении', 'Одобрен', 'Помощь оказана', 'Отклонен'
    RegistrationDate DATE NOT NULL DEFAULT CAST(GETDATE() AS DATE),
    CreatedAt AS CAST(RegistrationDate AS DATETIME2), -- Совместимость со спецификацией
    CONSTRAINT FK_Recipients_Categories FOREIGN KEY (CategoryId) 
        REFERENCES Categories(Id) ON DELETE NO ACTION ON UPDATE CASCADE,
    CONSTRAINT CK_Recipients_Status CHECK (Status IN (N'На рассмотрении', N'Одобрен', N'Помощь оказана', N'Отклонен'))
);
GO

-- 2.5. Таблица: Donations (Пожертвования)
CREATE TABLE Donations (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    DonorId INT NOT NULL,
    ProjectId INT NOT NULL,
    Amount DECIMAL(18,2) NOT NULL CHECK (Amount > 0),
    DonationDate DATETIME2 NOT NULL DEFAULT GETDATE(),
    PaymentMethod NVARCHAR(50) NOT NULL DEFAULT N'Банковская карта', -- 'Банковская карта', 'Банковский перевод', 'Наличные', 'Онлайн-платеж'
    Notes NVARCHAR(500) NULL,
    CONSTRAINT FK_Donations_Donors FOREIGN KEY (DonorId) 
        REFERENCES Donors(Id) ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT FK_Donations_Projects FOREIGN KEY (ProjectId) 
        REFERENCES Projects(Id) ON DELETE CASCADE ON UPDATE CASCADE
);
GO

-- 2.6. Таблица: Expenses / FundExpenses (Расходы фонда / целевая выдача помощи)
CREATE TABLE Expenses (
    Id INT IDENTITY(1,1) PRIMARY KEY,
    ProjectId INT NOT NULL,
    RecipientId INT NULL,
    ExpenseCategory NVARCHAR(100) NOT NULL, -- 'Адресная помощь', 'Закупка медикаментов', 'Оборудование', 'Транспорт', 'Административные расходы'
    Amount DECIMAL(18,2) NOT NULL CHECK (Amount > 0),
    ExpenseDate DATE NOT NULL DEFAULT CAST(GETDATE() AS DATE),
    DocumentNumber NVARCHAR(50) NOT NULL,
    Description NVARCHAR(500) NULL,
    Purpose AS ExpenseCategory, -- Совместимость со спецификацией
    ResponsiblePerson AS DocumentNumber, -- Совместимость со спецификацией
    CONSTRAINT FK_Expenses_Projects FOREIGN KEY (ProjectId) 
        REFERENCES Projects(Id) ON DELETE CASCADE ON UPDATE CASCADE,
    CONSTRAINT FK_Expenses_Recipients FOREIGN KEY (RecipientId) 
        REFERENCES Recipients(Id) ON DELETE SET NULL ON UPDATE CASCADE
);
GO

-- Представление-синоним FundExpenses для полного соответствия именованию ТЗ
CREATE VIEW dbo.FundExpenses AS
SELECT 
    Id,
    ProjectId,
    RecipientId,
    ExpenseCategory,
    Amount,
    ExpenseDate,
    DocumentNumber,
    Description,
    ExpenseCategory AS Purpose,
    DocumentNumber AS ResponsiblePerson
FROM dbo.Expenses;
GO

-- ============================================================================
-- 3. ИНДЕКСЫ ДЛЯ ОПТИМИЗАЦИИ ВЫБОРОК И ПОИСКА
-- ============================================================================
CREATE NONCLUSTERED INDEX IX_Donations_ProjectId ON Donations(ProjectId);
CREATE NONCLUSTERED INDEX IX_Donations_DonorId ON Donations(DonorId);
CREATE NONCLUSTERED INDEX IX_Donations_Date ON Donations(DonationDate);
CREATE NONCLUSTERED INDEX IX_Expenses_ProjectId ON Expenses(ProjectId);
CREATE NONCLUSTERED INDEX IX_Expenses_RecipientId ON Expenses(RecipientId);
CREATE NONCLUSTERED INDEX IX_Expenses_Date ON Expenses(ExpenseDate);
CREATE NONCLUSTERED INDEX IX_Projects_Status ON Projects(Status);
CREATE NONCLUSTERED INDEX IX_Projects_CategoryId ON Projects(CategoryId);
CREATE NONCLUSTERED INDEX IX_Recipients_CategoryId ON Recipients(CategoryId);
CREATE NONCLUSTERED INDEX IX_Recipients_Status ON Recipients(Status);
GO

-- ============================================================================
-- 4. ТРИГГЕРЫ (Автоматический пересчет CurrentAmount в Projects)
-- ============================================================================
CREATE TRIGGER dbo.TR_Donations_UpdateCurrentAmount
ON dbo.Donations
AFTER INSERT, UPDATE, DELETE
AS
BEGIN
    SET NOCOUNT ON;

    ;WITH AffectedProjects AS (
        SELECT ProjectId FROM inserted
        UNION
        SELECT ProjectId FROM deleted
    )
    UPDATE p
    SET p.CurrentAmount = ISNULL((
        SELECT SUM(d.Amount)
        FROM dbo.Donations d
        WHERE d.ProjectId = p.Id
    ), 0.00)
    FROM dbo.Projects p
    INNER JOIN AffectedProjects a ON p.Id = a.ProjectId;
END;
GO

-- ============================================================================
-- 5. ПРЕДСТАВЛЕНИЯ (VIEWS) ДЛЯ 5 ОБЯЗАТЕЛЬНЫХ АНАЛИТИЧЕСКИХ ЗАПРОСОВ
-- ============================================================================

-- Запрос 1: Сумма пожертвований (сводка по проектам)
CREATE VIEW vw_DonationsSummary AS
SELECT 
    p.Id AS ProjectId,
    p.Name AS ProjectName,
    c.Name AS CategoryName,
    COUNT(d.Id) AS DonationsCount,
    ISNULL(SUM(d.Amount), 0.00) AS TotalDonated,
    p.TargetAmount,
    MIN(d.DonationDate) AS FirstDonationDate,
    MAX(d.DonationDate) AS LastDonationDate
FROM Projects p
INNER JOIN Categories c ON p.CategoryId = c.Id
LEFT JOIN Donations d ON p.Id = d.ProjectId
GROUP BY p.Id, p.Name, c.Name, p.TargetAmount;
GO

-- Запрос 2: Активные проекты и процент выполнения сбора
CREATE VIEW vw_ActiveProjects AS
SELECT 
    p.Id AS ProjectId,
    p.Name AS ProjectName,
    c.Name AS CategoryName,
    p.TargetAmount,
    p.CurrentAmount,
    CASE 
        WHEN p.TargetAmount > 0 THEN ROUND((p.CurrentAmount / p.TargetAmount) * 100.0, 2)
        ELSE 0.00 
    END AS CompletionPercentage,
    CASE 
        WHEN p.TargetAmount > p.CurrentAmount THEN (p.TargetAmount - p.CurrentAmount)
        ELSE 0.00 
    END AS RemainingAmount,
    p.StartDate,
    p.EndDate,
    p.Status
FROM Projects p
INNER JOIN Categories c ON p.CategoryId = c.Id
WHERE p.Status = N'Активен';
GO

-- Запрос 3: Крупнейшие доноры (Рейтинг по общей сумме пожертвований)
CREATE VIEW vw_TopDonors AS
SELECT 
    d.Id AS DonorId,
    d.FullName AS DonorName,
    d.DonorType,
    d.Phone,
    d.Email,
    COUNT(dn.Id) AS DonationsCount,
    ISNULL(SUM(dn.Amount), 0.00) AS TotalDonated,
    ISNULL(MAX(dn.Amount), 0.00) AS MaxSingleDonation,
    MAX(dn.DonationDate) AS LastDonationDate
FROM Donors d
INNER JOIN Donations dn ON d.Id = dn.DonorId
GROUP BY d.Id, d.FullName, d.DonorType, d.Phone, d.Email;
GO

-- Запрос 4: Расходы фонда (Сравнение собранных и израсходованных средств)
CREATE VIEW vw_FundExpenses AS
SELECT 
    p.Id AS ProjectId,
    p.Name AS ProjectName,
    p.CurrentAmount AS TotalCollected,
    ISNULL(e.TotalSpent, 0.00) AS TotalSpent,
    (p.CurrentAmount - ISNULL(e.TotalSpent, 0.00)) AS Balance,
    ISNULL(e.ExpensesCount, 0) AS ExpensesCount
FROM Projects p
LEFT JOIN (
    SELECT 
        ProjectId, 
        SUM(Amount) AS TotalSpent, 
        COUNT(Id) AS ExpensesCount 
    FROM Expenses 
    GROUP BY ProjectId
) e ON p.Id = e.ProjectId;
GO

-- Запрос 5: Помощь по категориям
CREATE VIEW vw_CategoryAssistance AS
SELECT 
    c.Id AS CategoryId,
    c.Name AS CategoryName,
    COUNT(DISTINCT p.Id) AS ProjectsCount,
    COUNT(DISTINCT r.Id) AS RecipientsCount,
    ISNULL((SELECT SUM(dn.Amount) 
            FROM Donations dn 
            INNER JOIN Projects pr ON dn.ProjectId = pr.Id 
            WHERE pr.CategoryId = c.Id), 0.00) AS TotalDonated,
    ISNULL((SELECT SUM(ex.Amount) 
            FROM Expenses ex 
            INNER JOIN Projects pr ON ex.ProjectId = pr.Id 
            WHERE pr.CategoryId = c.Id), 0.00) AS TotalExpenses
FROM Categories c
LEFT JOIN Projects p ON c.Id = p.CategoryId
LEFT JOIN Recipients r ON c.Id = r.CategoryId
GROUP BY c.Id, c.Name;
GO

-- ============================================================================
-- 6. ХРАНИМЫЕ ПРОЦЕДУРЫ ДЛЯ АНАЛИТИЧЕСКИХ ВЫБОРОК С ПАРАМЕТРАМИ
-- ============================================================================

-- Процедура 1: Сумма пожертвований за период
CREATE PROCEDURE dbo.sp_GetDonationsSummary
    @StartDate DATETIME2 = NULL,
    @EndDate DATETIME2 = NULL
AS
BEGIN
    SET NOCOUNT ON;

    SELECT 
        p.Id AS ProjectId,
        p.Name AS ProjectName,
        c.Name AS CategoryName,
        p.TargetAmount,
        COUNT(d.Id) AS DonationsCount,
        ISNULL(SUM(d.Amount), 0.00) AS TotalDonated,
        MIN(d.DonationDate) AS FirstDonationDate,
        MAX(d.DonationDate) AS LastDonationDate
    FROM Projects p
    INNER JOIN Categories c ON p.CategoryId = c.Id
    LEFT JOIN Donations d ON p.Id = d.ProjectId
        AND (@StartDate IS NULL OR d.DonationDate >= @StartDate)
        AND (@EndDate IS NULL OR d.DonationDate <= @EndDate)
    GROUP BY p.Id, p.Name, c.Name, p.TargetAmount
    ORDER BY TotalDonated DESC;
END;
GO

-- Процедура 2: Активные проекты
CREATE PROCEDURE dbo.sp_GetActiveProjects
AS
BEGIN
    SET NOCOUNT ON;

    SELECT * FROM dbo.vw_ActiveProjects
    ORDER BY CompletionPercentage DESC;
END;
GO

-- Процедура 3: Крупнейшие доноры (ТОП-N)
CREATE PROCEDURE dbo.sp_GetTopDonors
    @TopCount INT = 10
AS
BEGIN
    SET NOCOUNT ON;

    SELECT TOP (@TopCount) *
    FROM dbo.vw_TopDonors
    ORDER BY TotalDonated DESC;
END;
GO

-- Процедура 4: Сводка баланса фонда
CREATE PROCEDURE dbo.sp_GetFundBalanceSummary
AS
BEGIN
    SET NOCOUNT ON;

    SELECT * FROM dbo.vw_FundExpenses
    ORDER BY Balance DESC;
END;
GO

-- Процедура 5: Помощь по категориям
CREATE PROCEDURE dbo.sp_GetCategoryHelpSummary
AS
BEGIN
    SET NOCOUNT ON;

    SELECT 
        c.Id AS CategoryId,
        c.Name AS CategoryName,
        COUNT(DISTINCT p.Id) AS ProjectsCount,
        COUNT(DISTINCT r.Id) AS RecipientsCount,
        ISNULL((SELECT SUM(dn.Amount) 
                FROM Donations dn 
                INNER JOIN Projects pr ON dn.ProjectId = pr.Id 
                WHERE pr.CategoryId = c.Id), 0.00) AS TotalDonated,
        ISNULL((SELECT SUM(ex2.Amount) 
                FROM Expenses ex2 
                INNER JOIN Projects pr ON ex2.ProjectId = pr.Id 
                WHERE pr.CategoryId = c.Id), 0.00) AS TotalExpenses,
        CAST(ROUND(
            CASE 
                WHEN (SELECT ISNULL(SUM(Amount), 0) FROM Expenses) > 0 
                THEN (ISNULL((SELECT SUM(ex3.Amount) 
                              FROM Expenses ex3 
                              INNER JOIN Projects pr ON ex3.ProjectId = pr.Id 
                              WHERE pr.CategoryId = c.Id), 0.00) * 100.0 / (SELECT SUM(Amount) FROM Expenses))
                ELSE 0.00 
            END, 2) AS DECIMAL(6,2)
        ) AS ExpenseSharePercent
    FROM Categories c
    LEFT JOIN Projects p ON c.Id = p.CategoryId
    LEFT JOIN Recipients r ON c.Id = r.CategoryId
    GROUP BY c.Id, c.Name
    ORDER BY TotalExpenses DESC;
END;
GO

-- ============================================================================
-- 7. ПЕРВИЧНОЕ НАПОЛНЕНИЕ ДЕМОНСТРАЦИОННЫМИ ДАННЫМИ
-- ============================================================================
PRINT N'Заполнение таблиц демонстрационными данными...';
GO

-- 7.1. Категории благотворительности
INSERT INTO Categories (Name, Description) VALUES
(N'Детское здоровье', N'Помощь тяжелобольным детям, оплата дорогостоящих операций и реабилитаций'),
(N'Поддержка малоимущих', N'Продуктовые наборы, теплая одежда и базовые средства жизнеобеспечения'),
(N'Медицинское оборудование', N'Приобретение диагностического и реанимационного оборудования для больниц'),
(N'Помощь пожилым и ветеранам', N'Адресный уход, покупка лекарств и содействие ветеранам и одиноким пенсионерам'),
(N'Приюты для животных', N'Корм, ветеринарная помощь и обустройство приютов для бездомных животных');
GO

-- 7.2. Доноры / Благотворители
INSERT INTO Donors (FullName, DonorType, Phone, Email, Address, RegistrationDate) VALUES
(N'ООО «Интерднестрком»', N'Юридическое лицо', N'+373 (533) 9-55-55', N'charity@idc.md', N'г. Тирасполь, ул. Правды, 10', '2026-01-10 10:00:00'),
(N'ЗАО «Тиротекс»', N'Юридическое лицо', N'+373 (533) 7-30-00', N'contact@tirotex.com', N'г. Тирасполь, проезд Магистральный, 1', '2026-01-12 11:30:00'),
(N'Смирнов Андрей Викторович', N'Физическое лицо', N'+373 (777) 12-345', N'smirnov.a@gmail.com', N'г. Тирасполь, ул. 25 Октября, 42', '2026-01-15 14:20:00'),
(N'Ковалева Наталья Ивановна', N'Физическое лицо', N'+373 (778) 54-321', N'kovaleva_n@mail.ru', N'г. Бендеры, ул. Суворова, 15', '2026-02-01 09:15:00'),
(N'ООО «Шериф»', N'Юридическое лицо', N'+373 (533) 8-88-88', N'support@sheriff.md', N'г. Тирасполь, ул. Шевченко, 99', '2026-02-05 16:40:00'),
(N'Васильев Петр Сергеевич', N'Физическое лицо', N'+373 (775) 99-887', N'vasiliev.p@yandex.ru', N'г. Дубоссары, ул. Ленина, 24', '2026-02-18 17:10:00'),
(N'Мельникова Ольга Дмитриевна', N'Физическое лицо', N'+373 (779) 33-211', N'melnikova_o@inbox.ru', N'г. Рыбница, ул. Кирова, 8', '2026-03-01 10:45:00');
GO

-- 7.3. Благотворительные проекты
INSERT INTO Projects (Name, CategoryId, TargetAmount, CurrentAmount, StartDate, EndDate, Status, Description) VALUES
(N'Здоровое сердце детям', 1, 150000.00, 115000.00, '2026-01-15', '2026-12-31', N'Активен', N'Сбор средств на проведение кардиохирургических операций детям Приднестровья'),
(N'Теплый дом для ветеранов', 4, 60000.00, 48000.00, '2026-02-01', '2026-09-30', N'Активен', N'Ремонт жилья и закупка твердого топлива для одиноких пожилых граждан'),
(N'Новое дыхание клиники', 3, 220000.00, 195000.00, '2026-03-01', '2026-11-15', N'Активен', N'Покупка аппаратов искусственной вентиляции легких для Республиканской клинической больницы'),
(N'Светлая пасха малоимущим', 2, 40000.00, 40000.00, '2026-03-15', '2026-05-10', N'Завершен', N'Формирование и раздача 500 продуктовых наборов к празднику Пасхи'),
(N'Друг человека: помощь приюту', 5, 35000.00, 21500.00, '2026-04-01', '2026-10-31', N'Активен', N'Строительство теплых вольеров и стерилизация животных в городском приюте');
GO

-- 7.4. Получатели помощи
INSERT INTO Recipients (FullName, CategoryId, Phone, Address, NeedDescription, Status, RegistrationDate) VALUES
(N'Иванов Максим (8 лет)', 1, N'+373 (777) 45-678', N'г. Тирасполь, ул. Мира, 12, кв. 4', N'Врожденный порок сердца, требуется окклюдер и оперативное лечение', N'Одобрен', '2026-02-10'),
(N'Григорьева Анна Павловна', 4, N'+373 (533) 4-11-22', N'г. Бендеры, ул. Калинина, 33', N'Инвалид II группы, одинокая, нуждается в замене окон и отопительного котла', N'Помощь оказана', '2026-02-20'),
(N'Республиканская клиническая больница', 3, N'+373 (533) 9-22-11', N'г. Тирасполь, ул. Мира, 33', N'Отделение интенсивной терапии, требуется монитор пациента и инфузоматы', N'Одобрен', '2026-03-05'),
(N'Семья Сидоровых (многодетные)', 2, N'+373 (778) 90-123', N'г. Слободзея, ул. Советская, 88', N'5 несовершеннолетних детей, отец потерял работу, нужны продукты и одежда к школе', N'Помощь оказана', '2026-03-18'),
(N'Приют для животных «Шанс»', 5, N'+373 (775) 66-554', N'г. Тирасполь, пер. Западный, 5', N'Необходимы средства на покупку медикаментов для вакцинации 80 собак', N'Одобрен', '2026-04-05');
GO

-- 7.5. Пожертвования (Donations)
INSERT INTO Donations (DonorId, ProjectId, Amount, DonationDate, PaymentMethod, Notes) VALUES
(1, 1, 50000.00, '2026-02-01 10:15:00', N'Банковский перевод', N'Корпоративное пожертвование на детскую кардиохирургию'),
(5, 3, 100000.00, '2026-03-10 14:30:00', N'Банковский перевод', N'Благотворительный взнос на закупку ИВЛ'),
(2, 1, 35000.00, '2026-03-12 11:20:00', N'Банковский перевод', N'Поддержка программы детского здоровья'),
(3, 1, 15000.00, '2026-03-25 16:45:00', N'Банковская карта', N'Личные средства на операцию ребенку'),
(4, 2, 20000.00, '2026-04-02 09:10:00', N'Онлайн-платеж', N'Помощь одиноким пожилым людям'),
(1, 2, 28000.00, '2026-04-10 15:00:00', N'Банковский перевод', N'Улучшение жилищных условий ветеранов'),
(5, 4, 40000.00, '2026-04-15 12:00:00', N'Банковский перевод', N'Оплата 500 пасхальных наборов'),
(6, 5, 12000.00, '2026-04-20 18:30:00', N'Банковская карта', N'На корм и обустройство вольеров'),
(7, 5, 9500.00, '2026-05-02 13:15:00', N'Наличные', N'В кассу фонда на медикаменты для приюта'),
(3, 3, 50000.00, '2026-05-18 17:00:00', N'Банковский перевод', N'Вторая часть благотворительного взноса на клинику'),
(2, 3, 45000.00, '2026-05-20 10:30:00', N'Банковский перевод', N'Целевой взнос на оборудование для РКБ'),
(4, 1, 15000.00, '2026-06-01 11:00:00', N'Онлайн-платеж', N'Пожелание скорейшего выздоровления детям');
GO

-- 7.6. Расходы фонда (Expenses / FundExpenses)
INSERT INTO Expenses (ProjectId, RecipientId, ExpenseCategory, Amount, ExpenseDate, DocumentNumber, Description) VALUES
(1, 1, N'Адресная помощь', 45000.00, '2026-02-25', N'ПП-2026-001', N'Оплата операции на сердце и пребывания в специализированном кардиоцентре'),
(4, 4, N'Адресная помощь', 38500.00, '2026-04-22', N'ТТН-54812', N'Закупка и адресная доставка 500 праздничных продуктовых наборов малоимущим'),
(2, 2, N'Адресная помощь', 25000.00, '2026-04-28', N'АКТ-1029', N'Оплата строительных материалов и монтажа металлопластиковых окон ветерану'),
(3, 3, N'Оборудование', 120000.00, '2026-05-25', N'СФ-980123', N'Оплата аванса за поставку современного аппарата искусственной вентиляции легких'),
(5, 5, N'Закупка медикаментов', 14000.00, '2026-05-30', N'НОВ-3041', N'Оплата вакцин и антибиотиков для животных приюта «Шанс»');
GO

PRINT N'============================================================================';
PRINT N'Скрипт CharityFund_CreateDB.sql успешно выполнен!';
PRINT N'База данных CharityFundDB наполнена и полностью готова к эксплуатации.';
PRINT N'============================================================================';
GO
