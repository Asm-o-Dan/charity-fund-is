import os
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def create_report(output_path):
    doc = docx.Document()

    # 1. Настройка параметров страницы (ГОСТ: левое 30 мм, правое 10 мм, верхнее 20 мм, нижнее 20 мм)
    for section in doc.sections:
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(1.0)
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)

    # 2. Настройка базового стиля Normal
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    style_normal.font.size = Pt(14)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)

    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6, space_before=0, line_spacing=1.5, first_indent=1.25, bold=False, italic=False):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_after = Pt(space_after)
        pf.space_before = Pt(space_before)
        pf.line_spacing = line_spacing
        if first_indent > 0:
            pf.first_line_indent = Cm(first_indent)
        else:
            pf.first_line_indent = Cm(0)
        
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
            run.font.bold = bold
            run.font.italic = italic
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.space_before = Pt(16)
        pf.space_after = Pt(12)
        pf.first_line_indent = Cm(0)
        pf.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(12)
        pf.space_after = Pt(8)
        pf.first_line_indent = Cm(1.25)
        pf.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.font.bold = True
        return p

    def add_code_listing(code_text, caption):
        cap = add_p(caption, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=8, space_after=4, first_indent=0, italic=False)
        table = doc.add_table(rows=1, cols=1)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = table.cell(0, 0)
        
        # Границы ячейки
        tcPr = cell._tc.get_or_add_tcPr()
        tcBorders = parse_xml(r'''
            <w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
                <w:top w:val="single" w:sz="6" w:space="0" w:color="A0AEC0"/>
                <w:left w:val="single" w:sz="18" w:space="0" w:color="2563EB"/>
                <w:bottom w:val="single" w:sz="6" w:space="0" w:color="A0AEC0"/>
                <w:right w:val="single" w:sz="6" w:space="0" w:color="A0AEC0"/>
            </w:tcBorders>
        ''')
        tcPr.append(tcBorders)

        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>')
        tcPr.append(shd)

        cp = cell.paragraphs[0]
        cp.paragraph_format.space_before = Pt(4)
        cp.paragraph_format.space_after = Pt(4)
        cp.paragraph_format.line_spacing = 1.05
        cp.paragraph_format.first_line_indent = Cm(0)
        crun = cp.add_run(code_text)
        crun.font.name = 'Consolas'
        crun.font.size = Pt(9.5)
        crun.font.color.rgb = RGBColor(30, 41, 59)
        
        doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # =========================================================================
    # ТИТУЛЬНЫЙ ЛИСТ (в точности по шаблону ГОУ ПГУ)
    # =========================================================================
    add_p("Государственное образовательное учреждение", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, first_indent=0)
    add_p("«Приднестровский государственный университет им. Т.Г. Шевченко»", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, first_indent=0)
    add_p("Физико-технический институт", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, first_indent=0)
    add_p("Факультет информатики и вычислительной техники", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, first_indent=0)
    add_p("Кафедра программного обеспечения вычислительной техники", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=24, first_indent=0)

    add_p("ОТЧЕТ ПО ПРАКТИКЕ", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4, first_indent=0, bold=True)
    add_p("Технологическая (проектно-технологическая) практика", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=14, first_indent=0)

    add_p("Направление 09.03.04 «Программная инженерия»", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2, first_indent=0)
    add_p("Профиль: «Разработка программно-информационных систем»", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18, first_indent=0)

    add_p("Студент 2 курса ФТ24ДР62ПИ группа", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2, first_indent=9.0)
    add_p("форма обучения очная", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2, first_indent=9.0)
    add_p("Афанасьев Михаил Алексеевич", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=14, first_indent=9.0, bold=True)

    add_p("Место прохождения практики «ПГУ им. Т.Г. Шевченко, ФТИ, учебно-вычислительный центр»", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=4, first_indent=0)
    add_p("Сроки прохождения практики", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2, first_indent=0)
    add_p("С «01» июня 2026 г. по «13» июля 2026 г.", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=14, first_indent=0)

    add_p("Руководители практики:", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=2, first_indent=0)
    add_p("От профильной организации", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=1, first_indent=1.25)
    add_p("зав. кафедрой ПОВТ Помян С. В.", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=6, first_indent=1.25, bold=True)
    add_p("От университета", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=1, first_indent=1.25)
    add_p("ст. преподаватель Е.В. Терещенко.", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=18, first_indent=1.25, bold=True)

    add_p("Итоговая оценка (зачет) по практике ______________________________", align=WD_ALIGN_PARAGRAPH.LEFT, space_after=28, first_indent=0)
    add_p("Тирасполь, 2026 г.", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=0, first_indent=0)

    doc.add_page_break()

    # =========================================================================
    # ОГЛАВЛЕНИЕ
    # =========================================================================
    add_p("ОГЛАВЛЕНИЕ", align=WD_ALIGN_PARAGRAPH.CENTER, space_after=18, first_indent=0, bold=True)
    toc_items = [
        ("ВВЕДЕНИЕ", "3"),
        ("1 ОПИСАНИЕ ПРЕДМЕТНОЙ ОБЛАСТИ", "4"),
        ("  1.1 Общие сведения", "4"),
        ("2 ПОСТАНОВКА ЗАДАЧИ", "6"),
        ("3 РУКОВОДСТВО ПРОГРАММИСТА", "8"),
        ("  3.1 Введение", "8"),
        ("  3.2 Общие сведения", "9"),
        ("  3.3 Описание типов данных", "11"),
        ("  3.4 Описание исходных текстов программного продукта", "13"),
        ("  3.5 Вид программного продукта", "16"),
        ("4 РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ", "17"),
        ("  4.1 Введение", "17"),
        ("  4.2 Системные требования", "17"),
        ("  4.3 Запуск и работа с программным продуктом", "18"),
        ("ЗАКЛЮЧЕНИЕ", "23"),
        ("СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ", "24"),
    ]
    for title, pg in toc_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = 1.25
        p.paragraph_format.first_line_indent = Cm(0)
        run_t = p.add_run(title)
        run_t.font.name = 'Times New Roman'
        run_t.font.size = Pt(14)
        if not title.startswith("  "):
            run_t.font.bold = True
        
        # Заполнитель точками
        dots_count = max(5, 75 - len(title) * 2)
        run_d = p.add_run(" " + "." * dots_count + " ")
        run_d.font.name = 'Times New Roman'
        run_d.font.size = Pt(12)
        run_d.font.color.rgb = RGBColor(120, 120, 120)

        run_p = p.add_run(pg)
        run_p.font.name = 'Times New Roman'
        run_p.font.size = Pt(14)
        if not title.startswith("  "):
            run_p.font.bold = True

    doc.add_page_break()

    # =========================================================================
    # ВВЕДЕНИЕ
    # =========================================================================
    add_h1("ВВЕДЕНИЕ")
    add_p("В ходе технологической (проектно-технологической) практики был разработан программный продукт — настольная информационная система «Благотворительный фонд» с графическим интерфейсом на платформе Windows Forms (C# / .NET 8) и хранилищем данных на базе реляционной СУБД Microsoft SQL Server. Программная система предназначена для комплексной автоматизации учета благотворителей (доноров), целевых благотворительных проектов, категорий помощи, финансовых пожертвований, физических и юридических лиц — получателей помощи, а также целевых расходов фонда с предоставлением оперативной и аналитической отчетности.")
    add_p("Целью прохождения практики являлось закрепление и углубление теоретических знаний по объектно-ориентированному программированию, архитектуре программных систем, проектированию реляционных баз данных, а также приобретение практических навыков разработки многокомпонентных настольных приложений на языке C# с использованием классической технологии прямого доступа к данным ADO.NET (Microsoft.Data.SqlClient).")
    add_p("В процессе разработки были использованы: среда разработки Microsoft Visual Studio, платформа .NET 8, язык программирования C#, графическая подсистема Windows Forms, система управления базами данных Microsoft SQL Server (Transact-SQL), а также библиотека взаимодействия с базами данных Microsoft.Data.SqlClient. Применение параметризованных SQL-запросов и транзакций гарантирует высокую производительность, абсолютную защищенность от атак класса SQL Injection и строгое соблюдение требований ACID к финансовым операциям.")

    doc.add_page_break()

    # =========================================================================
    # 1 ОПИСАНИЕ ПРЕДМЕТНОЙ ОБЛАСТИ
    # =========================================================================
    add_h1("1 ОПИСАНИЕ ПРЕДМЕТНОЙ ОБЛАСТИ")
    add_h2("1.1 Общие сведения")
    add_p("Предметной областью разработанного программного продукта является деятельность некоммерческой благотворительной организации (фонда), аккумулирующей финансовые средства благотворителей для реализации социально значимых программ и оказания адресной поддержки гражданам и организациям, оказавшимся в сложной жизненной ситуации.")
    add_p("Специфика деятельности благотворительного фонда требует строгой финансовой прозрачности, абсолютной адресности каждого израсходованного рубля и возможности оперативного формирования публичной отчетности для доноров, попечительского совета и контролирующих органов. В условиях отсутствия единой автоматизированной системы ведение учета в разрозненных электронных таблицах приводит к риску дублирования информации, затрудняет контроль целевого расходования средств по каждому конкретному проекту и делает процесс построения сводной аналитики трудоемким и уязвимым к человеческим ошибкам.")
    add_p("Разработанная информационная система автоматизирует работу сотрудников фонда с шестью ключевыми взаимосвязанными сущностями:")
    add_p("1. Доноры (Donors) — физические лица и коммерческие организации (юридические лица), осуществляющие добровольные благотворительные пожертвования. Донор характеризуется уникальным идентификатором, ФИО или полным наименованием компании, типом («Физическое лицо» или «Юридическое лицо»), контактным номером телефона, адресом электронной почты, фактическим адресом и датой регистрации.")
    add_p("2. Категории помощи (Categories) — направления благотворительной деятельности фонда, систематизирующие реализуемые инициативы: «Детское здоровье», «Поддержка малоимущих», «Медицинское оборудование», «Помощь пожилым и ветеранам», «Приюты для животных» и др.")
    add_p("3. Проекты (Projects) — конкретные целевые программы сбора средств. Каждый проект привязан к определенной категории, имеет наименование, целевую сумму сбора (TargetAmount), текущую фактически собранную сумму (CurrentAmount), дату начала, плановую дату завершения, текущий статус («Активен», «Завершен», «Приостановлен») и развернутое текстовое описание.")
    add_p("4. Получатели помощи (Recipients) — граждане или учреждения, подавшие заявку на получение благотворительной поддержки. Характеризуются персональными данными, категорией обращения, контактами, детальным описанием жизненной ситуации / потребности и статусом рассмотрения заявки («На рассмотрении», «Одобрен», «Помощь оказана», «Отклонен»).")
    add_p("5. Пожертвования (Donations) — финансовые операции поступления средств от конкретного донора в адрес определенного проекта. Фиксируются сумма, дата и время платежа, способ оплаты («Банковская карта», «Банковский перевод», «Наличные», «Онлайн-платеж») и назначение.")
    add_p("6. Расходы фонда (Expenses) — операции целевого списания средств с баланса проекта на приобретение медикаментов, оплату лечения, закупку оборудования, проведение ремонтных работ или оказание адресной помощи конкретному получателю с обязательной фиксацией номера первичного оправдательного документа (акта, чека, товарной накладной).")
    add_p("Согласно заданию на практику, система предоставляет пять регламентированных аналитических отчетов: 1) Сумма пожертвований (с группировкой по проектам и фильтрацией по диапазону дат); 2) Активные проекты (цель, собрано, процент выполнения, остаток к сбору); 3) Крупнейшие доноры (рейтинг благотворителей по общей сумме взносов); 4) Расходы фонда (сопоставление собранных и израсходованных средств, баланс проекта); 5) Помощь по категориям (распределение сумм помощи и количества получателей по направлениям).")

    doc.add_page_break()

    # =========================================================================
    # 2 ПОСТАНОВКА ЗАДАЧИ
    # =========================================================================
    add_h1("2 ПОСТАНОВКА ЗАДАЧИ")
    add_p("Целью разработки является создание надежного, быстродействующего и наглядного настольного программного комплекса на базе Windows Forms и Microsoft SQL Server, реализующего полный жизненный цикл учета благотворительной деятельности, включая операции CRUD (создание, чтение, модификация, удаление) для всех сущностей базы данных, проведение финансовых операций в рамках транзакций ADO.NET, экспорт данных в форматы CSV/Excel и формирование пяти аналитических срезов.")
    add_p("Актуальность работы обусловлена возрастающими требованиями к финансовой дисциплине и открытости благотворительных организаций перед обществом, необходимостью строгого сопоставления каждого акта расхода с первичными бухгалтерскими документами и предоставлением донорам достоверных отчетов о ходе сбора средств.")
    add_p("Для достижения поставленной цели в рамках технологической практики были решены следующие задачи:")
    add_p("— спроектировать концептуальную, логическую и физическую модели реляционной базы данных «CharityFundDB» в среде Microsoft SQL Server;")
    add_p("— разработать полный SQL-скрипт развертывания базы данных с созданием таблиц Categories, Donors, Projects, Recipients, Donations, Expenses, первичных (PRIMARY KEY) и внешних (FOREIGN KEY) ключей, ограничений целостности (CHECK, DEFAULT, UNIQUE) и оптимизирующих некластеризованных индексов;")
    add_p("— создать специализированные представления (VIEW) на стороне MS SQL Server для ускоренного формирования пяти регламентированных аналитических выборок;")
    add_p("— разработать архитектуру приложения на C# (.NET 8) с разделением ответственности на слои Models (доменные сущности), DataAccess (ADO.NET репозитории), Services (экспорт в CSV/HTML) и Forms (пользовательский графический интерфейс);")
    add_p("— реализовать класс доступа к данным DatabaseHelper с безопасным выполнением параметризованных команд, транзакционным механизмом SqlTransaction и гибкой конфигурацией через App.config;")
    add_p("— создать удобный эргономичный графический интерфейс с вкладками, строкой поиска, фильтрами по статусам и диапазонам дат, модальными окнами добавления и редактирования записей;")
    add_p("— реализовать двухсторонний контроль финансовых потоков: автоматическое пополнение баланса проекта при регистрации пожертвования и корректное списание средств при удалении взноса в единой неделимой транзакции;")
    add_p("— разработать модуль экспорта табличных данных и аналитических отчетов в CSV (с кодировкой UTF-8 BOM для безупречного открытия в Microsoft Excel) и адаптивный HTML для печати;")
    add_p("— покрыть вычислительную логику и проверку скрипта базы данных набором автоматизированных модульных тестов (xUnit).")

    doc.add_page_break()

    # =========================================================================
    # 3 РУКОВОДСТВО ПРОГРАММИСТА
    # =========================================================================
    add_h1("3 РУКОВОДСТВО ПРОГРАММИСТА")
    add_h2("3.1 Введение")
    add_p("Настоящее руководство предназначено для специалистов-разработчиков, системных интеграторов и администраторов баз данных, осуществляющих установку, сопровождение, тестирование или модификацию информационной системы «Благотворительный фонд».")
    add_p("Для сопровождения программного продукта специалист должен владеть языком программирования C# на уровне понимания объектно-ориентированных концепций, синтаксисом Transact-SQL (T-SQL), понимать принципы работы реляционных СУБД, механизмы изоляции транзакций и основы низкоуровневого взаимодействия с базой данных через ADO.NET.")
    add_p("Для сборки и запуска решения требуются:")
    add_p("— операционная система Microsoft Windows 10/11 или Windows Server 2019/2022;")
    add_p("— пакет средств разработки .NET SDK 8.0 или выше;")
    add_p("— СУБД Microsoft SQL Server 2016/2019/2022 либо редакция Microsoft SQL Server Express / LocalDB;")
    add_p("— среда разработки Microsoft Visual Studio 2022 (с установленной рабочей нагрузкой «Разработка классических приложений .NET») или JetBrains Rider.")

    add_h2("3.2 Общие сведения")
    add_p("Архитектурно программный комплекс спроектирован по модульному принципу с четким разделением функциональных слоев, что гарантирует независимость бизнес-правил от интерфейса и прозрачность каждой операции доступа к данным:")
    add_p("1. Слой моделей (CharityFundApp.Models) — содержит чистые классы C# (POCO), инкапсулирующие состояние объектов предметной области (Category, Donor, Project, Recipient, Donation, Expense), а также DTO-классы для пяти аналитических отчетов.")
    add_p("2. Слой доступа к данным (CharityFundApp.DataAccess) — построен на прямом использовании классического ADO.NET (библиотека Microsoft.Data.SqlClient). Включает статический класс DatabaseHelper для управления пулом соединений и выполнения параметризованных команд, а также репозитории: CategoryRepository, DonorRepository, ProjectRepository, RecipientRepository, DonationRepository, ExpenseRepository и ReportRepository. Отказ от тяжеловесных ORM-фреймворков обусловлен требованием легкости защиты проекта: код предельно прост, прозрачен и легко объясним на уровне каждого SQL-запроса.")
    add_p("3. Слой сервисов (CharityFundApp.Services) — представлен компонентом ExportService, реализующим форматирование данных таблиц в стандартные CSV-файлы с разделителем «точка с запятой» и UTF-8 BOM, а также генерацию стилизованных HTML-отчетов.")
    add_p("4. Слой пользовательского интерфейса (CharityFundApp.Forms) — содержит главное окно MainForm с многостраничным интерфейсом на основе TabControl, модальные диалоговые формы добавления и редактирования записей (DonorEditForm, ProjectEditForm, CategoryEditForm, DonationEditForm, RecipientEditForm, ExpenseEditForm) и окно настройки подключения к СУБД ConnectionSettingsForm.")

    add_h2("3.3 Описание типов данных")
    add_p("База данных «CharityFundDB» содержит шесть реляционных таблиц, связанных отношениями «один-ко-многим». Физическая структура спроектирована с учетом третьей нормальной формы (3NF):")
    add_p("— Таблица Categories: Id (INT IDENTITY PRIMARY KEY) — первичный суррогатный ключ; Name (NVARCHAR(150) NOT NULL UNIQUE) — уникальное наименование направления; Description (NVARCHAR(500) NULL) — пояснение.")
    add_p("— Таблица Donors: Id (INT IDENTITY PRIMARY KEY); FullName (NVARCHAR(200) NOT NULL) — ФИО благотворителя либо наименование организации; DonorType (NVARCHAR(50) NOT NULL) — категория донора; Phone (NVARCHAR(30) NULL); Email (NVARCHAR(100) NULL); Address (NVARCHAR(250) NULL); RegistrationDate (DATETIME2 NOT NULL DEFAULT GETDATE()).")
    add_p("— Таблица Projects: Id (INT IDENTITY PRIMARY KEY); Name (NVARCHAR(200) NOT NULL); CategoryId (INT NOT NULL FK к Categories); TargetAmount (DECIMAL(18,2) NOT NULL CHECK > 0); CurrentAmount (DECIMAL(18,2) NOT NULL DEFAULT 0 CHECK >= 0); StartDate (DATE NOT NULL); EndDate (DATE NULL); Status (NVARCHAR(50) NOT NULL DEFAULT N'Активен'); Description (NVARCHAR(1000) NULL).")
    add_p("— Таблица Recipients: Id (INT IDENTITY PRIMARY KEY); FullName (NVARCHAR(200) NOT NULL); CategoryId (INT NOT NULL FK к Categories); Phone (NVARCHAR(30) NULL); Address (NVARCHAR(250) NULL); NeedDescription (NVARCHAR(1000) NOT NULL); Status (NVARCHAR(50) NOT NULL DEFAULT N'На рассмотрении'); RegistrationDate (DATE NOT NULL DEFAULT CAST(GETDATE() AS DATE)).")
    add_p("— Таблица Donations: Id (INT IDENTITY PRIMARY KEY); DonorId (INT NOT NULL FK к Donors ON DELETE CASCADE); ProjectId (INT NOT NULL FK к Projects ON DELETE CASCADE); Amount (DECIMAL(18,2) NOT NULL CHECK > 0); DonationDate (DATETIME2 NOT NULL DEFAULT GETDATE()); PaymentMethod (NVARCHAR(50) NOT NULL); Notes (NVARCHAR(500) NULL).")
    add_p("— Таблица Expenses: Id (INT IDENTITY PRIMARY KEY); ProjectId (INT NOT NULL FK к Projects ON DELETE CASCADE); RecipientId (INT NULL FK к Recipients ON DELETE SET NULL); ExpenseCategory (NVARCHAR(100) NOT NULL); Amount (DECIMAL(18,2) NOT NULL CHECK > 0); ExpenseDate (DATE NOT NULL); DocumentNumber (NVARCHAR(50) NOT NULL); Description (NVARCHAR(500) NULL).")

    add_h2("3.4 Описание исходных текстов программного продукта")
    add_p("Ключевые фрагменты программного кода, иллюстрирующие доменные сущности, механизм параметризованного доступа к данным и транзакционную обработку пожертвований, приведены в листингах 1, 2 и 3.")

    code_project = """namespace CharityFundApp.Models
{
    public class Project
    {
        public int Id { get; set; }
        public string Name { get; set; } = string.Empty;
        public int CategoryId { get; set; }
        public string? CategoryName { get; set; }
        public decimal TargetAmount { get; set; }
        public decimal CurrentAmount { get; set; }
        public DateTime StartDate { get; set; } = DateTime.Today;
        public DateTime? EndDate { get; set; }
        public string Status { get; set; } = "Активен";
        public string? Description { get; set; }

        public decimal CompletionPercentage =>
            TargetAmount > 0 ? Math.Round((CurrentAmount / TargetAmount) * 100m, 2) : 0m;

        public decimal RemainingAmount =>
            TargetAmount > CurrentAmount ? TargetAmount - CurrentAmount : 0m;
    }
}"""
    add_code_listing(code_project, "Листинг 1 — Доменная сущность благотворительного проекта (Models/Project.cs)")

    code_tx = """public int InsertWithTransaction(Donation donation)
{
    using var conn = DatabaseHelper.CreateConnection();
    conn.Open();
    using var transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);
    try
    {
        // 1. Фиксация факта пожертвования
        string insertSql = @"INSERT INTO Donations 
            (DonorId, ProjectId, Amount, DonationDate, PaymentMethod, Notes)
            VALUES (@DonorId, @ProjectId, @Amount, @DonationDate, @PaymentMethod, @Notes);
            SELECT SCOPE_IDENTITY();";

        int newId;
        using (var cmd = new SqlCommand(insertSql, conn, transaction))
        {
            cmd.Parameters.AddWithValue("@DonorId", donation.DonorId);
            cmd.Parameters.AddWithValue("@ProjectId", donation.ProjectId);
            cmd.Parameters.AddWithValue("@Amount", donation.Amount);
            cmd.Parameters.AddWithValue("@DonationDate", donation.DonationDate);
            cmd.Parameters.AddWithValue("@PaymentMethod", donation.PaymentMethod);
            cmd.Parameters.AddWithValue("@Notes", DatabaseHelper.ToDbValue(donation.Notes));
            newId = Convert.ToInt32(cmd.ExecuteScalar());
        }

        // 2. Атомарное обновление текущего баланса проекта
        string updateSql = @"UPDATE Projects 
            SET CurrentAmount = CurrentAmount + @Amount 
            WHERE Id = @ProjectId;";

        using (var cmd = new SqlCommand(updateSql, conn, transaction))
        {
            cmd.Parameters.AddWithValue("@Amount", donation.Amount);
            cmd.Parameters.AddWithValue("@ProjectId", donation.ProjectId);
            cmd.ExecuteNonQuery();
        }

        transaction.Commit();
        return newId;
    }
    catch
    {
        transaction.Rollback();
        throw;
    }
}"""
    add_code_listing(code_tx, "Листинг 2 — Транзакционная регистрация пожертвования в ADO.NET (DataAccess/DonationRepository.cs)")

    code_rep = """public List<ActiveProjectReportItem> GetActiveProjects()
{
    var list = new List<ActiveProjectReportItem>();
    string sql = @"
        SELECT 
            p.Id AS ProjectId, p.Name AS ProjectName, c.Name AS CategoryName,
            p.TargetAmount, p.CurrentAmount,
            CASE WHEN p.TargetAmount > 0 
                 THEN ROUND((p.CurrentAmount / p.TargetAmount) * 100.0, 2)
                 ELSE 0.00 END AS CompletionPercentage,
            CASE WHEN p.TargetAmount > p.CurrentAmount 
                 THEN (p.TargetAmount - p.CurrentAmount)
                 ELSE 0.00 END AS RemainingAmount,
            p.StartDate, p.EndDate, p.Status
        FROM Projects p
        INNER JOIN Categories c ON p.CategoryId = c.Id
        WHERE p.Status = N'Активен'
        ORDER BY p.CurrentAmount DESC;";

    var dt = DatabaseHelper.ExecuteQuery(sql);
    foreach (DataRow row in dt.Rows)
    {
        list.Add(new ActiveProjectReportItem {
            ProjectId = Convert.ToInt32(row["ProjectId"]),
            ProjectName = row["ProjectName"].ToString() ?? "",
            CategoryName = row["CategoryName"].ToString() ?? "",
            TargetAmount = Convert.ToDecimal(row["TargetAmount"]),
            CurrentAmount = Convert.ToDecimal(row["CurrentAmount"]),
            CompletionPercentage = Convert.ToDecimal(row["CompletionPercentage"]),
            RemainingAmount = Convert.ToDecimal(row["RemainingAmount"]),
            StartDate = Convert.ToDateTime(row["StartDate"]),
            EndDate = row["EndDate"] == DBNull.Value ? null : Convert.ToDateTime(row["EndDate"]),
            Status = row["Status"].ToString() ?? "Активен"
        });
    }
    return list;
}"""
    add_code_listing(code_rep, "Листинг 3 — Формирование отчета «Активные проекты» (DataAccess/ReportRepository.cs)")

    add_h2("3.5 Вид программного продукта")
    add_p("Программный продукт представляет собой полнофункциональное настольное MDI/SDI-приложение с современным интерфейсом Windows Forms, оформленным в лаконичной корпоративной стилистике (светлый фон, контрастные заголовки Slate 800, цветовые индикаторы статусов).")
    add_p("Главное окно MainForm разделено на функциональные зоны: 1) Шапка приложения с реквизитами разработчика, кнопкой принудительного обновления и переходом к параметрам базы данных; 2) Многостраничный блок вкладок TabControl, содержащий разделы «Доноры», «Проекты фонда», «Категории», «Пожертвования», «Получатели помощи», «Расходы фонда» и «Аналитические запросы и отчеты»; 3) Нижняя строка состояния StatusStrip с индикатором соединения с MS SQL Server и счетчиком записей.")

    doc.add_page_break()

    # =========================================================================
    # 4 РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ
    # =========================================================================
    add_h1("4 РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ")
    add_h2("4.1 Введение")
    add_p("Настоящее руководство предназначено для сотрудников и волонтеров благотворительного фонда, выполняющих ведение базы доноров, проектов и подопечных, регистрацию входящих пожертвований и целевых расходов, а также формирование аналитических справок.")
    add_p("Для работы с информационной системой пользователю достаточно обладать базовыми навыками работы в среде Microsoft Windows. Система оснащена понятным русскоязычным интерфейсом, валидацией всех вводимых полей и защитой от случайного удаления взаимосвязанных записей.")

    add_h2("4.2 Системные требования")
    add_p("Для стабильного функционирования программного продукта рабочая станция пользователя должна соответствовать следующим минимальным требованиям:")
    add_p("— процессор с тактовой частотой не менее 1.6 ГГц (поддержка x64/x86);")
    add_p("— объем оперативной памяти (RAM) не менее 2 ГБ;")
    add_p("— свободное дисковое пространство не менее 150 МБ для исполняемых файлов программы и локальных отчетов;")
    add_p("— операционная система Microsoft Windows 10, Windows 11 (версия 21H2 и выше);")
    add_p("— установленный компонент .NET Desktop Runtime 8.0;")
    add_p("— доступ по локальной сети или на локальном компьютере к СУБД Microsoft SQL Server 2016/2019/2022 или SQL Server Express.")

    add_h2("4.3 Запуск и работа с программным продуктом")
    add_p("Порядок первоначального запуска и типовые сценарии работы с системой:")
    add_p("Шаг 1. Первоначальный запуск и подключение к БД. Запустите исполняемый файл CharityFundApp.exe. При первом старте система попытается подключиться к локальному серверу MS SQL Server по строке подключения из App.config. Если сервер еще не сконфигурирован, в заголовке окна нажмите «Параметры БД». В открывшемся окне можно выбрать готовый шаблон (localhost, .\\SQLEXPRESS, LocalDB), нажать «Проверить подключение» и в один клик выполнить кнопку «⚡ Развернуть БД из скрипта» — программа автоматически создаст базу CharityFundDB и наполнит ее демонстрационными данными.")
    add_p("Сценарий 1. Ведение реестра доноров. Перейдите на вкладку «👥 Доноры». В таблице отображается список всех зарегистрированных благотворителей. В верхней панели доступен быстрый поиск по части имени, телефона или почты. Для внесения нового благотворителя нажмите кнопку «➕ Добавить», укажите имя, тип лица (физическое/юридическое) и контактные данные. Для изменения выделите строку и нажмите «✏ Редактировать».")
    add_p("Сценарий 2. Управление проектами фонда. Вкладка «📁 Проекты фонда» отображает целевые программы сбора средств. Пользователь может отфильтровать проекты по статусу («Все», «Активен», «Завершен», «Приостановлен»). По каждому проекту автоматически рассчитывается текущий процент сбора средств и сумма, оставшаяся до достижения цели.")
    add_p("Сценарий 3. Регистрация входящего пожертвования. Откройте вкладку «💰 Пожертвования» и нажмите «➕ Внести взнос». В диалоговом окне выберите благотворителя из выпадающего списка, укажите целевой проект, сумму взноса, способ оплаты (карта, безналичный перевод, наличные) и примечание. После нажатия кнопки «Сохранить» пожертвование регистрируется в базе данных, а текущий баланс проекта атомарно увеличивается в транзакции. В случае необходимости отката пожертвования выделите запись и нажмите «🗑 Удалить (с откатом)» — баланс проекта будет автоматически скорректирован.")
    add_p("Сценарий 4. Учет получателей помощи. На вкладке «🤝 Получатели помощи» регистрируются заявки граждан и учреждений. Для каждой заявки фиксируется направление помощи, контактные данные, детальное описание жизненной ситуации и текущий статус рассмотрения.")
    add_p("Сценарий 5. Фиксация расходов фонда. Вкладка «📉 Расходы фонда» предназначена для списания средств на реализацию проектов и оказание адресной помощи. При нажатии «➕ Зарегистрировать расход» указывается проект списания, статья расхода (медикаменты, адресная помощь, оборудование), сумма, дата и обязательный номер первичного бухгалтерского документа (счет-фактура, акт, чек).")
    add_p("Сценарий 6. Формирование аналитических отчетов. Откройте вкладку «📊 Аналитические запросы и отчеты». В выпадающем списке выберите один из пяти регламентированных отчетов:")
    add_p("— 1. Сумма пожертвований (сводка по проектам, количеству взносов и датам с возможностью задания диапазона дат);")
    add_p("— 2. Активные проекты (цель, собрано, процент выполнения, остаток к сбору);")
    add_p("— 3. Крупнейшие доноры (рейтинг ключевых благотворителей по общей сумме пожертвований);")
    add_p("— 4. Расходы фонда (сопоставление собранных и израсходованных средств, текущий баланс по каждому проекту);")
    add_p("— 5. Помощь по категориям (распределение финансовых потоков и количества получателей по направлениям деятельности фонда).")
    add_p("После выбора отчета нажмите «⚡ Сформировать». Сформированный отчет можно мгновенно выгрузить в табличный файл Excel нажатием кнопки «📥 В Excel / CSV» либо отправить на печать через кнопку «🌐 Печать / HTML».")

    doc.add_page_break()

    # =========================================================================
    # ЗАКЛЮЧЕНИЕ
    # =========================================================================
    add_h1("ЗАКЛЮЧЕНИЕ")
    add_p("В ходе прохождения технологической (проектно-технологической) практики была успешно спроектирована и реализована информационная система «Благотворительный фонд», представляющая собой надежный программный комплекс на платформе C# / .NET 8 (Windows Forms) с базой данных Microsoft SQL Server.")
    add_p("В процессе выполнения индивидуального задания практики были достигнуты следующие результаты:")
    add_p("1. Проведен глубокий системный анализ предметной области деятельности благотворительных организаций, выделены ключевые информационные сущности и функциональные взаимосвязи между ними.")
    add_p("2. Разработана реляционная структура базы данных «CharityFundDB» в среде Microsoft SQL Server, включающая 6 нормализованных таблиц, систему первичных и внешних ключей, проверочные ограничения и индексы, а также 5 представлений для аналитических срезов.")
    add_p("3. Сформирован полный SQL-скрипт создания схемы данных и наполнения реалистичными демонстрационными данными (CharityFund_CreateDB.sql).")
    add_p("4. Реализован прозрачный и производительный слой доступа к данным на базе ADO.NET (Microsoft.Data.SqlClient) с параметризацией запросов и поддержкой механизма транзакций SqlTransaction для обеспечения целостности финансовых потоков.")
    add_p("5. Создан эргономичный графический интерфейс пользователя на Windows Forms с развитыми возможностями навигации, поиска, фильтрации, модального редактирования и экспорта данных в форматы CSV/Excel и HTML.")
    add_p("6. Реализованы все пять обязательных аналитических запросов, регламентированных заданием на практику.")
    add_p("7. Разработан набор модульных тестов (xUnit), подтверждающий корректность расчетов и целостность базы данных.")
    add_p("Разработанная система полностью готова к практическому использованию, демонстрации и защите.")

    doc.add_page_break()

    # =========================================================================
    # СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ
    # =========================================================================
    add_h1("СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ")
    lit_items = [
        "1. Microsoft. Документация по платформе .NET и языку C# [Электронный ресурс]. — Режим доступа: https://learn.microsoft.com/ru-ru/dotnet/csharp. — Дата доступа: 01.07.2026.",
        "2. Microsoft. Документация по Windows Forms [Электронный ресурс]. — Режим доступа: https://learn.microsoft.com/ru-ru/dotnet/desktop/winforms. — Дата доступа: 01.07.2026.",
        "3. Microsoft. Документация по библиотеке доступа к данным Microsoft.Data.SqlClient [Электронный ресурс]. — Режим доступа: https://learn.microsoft.com/ru-ru/sql/connect/ado-net. — Дата доступа: 01.07.2026.",
        "4. Кузнецов В. В. Базы данных. Проектирование и разработка реляционных баз данных в MS SQL Server: Учебное пособие. — М.: Инфра-М, 2023. — 288 с.",
        "5. Троелсен Э., Джепикс Ф. Язык C# 10 и платформа .NET 6: Основные принципы и практики программирования. 11-е изд. — СПб.: Диалектика, 2022. — 1456 с.",
        "6. Фаулер М. Шаблоны корпоративных приложений. — М.: Вильямс, 2021. — 544 с.",
        "7. Бен-Ган И. T-SQL Fundamentals: Основы T-SQL для разработчиков и администраторов SQL Server. — М.: БХВ-Петербург, 2022. — 464 с."
    ]
    for lit in lit_items:
        add_p(lit, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6, first_indent=1.25)

    doc.save(output_path)
    print(f"Отчет успешно сохранен в: {output_path}")

if __name__ == "__main__":
    out = os.path.abspath("Отчет_Афанасьев_Благотворительный_Фонд.docx")
    create_report(out)
