# -*- coding: utf-8 -*-
"""
Скрипт генерации презентации к защите практики в формате 16:9 PowerPoint (.pptx).
Тема: «Информационная система Благотворительный фонд»
Студент: Афанасьев М.А. (ГОУ ПГУ им. Т.Г. Шевченко, кафедра ПОВТ, 2026 г.)
Файл: Презентация_Благотворительный_Фонд.pptx
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    # Формат 16:9 (13.333 x 7.5 дюймов)
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Цветовая палитра
    NAVY_DARK = RGBColor(15, 23, 42)       # Slate 900
    NAVY_CARD = RGBColor(30, 41, 59)       # Slate 800
    BLUE_ACCENT = RGBColor(37, 99, 235)    # Blue 600
    BLUE_LIGHT = RGBColor(96, 165, 250)    # Blue 400
    EMERALD = RGBColor(16, 185, 129)       # Green 500
    AMBER = RGBColor(245, 158, 11)         # Amber 500
    WHITE = RGBColor(255, 255, 255)
    GRAY_TEXT = RGBColor(203, 213, 225)    # Slate 300
    GRAY_MUTED = RGBColor(148, 163, 184)   # Slate 400
    BG_LIGHT = RGBColor(248, 250, 252)     # Slate 50
    CARD_BORDER = RGBColor(51, 65, 85)     # Slate 700

    def add_bg(slide, dark=True):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
        bg.fill.solid()
        bg.fill.fore_color.rgb = NAVY_DARK if dark else BG_LIGHT
        bg.line.fill.background()
        return bg

    def add_header(slide, title_text, category_text="ИНФОРМАЦИОННАЯ СИСТЕМА «БЛАГОТВОРИТЕЛЬНЫЙ ФОНД»"):
        # Категория / верхний тег
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.5), Inches(0.4))
        p_cat = cat_box.text_frame.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.name = "Segoe UI"
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = BLUE_LIGHT

        # Основной заголовок слайда
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.7), Inches(11.5), Inches(0.8))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.text = title_text
        p_title.font.name = "Segoe UI"
        p_title.font.size = Pt(24)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE

    def add_card(slide, left, top, width, height, title="", body_items=None, dark_card=True, accent_color=BLUE_ACCENT):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = NAVY_CARD if dark_card else WHITE
        card.line.color.rgb = accent_color
        card.line.width = Pt(1.5)

        tf = card.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.25)
        tf.margin_right = Inches(0.25)
        tf.margin_top = Inches(0.2)
        tf.margin_bottom = Inches(0.2)

        if title:
            p0 = tf.paragraphs[0]
            p0.text = title
            p0.font.name = "Segoe UI"
            p0.font.size = Pt(14)
            p0.font.bold = True
            p0.font.color.rgb = accent_color
            p0.space_after = Pt(8)

        if body_items:
            for i, item in enumerate(body_items):
                p = tf.add_paragraph() if (title or i > 0) else tf.paragraphs[0]
                p.text = item
                p.font.name = "Segoe UI"
                p.font.size = Pt(11)
                p.font.color.rgb = GRAY_TEXT if dark_card else RGBColor(30, 41, 59)
                p.space_after = Pt(4)
        return card

    # ==========================================
    # СЛАЙД 1: ТИТУЛЬНЫЙ
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1, dark=True)

    # Верхний заголовок университета
    u_box = s1.shapes.add_textbox(Inches(1.0), Inches(0.7), Inches(11.3), Inches(0.8))
    pu = u_box.text_frame.paragraphs[0]
    pu.text = "ГОУ «ПРИДНЕСТРОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ ИМ. Т.Г. ШЕВЧЕНКО»\nФИЗИКО-ТЕХНИЧЕСКИЙ ИНСТИТУТ • КАФЕДРА ПОВТ"
    pu.font.name = "Segoe UI"
    pu.font.size = Pt(12)
    pu.font.bold = True
    pu.font.color.rgb = BLUE_LIGHT
    pu.alignment = PP_ALIGN.CENTER

    # Заголовок проекта
    t_box = s1.shapes.add_textbox(Inches(1.0), Inches(2.0), Inches(11.3), Inches(2.0))
    pt = t_box.text_frame.paragraphs[0]
    pt.text = "ИНФОРМАЦИОННАЯ СИСТЕМА\n«БЛАГОТВОРИТЕЛЬНЫЙ ФОНД»"
    pt.font.name = "Segoe UI"
    pt.font.size = Pt(36)
    pt.font.bold = True
    pt.font.color.rgb = WHITE
    pt.alignment = PP_ALIGN.CENTER

    psub = t_box.text_frame.add_paragraph()
    psub.text = "Отчет по технологической (проектно-технологической) практике"
    psub.font.name = "Segoe UI"
    psub.font.size = Pt(16)
    psub.font.color.rgb = AMBER
    psub.alignment = PP_ALIGN.CENTER
    psub.space_before = Pt(8)

    # Карточки студента и руководителей
    add_card(s1, Inches(1.5), Inches(4.5), Inches(4.8), Inches(2.0),
             title="ИСПОЛНИТЕЛЬ",
             body_items=[
                 "Студент 2 курса, группа ФТ24ДР62ПИ",
                 "Направление 09.03.04 «Программная инженерия»",
                 "Афанасьев М.А.",
                 "Очная форма обучения"
             ], accent_color=BLUE_ACCENT)

    add_card(s1, Inches(7.0), Inches(4.5), Inches(4.8), Inches(2.0),
             title="РУКОВОДИТЕЛИ ПРАКТИКИ",
             body_items=[
                 "От организации: зав. кафедрой ПОВТ Помян С.В.",
                 "От университета: ст. преп. Терещенко Е.В.",
                 "Место: УВЦ ФТИ ПГУ им. Т.Г. Шевченко",
                 "Тирасполь, 2026 год"
             ], accent_color=EMERALD)

    # ==========================================
    # СЛАЙД 2: АКТУАЛЬНОСТЬ И ЦЕЛЬ ПРОЕКТА
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2, dark=True)
    add_header(s2, "Актуальность и целеполагание разработки")

    add_card(s2, Inches(0.8), Inches(1.8), Inches(3.6), Inches(4.8),
             title="ПРОБЛЕМАТИКА",
             body_items=[
                 "• Ручной учет в разрозненных таблицах Excel и бумажных журналах.",
                 "• Риск потери данных о целевом назначении пожертвований.",
                 "• Отсутствие оперативного баланса между поступлениями и тратами фонда.",
                 "• Сложность формирования быстрой регламентированной аналитической отчетности."
             ], accent_color=AMBER)

    add_card(s2, Inches(4.8), Inches(1.8), Inches(3.6), Inches(4.8),
             title="ЦЕЛЬ РАЗРАБОТКИ",
             body_items=[
                 "• Создание высокопроизводительной настольной информационной системы на C# (.NET 8 Windows Forms).",
                 "• Обеспечение абсолютной прозрачности каждого поступившего и израсходованного рубля.",
                 "• Автоматизация учета доноров, проектов, пожертвований и получателей помощи.",
                 "• Мгновенное формирование 5 аналитических отчетов."
             ], accent_color=BLUE_ACCENT)

    add_card(s2, Inches(8.8), Inches(1.8), Inches(3.7), Inches(4.8),
             title="ПРАКТИЧЕСКАЯ ПОЛЬЗА",
             body_items=[
                 "• Исключение ошибок человеческого фактора благодаря валидации.",
                 "• Триггерная синхронизация собранных сумм в MS SQL Server.",
                 "• Легкая адаптация под любую благотворительную организацию.",
                 "• Подготовка к защите практики на кафедре ПОВТ."
             ], accent_color=EMERALD)

    # ==========================================
    # СЛАЙД 3: ПРЕДМЕТНАЯ ОБЛАСТЬ И СУЩНОСТИ
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3, dark=True)
    add_header(s3, "Информационная модель и сущности предметной области")

    entities = [
        ("Категории (Categories)", "Классификатор программ (детское здоровье, поддержка ветеранов, сироты, экология)."),
        ("Доноры (Donors)", "Физические, юридические лица и анонимные благотворители. Контакты, реквизиты, история."),
        ("Проекты (Projects)", "Целевые сборы фонда: название, бюджет (TargetAmount), собрано (CurrentAmount), даты, статус."),
        ("Пожертвования (Donations)", "Транзакции взносов: привязка донора к проекту, сумма (>0), дата, способ оплаты (карта, банк, наличные)."),
        ("Получатели (Recipients)", "Нуждающиеся граждане и учреждения: социальный статус, описание трудной ситуации, контакты."),
        ("Расходы (Expenses)", "Целевое списание средств по проектам: покупка лекарств, оборудования, оплата операций, ответственный.")
    ]

    for idx, (ent_title, ent_desc) in enumerate(entities):
        col = idx % 3
        row = idx // 3
        left = Inches(0.8 + col * 4.0)
        top = Inches(1.8 + row * 2.5)
        add_card(s3, left, top, Inches(3.7), Inches(2.2),
                 title=ent_title,
                 body_items=[ent_desc],
                 accent_color=BLUE_LIGHT if row == 0 else EMERALD)

    # ==========================================
    # СЛАЙД 4: СТРУКТУРА БАЗЫ ДАННЫХ (MS SQL SERVER)
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_bg(s4, dark=True)
    add_header(s4, "Схема базы данных CharityFundDB и целостность данных")

    add_card(s4, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8),
             title="АРХИТЕКТУРА БАЗЫ ДАННЫХ (3НФ)",
             body_items=[
                 "• Categories (Id, Name, Description) — PK, UNIQUE Name",
                 "• Donors (Id, FullName, DonorType, Phone, Email, Address, RegistrationDate)",
                 "• Projects (Id, Name, CategoryId, TargetAmount, CurrentAmount, Dates, Status) — FK к Categories",
                 "• Donations (Id, DonorId, ProjectId, Amount, DonationDate, PaymentMethod) — FK к Donors и Projects",
                 "• Recipients (Id, FullName, CategoryId, Phone, Address, NeedDescription, Status) — FK к Categories",
                 "• Expenses (Id, ProjectId, RecipientId, Amount, ExpenseDate, Purpose, ResponsiblePerson) — FK к Projects и Recipients"
             ], accent_color=BLUE_ACCENT)

    add_card(s4, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8),
             title="ОГРАНИЧЕНИЯ И СРЕДСТВА ЦЕЛОСТНОСТИ",
             body_items=[
                 "• Первичные ключи (PRIMARY KEY): суррогатные ключи INT IDENTITY(1,1).",
                 "• Внешние ключи (FOREIGN KEY): целостность связей 1:M с правилами каскадного обновления.",
                 "• Ограничения CHECK: Amount > 0, TargetAmount > 0, EndDate >= StartDate, списки допустимых статусов.",
                 "• SQL-триггер TR_Donations_UpdateCurrentAmount: автоматический пересчет Projects.CurrentAmount при любых INSERT/UPDATE/DELETE.",
                 "• 5 аналитических представлений (VIEW): инкапсуляция сложных многотабличных JOIN-запросов."
             ], accent_color=AMBER)

    # ==========================================
    # СЛАЙД 5: АРХИТЕКТУРА ПРИЛОЖЕНИЯ И ADO.NET
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_bg(s5, dark=True)
    add_header(s5, "Архитектура программного комплекса и доступ к данным")

    add_card(s5, Inches(0.8), Inches(1.8), Inches(3.6), Inches(4.8),
             title="СЛОЙ ПРЕДСТАВЛЕНИЯ (UI)",
             body_items=[
                 "• Платформа: .NET 8 (Windows Forms).",
                 "• Главная форма MainForm с удобной навигацией по 7 разделам.",
                 "• Модальные формы добавления и редактирования записей.",
                 "• Мгновенный контекстный поиск и фильтрация данных в DataGridView.",
                 "• Окно настроек подключения к SQL Server."
             ], accent_color=BLUE_ACCENT)

    add_card(s5, Inches(4.8), Inches(1.8), Inches(3.6), Inches(4.8),
             title="ДОСТУП К ДАННЫМ (ADO.NET)",
             body_items=[
                 "• Библиотека Microsoft.Data.SqlClient.",
                 "• Прозрачный класс DatabaseHelper: методы ExecuteQuery, ExecuteNonQuery, ExecuteScalar.",
                 "• Строго параметризованные запросы (SqlParameter): 100% защита от SQL Injection.",
                 "• Отсутствие скрытой магии ORM: абсолютная ясность кода для защиты на кафедре.",
                 "• Репозитории по каждой предметной области."
             ], accent_color=EMERALD)

    add_card(s5, Inches(8.8), Inches(1.8), Inches(3.7), Inches(4.8),
             title="СЕРВИСНЫЙ УРОВЕНЬ",
             body_items=[
                 "• Класс ExportService: потоковый экспорт таблиц в формат CSV.",
                 "• Кодировка UTF-8 с маркером BOM для автоматического открытия в Microsoft Excel.",
                 "• Автоматическая валидация обязательных полей при вводе.",
                 "• Индикация сетевого статуса подключения к БД."
             ], accent_color=AMBER)

    # ==========================================
    # СЛАЙД 6: ПОЛЬЗОВАТЕЛЬСКИЙ ИНТЕРФЕЙС
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_bg(s6, dark=True)
    add_header(s6, "Эргономика и возможности пользовательского интерфейса")

    add_card(s6, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8),
             title="СТРУКТУРА ГЛАВНОГО ОКНА (MAINFORM)",
             body_items=[
                 "1. Верхняя панель (Top Header): наименование системы, сведения об авторе, кнопки «🔄 Обновить всё» и «⚙️ Настройки БД».",
                 "2. Вкладочная навигация (7 разделов): Доноры, Проекты, Категории, Пожертвования, Получатели, Расходы, Аналитика.",
                 "3. Интерактивные таблицы DataGridView с чередующимися цветами строк для удобства чтения оператором.",
                 "4. Панели управления: контекстный поиск, фильтры по статусам, фильтры по датам (DateTimePicker).",
                 "5. Нижняя статусная строка: состояние БД, имя сервера и подсказки."
             ], accent_color=BLUE_ACCENT)

    add_card(s6, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8),
             title="УДОБСТВО ОПЕРАТОРА И ЗАЩИТА ОТ ОШИБОК",
             body_items=[
                 "• Выпадающие списки (ComboBox) с автоподстановкой внешних ключей (нет необходимости вручную вводить ID).",
                 "• Блокировка отрицательных и нулевых сумм финансовой помощи.",
                 "• Диалоги подтверждения перед удалением критически важных записей.",
                 "• Подсчет и отображение суммарных показателей (всего собрано, всего потрачено, баланс фонда) в реальном времени.",
                 "• Возможность развернуть БД прямо из приложения кнопкой «Инициализировать БД»."
             ], accent_color=EMERALD)

    # ==========================================
    # СЛАЙД 7: 5 АНАЛИТИЧЕСКИХ ОТЧЕТОВ
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_bg(s7, dark=True)
    add_header(s7, "Реализация 5 обязательных аналитических запросов")

    reports = [
        ("1. Сумма пожертвований", "Сводная ведомость сборов по проектам: общая сумма, количество взносов, даты первого/последнего пожертвования, фильтр по диапазону дат."),
        ("2. Активные проекты", "Мониторинг текущих программ со статусом «Активен»: целевой бюджет, фактически собрано, процент выполнения цели и сумма к добору."),
        ("3. Крупнейшие доноры", "Рейтинг ТОП-благотворителей (физ. и юр. лица) по совокупной сумме взносов с отображением даты последнего перевода."),
        ("4. Расходы фонда", "Финансовый баланс программ: собрано, израсходовано, текущий свободный остаток средств и детальный журнал списаний."),
        ("5. Помощь по категориям", "Анализ эффективности работы фонда по направлениям: охват проектов, количество получателей помощи, объем привлеченных и потраченных средств.")
    ]

    for idx, (rep_title, rep_desc) in enumerate(reports):
        top = Inches(1.7 + idx * 1.05)
        add_card(s7, Inches(0.8), top, Inches(11.7), Inches(0.95),
                 title=rep_title,
                 body_items=[rep_desc],
                 accent_color=AMBER if idx % 2 == 0 else BLUE_LIGHT)

    # ==========================================
    # СЛАЙД 8: ФИНАНСОВЫЙ БАЛАНС И КОНТРОЛЬ СРЕДСТВ
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_bg(s8, dark=True)
    add_header(s8, "Финансовый баланс фонда и целевое расходование")

    add_card(s8, Inches(0.8), Inches(1.8), Inches(3.6), Inches(4.8),
             title="ПОСТУПЛЕНИЯ (ДОХОДЫ)",
             body_items=[
                 "• Аккумуляция взносов от физических лиц, компаний-меценатов и анонимных дарителей.",
                 "• Поддержка всех популярных способов оплаты: банковские переводы, карты, наличные.",
                 "• Строгая привязка каждого платежа к конкретной благотворительной программе.",
                 "• Автоматический рост прогресса сбора проекта через SQL-триггер."
             ], accent_color=EMERALD)

    add_card(s8, Inches(4.8), Inches(1.8), Inches(3.6), Inches(4.8),
             title="ЦЕЛЕВЫЕ РАСХОДЫ",
             body_items=[
                 "• Прозрачное списание средств строго в рамках утвержденного бюджета проекта.",
                 "• Персонализированная привязка расходов к конкретным благополучателям.",
                 "• Фиксация ответственного сотрудника фонда за каждую расходную операцию.",
                 "• Исключение нецелевого расходования пожертвованных средств."
             ], accent_color=AMBER)

    add_card(s8, Inches(8.8), Inches(1.8), Inches(3.7), Inches(4.8),
             title="ТЕКУЩИЙ ОСТАТОК (БАЛАНС)",
             body_items=[
                 "• Мгновенный расчет формулы:\n  Баланс = Собрано - Израсходовано.",
                 "• Предотвращение кассовых разрывов и перерасхода бюджетов программ.",
                 "• Экспорт сводного финансового отчета в CSV для аудиторов и попечителей.",
                 "• Полное соответствие требованиям некоммерческой отчетности."
             ], accent_color=BLUE_ACCENT)

    # ==========================================
    # СЛАЙД 9: РУКОВОДСТВО ПО ОПЕРАЦИЯМ И ЗАПУСКУ
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_bg(s9, dark=True)
    add_header(s9, "Порядок запуска, развертывания и демонстрации")

    steps = [
        ("Шаг 1: Развертывание базы данных", "Запуск скрипта CharityFund_CreateDB.sql в SSMS либо нажатие кнопки «Инициализировать БД из скрипта» в окне настроек приложения. Создаются все таблицы, триггер, представления и демо-данные."),
        ("Шаг 2: Настройка строки подключения", "В файле App.config задана строка подключения к SQL Server (Server=localhost или (localdb)\\mssqllocaldb;Database=CharityFundDB;Trusted_Connection=True;). Возможно горячее изменение в окне настроек."),
        ("Шаг 3: Запуск и демонстрация операций", "Сборка и запуск CharityFundApp.exe через Visual Studio 2022 или команду dotnet run. Демонстрация добавления донора, регистрации взноса, обновления прогресса сбора и списания расхода."),
        ("Шаг 4: Демонстрация 5 аналитических отчетов", "Переход на вкладку «📊 Аналитика и Отчеты», выбор каждого из 5 отчетов, показ динамического расчета сумм, процентов и остатков, демонстрация экспорта отчета в CSV.")
    ]

    for idx, (st_title, st_desc) in enumerate(steps):
        top = Inches(1.8 + idx * 1.25)
        add_card(s9, Inches(0.8), top, Inches(11.7), Inches(1.1),
                 title=st_title,
                 body_items=[st_desc],
                 accent_color=BLUE_ACCENT if idx % 2 == 0 else EMERALD)

    # ==========================================
    # СЛАЙД 10: ЗАКЛЮЧЕНИЕ И ИТОГИ
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_bg(s10, dark=True)
    add_header(s10, "Результаты выполнения технологической практики")

    add_card(s10, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8),
             title="ДОСТИГНУТЫЕ РЕЗУЛЬТАТЫ",
             body_items=[
                 "✔ Разработана и развернута полнофункциональная реляционная база данных CharityFundDB в MS SQL Server.",
                 "✔ Создано настольное приложение Windows Forms на платформе .NET 8 (C# 12).",
                 "✔ Реализован прозрачный и безопасный слой доступа к данным на классическом ADO.NET (параметризованные запросы).",
                 "✔ Реализованы все операции добавления, редактирования и удаления (CRUD) для всех 6 таблиц.",
                 "✔ В полном объеме реализованы 5 обязательных аналитических выборок с функцией экспорта в CSV.",
                 "✔ Подготовлен развернутый академический отчет по практике и комплект документации."
             ], accent_color=EMERALD)

    add_card(s10, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8),
             title="ОЦЕНКА И ГОТОВНОСТЬ К ЗАЩИТЕ",
             body_items=[
                 "• Код проекта полностью прозрачен, структурирован и свободен от сторонней магии сложных ORM.",
                 "• Любая строчка кода легко комментируется и объясняется перед экзаменационной комиссией кафедры ПОВТ.",
                 "• Проект собирается с 0 ошибок и 0 предупреждений через dotnet build.",
                 "• Решение готово к показу на локальном компьютере или сервере кафедры.",
                 "• Студент Афанасьев М.А. готов к защите практики на «Отлично»!"
             ], accent_color=AMBER)

    # ==========================================
    # СЛАЙД 11: СПАСИБО ЗА ВНИМАНИЕ!
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    add_bg(s11, dark=True)

    c_box = s11.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.3), Inches(3.0))
    pc = c_box.text_frame.paragraphs[0]
    pc.text = "СПАСИБО ЗА ВНИМАНИЕ!"
    pc.font.name = "Segoe UI"
    pc.font.size = Pt(44)
    pc.font.bold = True
    pc.font.color.rgb = WHITE
    pc.alignment = PP_ALIGN.CENTER

    pc2 = c_box.text_frame.add_paragraph()
    pc2.text = "Готов ответить на вопросы комиссии"
    pc2.font.name = "Segoe UI"
    pc2.font.size = Pt(20)
    pc2.font.color.rgb = BLUE_LIGHT
    pc2.alignment = PP_ALIGN.CENTER
    pc2.space_before = Pt(14)

    pc3 = c_box.text_frame.add_paragraph()
    pc3.text = "Разработчик: Афанасьев М.А., студент гр. ФТ24ДР62ПИ\nГОУ «ПГУ им. Т.Г. Шевченко», Тирасполь, 2026 г."
    pc3.font.name = "Segoe UI"
    pc3.font.size = Pt(14)
    pc3.font.color.rgb = GRAY_MUTED
    pc3.alignment = PP_ALIGN.CENTER
    pc3.space_before = Pt(20)

    pptx_path = os.path.abspath("Презентация_Благотворительный_Фонд.pptx")
    prs.save(pptx_path)
    print(f"Презентация успешно сохранена: {pptx_path}")

if __name__ == '__main__':
    create_presentation()
