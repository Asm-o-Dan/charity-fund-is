import os
import pptx
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_presentation(output_path):
    prs = Presentation()
    # 16:9 Widescreen: 13.333 x 7.5 inches
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Цветовая палитра
    C_BG_DARK = RGBColor(15, 23, 42)      # Slate 900
    C_BG_LIGHT = RGBColor(248, 250, 252)  # Slate 50
    C_PRIMARY = RGBColor(37, 99, 235)     # Royal Blue
    C_PRIMARY_DARK = RGBColor(30, 58, 138)
    C_TEXT_MAIN = RGBColor(30, 41, 59)    # Slate 800
    C_TEXT_MUTED = RGBColor(100, 116, 139) # Slate 500
    C_ACCENT_GREEN = RGBColor(16, 185, 129)
    C_ACCENT_ORANGE = RGBColor(245, 158, 11)
    C_CARD_BG = RGBColor(255, 255, 255)
    C_CARD_BORDER = RGBColor(226, 232, 240)

    def add_header(slide, title, category="ОТЧЕТ ПО ТЕХНОЛОГИЧЕСКОЙ ПРАКТИКЕ"):
        # Header background banner
        header_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
        header_shape.fill.solid()
        header_shape.fill.fore_color.rgb = C_BG_DARK
        header_shape.line.fill.background()

        # Category text
        cat_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.12), Inches(11.5), Inches(0.3))
        tf_cat = cat_box.text_frame
        tf_cat.word_wrap = True
        p_cat = tf_cat.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.name = "Segoe UI"
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = C_ACCENT_GREEN

        # Title text
        title_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.38), Inches(11.5), Inches(0.65))
        tf_title = title_box.text_frame
        tf_title.word_wrap = True
        p_title = tf_title.paragraphs[0]
        p_title.text = title
        p_title.font.name = "Segoe UI"
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = RGBColor(255, 255, 255)

    def add_card(slide, left, top, width, height, title, items, border_color=C_PRIMARY):
        # Card shape
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(left), Inches(top), Inches(width), Inches(height))
        card.fill.solid()
        card.fill.fore_color.rgb = C_CARD_BG
        card.line.color.rgb = border_color
        card.line.width = Pt(1.5)

        # Title
        t_box = slide.shapes.add_textbox(Inches(left + 0.2), Inches(top + 0.15), Inches(width - 0.4), Inches(0.45))
        tf = t_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = title
        p.font.name = "Segoe UI"
        p.font.size = Pt(15)
        p.font.bold = True
        p.font.color.rgb = C_PRIMARY_DARK

        # Items
        c_box = slide.shapes.add_textbox(Inches(left + 0.2), Inches(top + 0.65), Inches(width - 0.4), Inches(height - 0.75))
        ctf = c_box.text_frame
        ctf.word_wrap = True
        for i, item in enumerate(items):
            p = ctf.add_paragraph() if i > 0 else ctf.paragraphs[0]
            p.text = f"• {item}" if not item.startswith("✔") and not item.startswith("➡") else item
            p.font.name = "Segoe UI"
            p.font.size = Pt(12)
            p.font.color.rgb = C_TEXT_MAIN
            p.space_after = Pt(4)

    # =========================================================================
    # СЛАЙД 1: ТИТУЛЬНЫЙ СЛАЙД
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = C_BG_DARK
    bg1.line.fill.background()

    # Decorative accent bar
    bar1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.0), Inches(1.2), Inches(0.15), Inches(4.8))
    bar1.fill.solid()
    bar1.fill.fore_color.rgb = C_PRIMARY
    bar1.line.fill.background()

    t_box1 = s1.shapes.add_textbox(Inches(1.4), Inches(1.1), Inches(11.0), Inches(3.2))
    tf1 = t_box1.text_frame
    tf1.word_wrap = True

    p_org = tf1.paragraphs[0]
    p_org.text = "ГОУ «ПРИДНЕСТРОВСКИЙ ГОСУДАРСТВЕННЫЙ УНИВЕРСИТЕТ им. Т.Г. ШЕВЧЕНКО»\nФИЗИКО-ТЕХНИЧЕСКИЙ ИНСТИТУТ • КАФЕДРА ПОВТ"
    p_org.font.name = "Segoe UI"
    p_org.font.size = Pt(13)
    p_org.font.color.rgb = C_ACCENT_GREEN
    p_org.font.bold = True
    p_org.space_after = Pt(16)

    p_main = tf1.add_paragraph()
    p_main.text = "ИНФОРМАЦИОННАЯ СИСТЕМА\n«БЛАГОТВОРИТЕЛЬНЫЙ ФОНД»"
    p_main.font.name = "Segoe UI"
    p_main.font.size = Pt(32)
    p_main.font.bold = True
    p_main.font.color.rgb = RGBColor(255, 255, 255)
    p_main.space_after = Pt(12)

    p_sub = tf1.add_paragraph()
    p_sub.text = "Отчет по технологической (проектно-технологической) практике"
    p_sub.font.name = "Segoe UI"
    p_sub.font.size = Pt(16)
    p_sub.font.color.rgb = RGBColor(203, 213, 225)

    info_box = s1.shapes.add_textbox(Inches(1.4), Inches(4.6), Inches(11.0), Inches(2.2))
    itf = info_box.text_frame
    itf.word_wrap = True

    p_st = itf.paragraphs[0]
    p_st.text = "Выполнил: Афанасьев М.А., студент 2 курса группы ФТ24ДР62ПИ\nНаправление: 09.03.04 «Программная инженерия» (профиль «Разработка ПИС»)"
    p_st.font.name = "Segoe UI"
    p_st.font.size = Pt(13)
    p_st.font.color.rgb = RGBColor(255, 255, 255)
    p_st.space_after = Pt(6)

    p_lead = itf.add_paragraph()
    p_lead.text = "Руководитель от профильной организации: зав. кафедрой ПОВТ Помян С.В.\nРуководитель от университета: ст. преподаватель Е.В. Терещенко"
    p_lead.font.name = "Segoe UI"
    p_lead.font.size = Pt(12)
    p_lead.font.color.rgb = RGBColor(148, 163, 184)
    p_lead.space_after = Pt(8)

    p_loc = itf.add_paragraph()
    p_loc.text = "Тирасполь, 2026 г."
    p_loc.font.name = "Segoe UI"
    p_loc.font.size = Pt(12)
    p_loc.font.bold = True
    p_loc.font.color.rgb = C_ACCENT_ORANGE

    # =========================================================================
    # СЛАЙД 2: АКТУАЛЬНОСТЬ И ЦЕЛЬ ПРОЕКТА
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Актуальность и цель разработки")
    add_card(s2, 0.8, 1.5, 5.6, 5.4, "Актуальность темы", [
        "Необходимость строгой финансовой прозрачности некоммерческих благотворительных фондов.",
        "Требование персонального учета каждого поступившего взноса и целевого сопоставления с проектами.",
        "Исключение нецелевого расходования средств: каждый акт расхода привязывается к первичному документу (чеку, акту, накладной).",
        "Преодоление разрозненности Excel-файлов за счет централизованной реляционной БД MS SQL Server.",
        "Предоставление оперативной аналитики для попечительского совета и публичной отчетности."
    ], C_PRIMARY)

    add_card(s2, 6.8, 1.5, 5.7, 5.4, "Цель и ключевые задачи", [
        "Разработка настольного программного комплекса на платформе .NET 8 (Windows Forms / C#).",
        "Проектирование нормализованной реляционной базы данных «CharityFundDB» в MS SQL Server.",
        "Реализация прямого, быстрого и защищенного доступа к данным через чистый ADO.NET (Microsoft.Data.SqlClient).",
        "Обеспечение транзакционной целостности при проведении пожертвований и списаний.",
        "Построение 5 регламентированных аналитических отчетов согласно индивидуальному заданию практики."
    ], C_ACCENT_GREEN)

    # =========================================================================
    # СЛАЙД 3: СУЩНОСТИ ПРЕДМЕТНОЙ ОБЛАСТИ
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Сущности предметной области (ER-модель)")
    add_card(s3, 0.8, 1.5, 3.6, 2.6, "1. Доноры (Donors)", [
        "Физические и юридические лица.",
        "ФИО / Наименование организации.",
        "Телефон, Email, фактический адрес.",
        "Дата регистрации в фонде."
    ])
    add_card(s3, 4.8, 1.5, 3.6, 2.6, "2. Проекты (Projects)", [
        "Целевые программы фонда.",
        "Целевая и собранная суммы.",
        "Сроки проведения сбора средств.",
        "Статус: Активен, Завершен, Приостановлен."
    ])
    add_card(s3, 8.8, 1.5, 3.7, 2.6, "3. Категории (Categories)", [
        "Направления благотворительности.",
        "Детское здоровье, медицина, малоимущие, ветераны, животные.",
        "Контроль целостности при удалении."
    ])

    add_card(s3, 0.8, 4.3, 3.6, 2.7, "4. Пожертвования (Donations)", [
        "Связь донора и целевого проекта.",
        "Сумма, дата и время платежа.",
        "Способ оплаты (карта, счет, нал).",
        "Транзакционное зачисление средств."
    ], C_ACCENT_ORANGE)
    add_card(s3, 4.8, 4.3, 3.6, 2.7, "5. Получатели (Recipients)", [
        "Граждане и социальные учреждения.",
        "Детальное описание жизненной ситуации.",
        "Категория благотворительности.",
        "Статус: На рассмотрении, Одобрен, Помощь оказана."
    ], C_ACCENT_ORANGE)
    add_card(s3, 8.8, 4.3, 3.7, 2.7, "6. Расходы фонда (Expenses)", [
        "Списание средств по проекту.",
        "Опциональная привязка к получателю.",
        "Статья расходов (медикаменты, помощь).",
        "Обязательный номер первичного документа."
    ], C_ACCENT_ORANGE)

    # =========================================================================
    # СЛАЙД 4: СТРУКТУРА БАЗЫ ДАННЫХ И ЦЕЛОСТНОСТЬ
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Проектирование базы данных в MS SQL Server")
    add_card(s4, 0.8, 1.5, 5.6, 5.4, "Схема и ограничения целостности", [
        "СУБД: Microsoft SQL Server 2022 / SQLEXPRESS / LocalDB.",
        "Первичные ключи: суррогатные идентификаторы INT IDENTITY(1,1).",
        "Внешние ключи (FOREIGN KEY):",
        "  • FK_Projects_Categories (ON DELETE RESTRICT)",
        "  • FK_Recipients_Categories (ON DELETE RESTRICT)",
        "  • FK_Donations_Donors / Projects (ON DELETE CASCADE)",
        "  • FK_Expenses_Projects (ON DELETE CASCADE)",
        "  • FK_Expenses_Recipients (ON DELETE SET NULL)",
        "Ограничения CHECK:",
        "  • TargetAmount > 0, CurrentAmount >= 0",
        "  • Donation.Amount > 0, Expense.Amount > 0",
        "Некластеризованные индексы на внешние ключи и даты для ускорения отчетов."
    ], C_PRIMARY)

    add_card(s4, 6.8, 1.5, 5.7, 5.4, "Представления (Views) для аналитики", [
        "vw_DonationsSummary — агрегация сумм пожертвований по проектам, подсчет взносов и датировок.",
        "vw_ActiveProjects — расчет выполнения целевого сбора в процентах и остатка средств к сбору.",
        "vw_TopDonors — группировка доноров по суммарному объему помощи с определением максимального взноса.",
        "vw_FundExpenses — балансовый срез: сравнение собранных и списанных сумм с расчетом остатка бюджета.",
        "vw_CategoryAssistance — кросс-аналитика количества проектов, подопечных и финансовых потоков по направлениям.",
        "Полный скрипт развертывания: CharityFund_CreateDB.sql."
    ], C_ACCENT_GREEN)

    # =========================================================================
    # СЛАЙД 5: АРХИТЕКТУРА И ТЕХНОЛОГИЧЕСКИЙ СТЕК
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Архитектура приложения и стек технологий")
    add_card(s5, 0.8, 1.5, 5.6, 5.4, "Стек технологий и компоненты", [
        "Платформа: Microsoft .NET 8 (C# 12).",
        "Графический стек: Windows Forms (.NET Desktop).",
        "СУБД: Microsoft SQL Server (T-SQL).",
        "Драйвер данных: Microsoft.Data.SqlClient (ADO.NET).",
        "Конфигурация: App.config + ConfigurationManager.",
        "Тестирование: xUnit Test Framework + .NET Test SDK.",
        "Экспорт: CSV (RFC 4180 с UTF-8 BOM) + HTML-генерация.",
        "Сборка: чистый dotnet build и поддержка Visual Studio 2022."
    ], C_PRIMARY)

    add_card(s5, 6.8, 1.5, 5.7, 5.4, "Многоуровневая структура проекта", [
        "1. Models: POCO-классы домена (Donor, Project, Donation...) и DTO для отчетов с вычисляемыми свойствами.",
        "2. DataAccess: статический фасад DatabaseHelper и 7 репозиториев с параметризованными запросами.",
        "3. Forms: главное окно MainForm (вкладки, фильтрация, события) и специализированные модальные формы.",
        "4. Services: сервис ExportService для независимой выгрузки табличных представлений.",
        "Преимущество ADO.NET на защите: чистый и прозрачный код, каждый SQL-запрос под полным контролем студента, отсутствие «магии» тяжелых ORM."
    ], C_ACCENT_GREEN)

    # =========================================================================
    # СЛАЙД 6: ФИНАНСОВЫЕ ТРАНЗАКЦИИ (SqlTransaction)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Транзакционная надежность (ADO.NET SqlTransaction)")
    add_card(s6, 0.8, 1.5, 5.6, 5.4, "Проблема рассинхронизации", [
        "В финансовой системе недопустимо раздельное выполнение записи платежа и обновления баланса проекта.",
        "Если при сбое сервера пожертвование добавилось, а баланс проекта не обновился — возникает финансовая коллизия.",
        "Решение: применение SqlTransaction с уровнем изоляции ReadCommitted.",
        "Принцип неделимости (ACID): либо обе операции завершаются успешно (Commit), либо при любой ошибке происходит полный откат (Rollback)."
    ], C_PRIMARY)

    add_card(s6, 6.8, 1.5, 5.7, 5.4, "Реализация в DonationRepository", [
        "using var conn = DatabaseHelper.CreateConnection();\nconn.Open();\nusing var tx = conn.BeginTransaction();\ntry {",
        "  1. INSERT INTO Donations (...) VALUES (...);\n     cmd1.Transaction = tx;\n     int newId = cmd1.ExecuteScalar();",
        "  2. UPDATE Projects\n     SET CurrentAmount = CurrentAmount + @Amount\n     WHERE Id = @ProjectId;\n     cmd2.Transaction = tx;\n     cmd2.ExecuteNonQuery();",
        "  tx.Commit();\n} catch { tx.Rollback(); throw; }",
        "Аналогичный транзакционный механизм реализован при откате пожертвования (DeleteWithTransaction)."
    ], C_ACCENT_GREEN)

    # =========================================================================
    # СЛАЙД 7: ПОЛЬЗОВАТЕЛЬСКИЙ ИНТЕРФЕЙС
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Пользовательский интерфейс системы")
    add_card(s7, 0.8, 1.5, 5.6, 5.4, "Эргономика и дизайн Windows Forms", [
        "Стильное лаконичное оформление: цветовая гамма Slate & Royal Blue, отсутствие перегруженности.",
        "Шапка приложения с информацией о студенте и быстрым доступом к настройкам БД.",
        "Таблицы DataGridView с чередующимися строками, подсветкой курсора и автоформатированием денежных сумм (120 000,00 руб.).",
        "Мгновенный поиск по подстроке во всех ключевых реестрах.",
        "Фильтрация по статусам (Активен/Завершен) и временным периодам.",
        "Нижняя строка StatusStrip с онлайн-индикатором соединения."
    ], C_PRIMARY)

    add_card(s7, 6.8, 1.5, 5.7, 5.4, "Удобство настройки и демонстрации", [
        "Окно «Параметры БД» (ConnectionSettingsForm):",
        "  • 4 готовых пресета: localhost, .\\SQLEXPRESS, LocalDB, Custom.",
        "  • Кнопка «Проверить подключение» с моментальной диагностикой.",
        "  • Кнопка «⚡ Развернуть БД из скрипта» — выполняет скрипт CharityFund_CreateDB.sql прямо из приложения!",
        "Модальные диалоговые окна с проверкой обязательных полей, валидацией сумм (> 0) и выбором из связанных справочников.",
        "Безопасное каскадное удаление с предупреждающими диалогами."
    ], C_ACCENT_GREEN)

    # =========================================================================
    # СЛАЙД 8: 5 АНАЛИТИЧЕСКИХ ОТЧЕТОВ
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "5 обязательных аналитических запросов и отчетов")
    add_card(s8, 0.8, 1.5, 5.6, 5.4, "Реализация аналитических выборок", [
        "1. Сумма пожертвований:",
        "   Группировка по проектам, количество взносов, общая сумма, даты первого и последнего взноса + фильтр по диапазону дат.",
        "2. Активные проекты:",
        "   Проекты в статусе «Активен», расчет процента сбора (Current/Target * 100) и оставшейся суммы к сбору.",
        "3. Крупнейшие доноры:",
        "   Рейтинг TOP благотворителей по общей сумме помощи с фиксацией максимального разового взноса.",
        "4. Расходы фонда:",
        "   Сравнение собранных средств и фактических списаний по каждому проекту, расчет текущего остатка баланса.",
        "5. Помощь по категориям:",
        "   Распределение проектов, получателей, сумм пожертвований и выплат по направлениям помощи."
    ], C_PRIMARY)

    add_card(s8, 6.8, 1.5, 5.7, 5.4, "Экспорт и документирование", [
        "Интерактивное переключение отчетов через выпадающий список.",
        "Автоматический подсчет итоговой сводки в панели управления (Всего собрано, Количество операций, Текущий баланс).",
        "Экспорт в CSV / Excel:",
        "  • Поддержка кодировки UTF-8 BOM для автоматического открытия кириллицы в Microsoft Excel без кракозябр.",
        "Экспорт в HTML для печати:",
        "  • Генерация адаптивного адаптивного отчета с фирменной стилизацией и метаданными формирования документа."
    ], C_ACCENT_GREEN)

    # =========================================================================
    # СЛАЙД 9: МОДУЛЬНОЕ ТЕСТИРОВАНИЕ (xUnit)
    # =========================================================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Контроль качества и модульное тестирование")
    add_card(s9, 0.8, 1.5, 5.6, 5.4, "Набор тестов (CharityFundApp.Tests)", [
        "Использован современный тестовый фреймворк xUnit.",
        "Группа 1: Валидация вычислений доменных моделей:",
        "  • Корректность расчета Project.CompletionPercentage.",
        "  • Обработка нулевой цели сбора (деление на ноль исключено).",
        "  • Корректность вычисления остатка Project.RemainingAmount.",
        "  • Обработка ситуации сбора сверх цели (Overfunding).",
        "Группа 2: Расчет аналитических показателей:",
        "  • Проверка формулы баланса расходов (Собрано - Израсходовано).",
        "  • Сортировка и агрегация рейтинга благотворителей.",
        "Группа 3: Целостность скрипта базы данных:",
        "  • Проверка наличия всех 6 таблиц и 5 представлений."
    ], C_PRIMARY)

    add_card(s9, 6.8, 1.5, 5.7, 5.4, "Результаты прогона тестов", [
        "Команда выполнения: dotnet test CharityFundApp.sln",
        "✔ Пройдено тестов: 12 из 12 (100%).",
        "✔ Ошибок: 0.",
        "✔ Время выполнения: 32 мс.",
        "✔ Сборка: 0 предупреждений, 0 ошибок компилятора.",
        "Подтверждена корректность алгоритмической логики и надежность программного кода перед защитой."
    ], C_ACCENT_GREEN)

    # =========================================================================
    # СЛАЙД 10: ЗАКЛЮЧЕНИЕ
    # =========================================================================
    s10 = prs.slides.add_slide(blank_layout)
    bg10 = s10.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg10.fill.solid()
    bg10.fill.fore_color.rgb = C_BG_DARK
    bg10.line.fill.background()

    add_card(s10, 1.0, 1.2, 11.333, 4.2, "Итоги выполнения индивидуального задания практики", [
        "✔ Полностью спроектирована и реализована ИС «Благотворительный фонд» на C# / .NET 8 (Windows Forms).",
        "✔ Развернута база данных «CharityFundDB» в MS SQL Server с 6 таблицами, внешними ключами и 5 VIEW.",
        "✔ Реализован чистый и прозрачный доступ через ADO.NET (параметризованные запросы, SqlTransaction).",
        "✔ Реализованы все операции добавления, изменения, удаления и просмотра для каждой сущности.",
        "✔ Реализованы 5 аналитических отчетов с интерактивной фильтрацией и экспортом в CSV / HTML.",
        "✔ Подготовлен структурированный отчет по практике (.docx) в строгом соответствии с требованиями кафедры ПОВТ.",
        "✔ Разработан комплект слайдов (.pptx) для уверенной и наглядной защиты проекта."
    ], C_ACCENT_GREEN)

    t_end = s10.shapes.add_textbox(Inches(1.0), Inches(5.7), Inches(11.333), Inches(1.3))
    te_tf = t_end.text_frame
    p_thanks = te_tf.paragraphs[0]
    p_thanks.text = "СПАСИБО ЗА ВНИМАНИЕ!"
    p_thanks.font.name = "Segoe UI"
    p_thanks.font.size = Pt(26)
    p_thanks.font.bold = True
    p_thanks.font.color.rgb = RGBColor(255, 255, 255)
    p_thanks.alignment = PP_ALIGN.CENTER

    p_qa = te_tf.add_paragraph()
    p_qa.text = "Готов ответить на ваши вопросы. Афанасьев М.А., гр. ФТ24ДР62ПИ, ПГУ им. Т.Г. Шевченко"
    p_qa.font.name = "Segoe UI"
    p_qa.font.size = Pt(13)
    p_qa.font.color.rgb = C_ACCENT_ORANGE
    p_qa.alignment = PP_ALIGN.CENTER

    prs.save(output_path)
    print(f"Презентация успешно сохранена в: {output_path}")

if __name__ == "__main__":
    out = os.path.abspath("Презентация_Афанасьев_Благотворительный_Фонд.pptx")
    create_presentation(out)
