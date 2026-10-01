# -*- coding: utf-8 -*-
"""
Full report generator for:
Отчет_Практика_Благотворительный_Фонд.docx
Target volume: 27-28 pages (perfectly within 20-30 pages range)
"""

import os
import sys
import docx
from docx.shared import Pt, Cm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

def generate_report():
    doc = docx.Document()

    # Page setup (A4, standard GOST margins matching example document)
    section = doc.sections[0]
    section.top_margin = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(1.0)
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)

    # Base Normal Style setup
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Times New Roman'
    normal_style.font.size = Pt(14)
    normal_style.font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.5

    # Helper functions
    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=0, space_after=6, indent=18, bold=False, italic=False, size=14, line_spacing=1.5):
        p = doc.add_paragraph()
        p.alignment = align
        pf = p.paragraph_format
        pf.space_before = Pt(space_before)
        pf.space_after = Pt(space_after)
        pf.line_spacing = line_spacing
        if indent > 0:
            pf.first_line_indent = Pt(indent)
        else:
            pf.first_line_indent = Pt(0)
        if text:
            r = p.add_run(text)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(size)
            r.bold = bold
            r.italic = italic
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pf = p.paragraph_format
        pf.space_before = Pt(14)
        pf.space_after = Pt(8)
        pf.first_line_indent = Pt(0)
        pf.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
        r.bold = True
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(10)
        pf.space_after = Pt(6)
        pf.first_line_indent = Pt(0)
        pf.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
        r.bold = True
        return p

    def add_listing_caption(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(5)
        pf.space_after = Pt(4)
        pf.first_line_indent = Pt(0)
        pf.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
        r.bold = False
        return p

    def add_table_caption(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(6)
        pf.space_after = Pt(3)
        pf.first_line_indent = Pt(0)
        pf.keep_with_next = True
        r = p.add_run(text)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)
        r.bold = False
        return p

    def add_code_listing(code_lines):
        tbl = doc.add_table(rows=1, cols=1)
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        cell = tbl.cell(0, 0)
        
        tcPr = cell._tc.get_or_add_tcPr()
        tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:type="dxa" w:w="9638"/>')
        tcPr.append(tcW)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F2F2F2"/>')
        tcPr.append(shd)

        tcMar = parse_xml(
            f'<w:tcMar {nsdecls("w")}>'
            f'  <w:top w:w="60" w:type="dxa"/>'
            f'  <w:left w:w="120" w:type="dxa"/>'
            f'  <w:bottom w:w="60" w:type="dxa"/>'
            f'  <w:right w:w="120" w:type="dxa"/>'
            f'</w:tcMar>'
        )
        tcPr.append(tcMar)

        tcBorders = parse_xml(
            f'<w:tcBorders {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>'
            f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>'
            f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="D0D5DD"/>'
            f'</w:tcBorders>'
        )
        tcPr.append(tcBorders)

        for i, line in enumerate(code_lines):
            p = cell.paragraphs[0] if i == 0 else cell.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            pf.first_line_indent = Pt(0)
            r = p.add_run(line)
            r.font.name = 'Courier New'
            r.font.size = Pt(9.5)

        after_p = doc.add_paragraph()
        after_p.paragraph_format.space_before = Pt(0)
        after_p.paragraph_format.space_after = Pt(4)
        after_p.paragraph_format.line_spacing = 1.0

    def add_data_table(headers, rows, col_widths=None):
        tbl = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        tblPr = tbl._tbl.tblPr
        tblBorders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>'
            f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="B0B0B0"/>'
            f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="B0B0B0"/>'
            f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="B0B0B0"/>'
            f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="B0B0B0"/>'
            f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
            f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
            f'</w:tblBorders>'
        )
        tblPr.append(tblBorders)

        # Header Row
        hdr_row = tbl.rows[0]
        hdr_row._tr.get_or_add_trPr().append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
        for ci, h in enumerate(headers):
            cell = hdr_row.cells[ci]
            tcPr = cell._tc.get_or_add_tcPr()
            tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="EBF1F5"/>'))
            tcMar = parse_xml(
                f'<w:tcMar {nsdecls("w")}>'
                f'  <w:top w:w="40" w:type="dxa"/>'
                f'  <w:bottom w:w="40" w:type="dxa"/>'
                f'  <w:left w:w="60" w:type="dxa"/>'
                f'  <w:right w:w="60" w:type="dxa"/>'
                f'</w:tcMar>'
            )
            tcPr.append(tcMar)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            pf = p.paragraph_format
            pf.space_before = Pt(1)
            pf.space_after = Pt(1)
            pf.line_spacing = 1.0
            pf.first_line_indent = Pt(0)
            r = p.add_run(h)
            r.font.name = 'Times New Roman'
            r.font.size = Pt(10)
            r.bold = True

        # Data Rows
        for ri, row_data in enumerate(rows):
            row = tbl.rows[ri + 1]
            row_trPr = row._tr.get_or_add_trPr()
            row_trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            for ci, val in enumerate(row_data):
                cell = row.cells[ci]
                tcPr = cell._tc.get_or_add_tcPr()
                if ri % 2 == 1:
                    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="F9FBFC"/>'))
                tcMar = parse_xml(
                    f'<w:tcMar {nsdecls("w")}>'
                    f'  <w:top w:w="30" w:type="dxa"/>'
                    f'  <w:bottom w:w="30" w:type="dxa"/>'
                    f'  <w:left w:w="60" w:type="dxa"/>'
                    f'  <w:right w:w="60" w:type="dxa"/>'
                    f'</w:tcMar>'
                )
                tcPr.append(tcMar)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if ci > 0 else WD_ALIGN_PARAGRAPH.CENTER
                pf = p.paragraph_format
                pf.space_before = Pt(1)
                pf.space_after = Pt(1)
                pf.line_spacing = 1.0
                pf.first_line_indent = Pt(0)
                r = p.add_run(str(val))
                r.font.name = 'Times New Roman'
                r.font.size = Pt(10)

        # Set column widths if provided
        if col_widths:
            for row in tbl.rows:
                for ci, w in enumerate(col_widths):
                    row.cells[ci].width = w

        after_p = doc.add_paragraph()
        after_p.paragraph_format.space_before = Pt(0)
        after_p.paragraph_format.space_after = Pt(4)
        after_p.paragraph_format.line_spacing = 1.0

    print("Generating Academic Report...")

    # ==========================================
    # 1. ТИТУЛЬНЫЙ ЛИСТ (Page 1)
    # ==========================================
    add_p("ГОУ «Приднестровский государственный университет им. Т.Г. Шевченко»", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("Физико-технический институт", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("Факультет информатики и вычислительной техники", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("Кафедра программного обеспечения вычислительной техники", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=18, indent=0, line_spacing=1.0)

    add_p("ОТЧЕТ ПО ПРАКТИКЕ", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=8, indent=0, bold=True, line_spacing=1.0)
    add_p("Технологическая (проектно-технологическая) практика", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=18, indent=0, line_spacing=1.0)

    add_p("Направление 09.03.04 «Программная инженерия»", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("Профиль: «Разработка программно-информационных систем»", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=12, indent=0, line_spacing=1.0)

    add_p("Студент 2 курса ФТ24ДР62ПИ группа", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("форма обучения очная", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("Афанасьев М.А.", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=12, indent=0, line_spacing=1.0)

    add_p("Место прохождения практики «ПГУ им. Т.Г. Шевченко, ФТИ, учебно-вычислительный центр»", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("Сроки прохождения практики", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("С «01» июня 2026 г. по «13» июля 2026 г.", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=12, indent=0, line_spacing=1.0)

    add_p("Руководители практики:", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("От профильной организации", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("зав. кафедрой ПОВТ Помян С. В.", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=8, indent=0, line_spacing=1.0)
    add_p("От университета", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0, indent=0, line_spacing=1.0)
    add_p("ст. преподаватель Е.В. Терещенко.", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=12, indent=0, line_spacing=1.0)

    add_p("Итоговая оценка (зачет) по практике ______________________________", align=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=20, indent=0, line_spacing=1.0)
    add_p("Тирасполь, 2026 г.", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=0, indent=0, line_spacing=1.0)

    doc.add_page_break()

    # ==========================================
    # 2. ОГЛАВЛЕНИЕ (Page 2)
    # ==========================================
    add_p("ОГЛАВЛЕНИЕ", align=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=14, indent=0, bold=True, line_spacing=1.0)

    toc_items = [
        ("ВВЕДЕНИЕ", "toc_intro"),
        ("1 ОПИСАНИЕ ПРЕДМЕТНОЙ ОБЛАСТИ", "toc_ch1"),
        ("1.1 Общие сведения", "toc_ch1_1"),
        ("2 ПОСТАНОВКА ЗАДАЧИ", "toc_ch2"),
        ("3 РУКОВОДСТВО ПРОГРАММИСТА", "toc_ch3"),
        ("3.1 Введение", "toc_ch3_1"),
        ("3.2 Общие сведения и архитектура решения", "toc_ch3_2"),
        ("3.3 Описание типов данных и схемы БД", "toc_ch3_3"),
        ("3.4 Описание исходных текстов программного продукта", "toc_ch3_4"),
        ("3.5 Вид программного продукта", "toc_ch3_5"),
        ("4 РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ", "toc_ch4"),
        ("4.1 Введение", "toc_ch4_1"),
        ("4.2 Системные требования", "toc_ch4_2"),
        ("4.3 Запуск и подробные сценарии работы со сквозными примерами", "toc_ch4_3"),
        ("ЗАКЛЮЧЕНИЕ", "toc_conclusion"),
        ("СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ", "toc_lit"),
    ]

    toc_paragraphs = {}
    for title, key in toc_items:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(3)
        pf.line_spacing = 1.0
        pf.first_line_indent = Pt(0)
        
        pPr = p._p.get_or_add_pPr()
        tabs = parse_xml(f'<w:tabs {nsdecls("w")}><w:tab w:val="right" w:leader="dot" w:pos="9638"/></w:tabs>')
        pPr.append(tabs)

        r_title = p.add_run(title)
        r_title.font.name = 'Times New Roman'
        r_title.font.size = Pt(14)

        r_tab = p.add_run("\t")
        r_tab.font.name = 'Times New Roman'

        r_num = p.add_run("0")
        r_num.font.name = 'Times New Roman'
        r_num.font.size = Pt(14)
        
        toc_paragraphs[key] = (p, r_num)

    doc.add_page_break()

    # ==========================================
    # 3. ВВЕДЕНИЕ
    # ==========================================
    add_h1("ВВЕДЕНИЕ")

    add_p("В современных социально-экономических реалиях деятельность некоммерческих и благотворительных организаций приобретает стратегическое значение для поддержки уязвимых слоёв населения, финансирования высокотехнологичной медицинской помощи, развития образовательных, экологических и гуманитарных программ. Однако расширение масштабов деятельности благотворительных фондов неизбежно сталкивается с вызовами строгого финансового контроля, прозрачности движения целевых пожертвований и своевременного формирования отчетности перед донорами, надзорными органами и обществом.")

    add_p("Традиционные подходы к учету пожертвований и расходов, основанные на ведении разрозненных электронных таблиц или бумажных журналов, характеризуются высоким риском человеческой ошибки, отсутствием разграничения прав доступа, фрагментарностью данных и невозможностью оперативного получения комплексной финансово-аналитической информации. В связи с этим разработка и внедрение специализированной программной системы автоматизации учета благотворительного фонда является актуальной научно-практической задачей современной программной инженерии.")

    add_p("Целью настоящей технологической (проектно-технологической) практики являлось закрепление фундаментальных теоретических знаний в области объектно-ориентированного программирования, проектирования реляционных баз данных и многослойных информационных систем, а также практическая разработка законченного программного комплекса «Информационная система Благотворительный фонд». Разработанный продукт призван автоматизировать учет благотворителей (доноров), целевых благотворительных программ и проектов, входящих пожертвований, получателей помощи и расходных операций фонда, а также предоставить аналитический инструментарий для оценки эффективности деятельности организации.")

    add_p("Для достижения поставленной цели в ходе прохождения технологической практики были сформулированы и успешно решены следующие задачи:")
    add_p("— провести детальный системный анализ предметной области благотворительной организации, формализовать основные бизнес-процессы фонда и построить инфологическую модель данных;", indent=0, space_after=3)
    add_p("— спроектировать и реализовать в среде Microsoft SQL Server реляционную базу данных в третьей нормальной форме с необходимыми первичными и внешними ключами, проверочными ограничениями и индексами;", indent=0, space_after=3)
    add_p("— обосновать выбор архитектурного шаблона Layered Architecture и реализовать разделение системы на слои Models, DataAccess (Repositories), Services и Presentation;", indent=0, space_after=3)
    add_p("— разработать программный компонент взаимодействия с СУБД с применением технологии прозрачного ADO.NET (библиотека Microsoft.Data.SqlClient) на основе параметризованных команд, обеспечивающих абсолютную защиту от SQL-инъекций;", indent=0, space_after=3)
    add_p("— реализовать пять обязательных регламентных аналитических выборок: расчет сумм пожертвований по проектам за произвольный период, реестр активных проектов с процентом выполнения, рейтинг крупнейших доноров, сводный баланс фонда и постатейные расходы, а также распределение помощи по целевым категориям;", indent=0, space_after=3)
    add_p("— разработать эргономичный графический интерфейс пользователя на платформе Windows Forms (C# .NET 8) с удобной навигацией через единое окно, табличным представлением данных, фильтрацией, поиском и функцией экспорта отчетов в формат CSV;", indent=0, space_after=3)
    add_p("— провести всестороннее модульное и интеграционное тестирование функциональности программного комплекса и оформить комплект программной документации в строгом соответствии с требованиями ЕСПД и ГОСТ.", indent=0, space_after=6)

    add_p("В процессе разработки программного комплекса были использованы: современный объектно-ориентированный язык программирования C# 12, платформа .NET 8.0, графическая подсистема Windows Forms, реляционная СУБД Microsoft SQL Server (с поддержкой локальной базы LocalDB и серверных редакций 2019/2022), официальный ADO.NET-провайдер Microsoft.Data.SqlClient, а также встроенные средства экспорта данных в стандартные табличные форматы.")

    # ==========================================
    # 4. 1 ОПИСАНИЕ ПРЕДМЕТНОЙ ОБЛАСТИ
    # ==========================================
    add_h1("1 ОПИСАНИЕ ПРЕДМЕТНОЙ ОБЛАСТИ")
    add_h2("1.1 Общие сведения")

    add_p("Предметной областью разработанного программного комплекса является операционная, финансовая и аналитическая деятельность некоммерческой благотворительной организации. Благотворительный фонд осуществляет аккумулирование добровольных имущественных и денежных взносов граждан и юридических лиц для последующего строго целевого финансирования социально значимых программ, адресной помощи гражданам, оказавшимся в трудной жизненной ситуации, а также поддержки медицинских, образовательных и социальных учреждений.")

    add_p("Специфика функционирования благотворительного фонда накладывает жесткие требования к учету движения денежных средств. В отличие от коммерческих предприятий, основной целью которых является извлечение прибыли, фонд оперирует целевыми пожертвованиями. Каждая поступившая денежная сумма имеет строго определенное назначение: либо она направляется на реализацию конкретного благотворительного проекта (например, сбор средств на проведение сложной хирургической операции ребенку или приобретение оборудования для детского дома), либо аккумулируется в рамках уставной категории помощи (медицина, поддержка сирот, помощь пожилым людям, экология, ликвидация последствий стихийных бедствий). Нецелевое расходование пожертвований недопустимо ни с юридической, ни с морально-этической точки зрения.")

    add_p("В информационной системе выделены шесть ключевых взаимосвязанных сущностей предметной области, всесторонне отражающих все контуры операционной работы фонда:")
    add_p("1. Категории программ и помощи (Categories) — базовый классификатор уставных направлений деятельности фонда. Категории определяют профиль оказываемой помощи (например, «Медицина и лечение», «Поддержка детей-сирот», «Забота о пожилых людях», «Экологические инициативы», «Экстренная гуманитарная помощь»). Классификатор позволяет структурировать проекты и заявки нуждающихся, а также строить сводную аналитику о приоритетах благотворительной активности.", indent=0, space_after=3)
    add_p("2. Доноры (Donors) — физические лица, коммерческие организации, корпоративные партнеры и анонимные благотворители, совершающие пожертвования в пользу фонда. Сущность донора характеризуется полным наименованием (ФИО гражданина или юридическое наименование предприятия), типом донора («Физическое лицо», «Юридическое лицо», «Анонимный благотворитель»), контактным лицом (для юридических лиц), номером контактного телефона, адресом электронной почты, фактическим адресом и датой первичной регистрации в системе фонда.", indent=0, space_after=3)
    add_p("3. Благотворительные проекты (Projects) — целевые публичные программы сбора средств, инициируемые фондом. Проект характеризуется уникальным идентификатором, понятным наименованием, привязкой к конкретной категории помощи, целевым объемом сбора (TargetAmount в рублях), текущей фактически аккумулированной суммой (CurrentAmount), датой старта сбора, плановой датой закрытия, статусом реализации («Активен», «Завершен», «Приостановлен») и подробным описанием решаемой проблемы.", indent=0, space_after=3)
    add_p("4. Пожертвования (Donations) — финансовые транзакции поступления средств от благотворителей. Каждое пожертвование строго связывается с конкретным донором и целевым проектом. Запись содержит сумму платежа (Amount), точную дату и время транзакции, способ перечисления («Банковский перевод», «Банковская карта», «Наличные средства», «Электронные платежи») и сопроводительные заметки или назначение платежа. Проведение пожертвования должно автоматически и атомарно увеличивать текущую сумму сбора соответствующего проекта.", indent=0, space_after=3)
    add_p("5. Получатели благотворительной помощи (Recipients) — граждане, семьи или социальные учреждения, официально обратившиеся в фонд за поддержкой. Запись получателя включает ФИО или наименование учреждения, профильную категорию нужды, телефон, фактический адрес проживания или местонахождения, подробное описание жизненной ситуации и требуемой помощи (NeedDescription), статус рассмотрения обращения («На рассмотрении», «Одобрено», «Помощь оказана», «Отклонено») и дату подачи заявки.", indent=0, space_after=3)
    add_p("6. Расходы фонда (FundExpenses) — финансовые операции целевого расходования аккумулированных средств. Расход оформляется в рамках конкретного проекта и может быть адресован конкретному зарегистрированному и одобренному получателю (либо носить общий проектный характер, например закупка партии медикаментов). Расходная запись фиксирует сумму, дату операции, назначение платежа (Purpose — например, оплата счета клиники, закупка инвалидных колясок, продуктовых наборов), а также ФИО ответственного сотрудника фонда, санкционировавшего выдачу средств.", indent=0, space_after=6)

    add_p("Жизненный цикл движения денежных средств в рамках предметной области фонда представлен следующей цепочкой операционных событий: регистрация донора -> открытие целевого проекта -> поступление пожертвований с пополнением баланса проекта -> регистрация и экспертная проверка заявки нуждающегося лица -> вынесение решения об одобрении помощи -> оформление расходной операции с указанием конкретной статьи затрат -> закрытие проекта по достижении целевой суммы или наступлении планового срока -> формирование публичной отчетности.")

    add_p("Ключевыми бизнес-правилами предметной области, реализация которых заложена в архитектуру базы данных и логику приложения, являются: абсолютная недопустимость отрицательных или нулевых сумм в транзакциях (CHECK Amount > 0); соблюдение финансового лимита (списания по проекту не могут превышать сумму собранных по нему средств); сохранение ссылочной целостности истории операций (блокировка удаления проектов и доноров с финансовыми проводками); обеспечение конфиденциальности для анонимных благотворителей при соблюдении бухгалтерской дисциплины.")

    # ==========================================
    # 5. 2 ПОСТАНОВКА ЗАДАЧИ
    # ==========================================
    add_h1("2 ПОСТАНОВКА ЗАДАЧИ")

    add_p("Главной целью разработки программного продукта является создание надежного, быстродействующего и безопасного настольного приложения для операционной системы Windows, обеспечивающего комплексную автоматизацию процессов учета доноров, сборов пожертвований, ведения проектов, учета обращений нуждающихся и расходных операций благотворительного фонда с централизованным хранением информации в реляционной СУБД Microsoft SQL Server.")

    add_p("Актуальность разработки обусловлена необходимостью перехода благотворительных организаций от неструктурированных электронных таблиц к промышленным системам управления базами данных, обеспечивающим принцип единого источника достоверных данных (Single Source of Truth), высокую надежность при многопользовательском доступе, атомарность финансовых транзакций и возможность глубокого аналитического мониторинга.")

    add_p("Для реализации поставленной цели в рамках технологической практики были сформулированы следующие требования и задачи:")
    add_p("— спроектировать концептуальную, инфологическую и физическую схемы базы данных, обеспечив третью нормальную форму (3NF), отсутствие аномалий обновления, удаления и вставки, а также индексирование внешних ключей и часто запрашиваемых полей;", indent=0, space_after=3)
    add_p("— реализовать структуру базы данных CharityFund в Microsoft SQL Server с помощью скрипта CharityFund_CreateDB.sql, содержащего DDL-инструкции создания таблиц, связей, индексов и начального набора тестовых данных;", indent=0, space_after=3)
    add_p("— организовать подключение к СУБД с поддержкой как локальных экземпляров (LocalDB, SQL Express), так и выделенных сетевых серверов SQL Server с возможностью интерактивного тестирования соединения из интерфейса программы;", indent=0, space_after=3)
    add_p("— реализовать компонент DataAccess с использованием прямого ADO.NET (Microsoft.Data.SqlClient), инкапсулирующий создание подключений (SqlConnection), параметризованных команд (SqlCommand) и чтение потока данных (SqlDataReader);", indent=0, space_after=3)
    add_p("— обеспечить строгую транзакционную целостность при регистрации пожертвований: вставка записи в таблицу пожертвований и инкремент собранной суммы проекта должны выполняться в единой транзакции (SqlTransaction) с уровнем изоляции Read Committed;", indent=0, space_after=3)
    add_p("— реализовать пять обязательных аналитических выборок, регламентированных кафедрой ПОВТ, для всесторонней оценки финансового состояния и результативности фонда;", indent=0, space_after=3)
    add_p("— разработать модульный сервис экспорта любых табличных представлений и аналитических выборок в формат CSV (RFC 4180) с поддержкой кодировки UTF-8 с маркером последовательности байтов (BOM) для корректного отображения в Microsoft Excel;", indent=0, space_after=3)
    add_p("— создать эргономичный графический интерфейс (Windows Forms) с использованием паттерна Single-Form Navigation (боковая панель меню, динамическое переключение представлений без открытия множества хаотичных окон);", indent=0, space_after=3)
    add_p("— разработать систему модальных диалогов для создания и редактирования записей с валидацией пользовательского ввода, подсказками и перехватом исключений.", indent=0, space_after=6)

    add_p("Входными данными системы являются: конфигурационные параметры подключения к СУБД; регистрационные анкеты доноров (ФИО/наименование, тип, контактное лицо, телефон, email, адрес); параметры проектов (наименование, категория, плановый бюджет, даты, описание); реквизиты платежных проводок (донор, проект, сумма, способ оплаты, примечание); анкеты нуждающихся лиц (категория, телефон, адрес, описание нужды); финансовые реквизиты списаний (проект, получатель, сумма, назначение, ответственный сотрудник); команды фильтрации, поиска и экспорта.")

    add_p("Выходными данными системы являются: интерактивные табличные представления (DataGridView) реестров доноров, проектов, транзакций, обращений и расходов; информационные панели ключевых метрик дашборда (общие сборы, расходы, баланс фонда, активные проекты); пять регламентных аналитических отчетов со сводными финансовыми показателями; структурированные файлы формата CSV для интеграции с Excel; системные диалоговые сообщения об успешных проводках и ошибках валидации.")

    add_p("В соответствии с заданием кафедры реализованы пять обязательных аналитических выборок, сведенных в таблицу 7.")

    add_table_caption("Таблица 7 — Сводная матрица аналитических запросов ИС «Благотворительный фонд»")
    tbl7_headers = ["№", "Наименование выборки", "Используемые конструкции SQL", "Управленческое назначение"]
    tbl7_rows = [
        ["1", "Сумма пожертвований", "SUM, COUNT, GROUP BY Projects.Title, HAVING, BETWEEN", "Контроль динамики сборов за период, оценка сборов по проектам."],
        ["2", "Активные проекты", "WHERE Status = 'Активен', (Current/Target)*100, Target - Current", "Мониторинг сборов, выявление проектов, отстающих от графика."],
        ["3", "Крупнейшие доноры", "SUM(Amount), COUNT(Id), MAX(Date), GROUP BY, ORDER BY SUM DESC", "Формирование пула ключевых партнеров, персонализация благодарностей."],
        ["4", "Расходы фонда", "SUM(Donations) - SUM(Expenses), GROUP BY Project, Purpose", "Обеспечение финансовой устойчивости, контроль целевого расхода."],
        ["5", "Помощь по категориям", "COUNT(Recipients), SUM(Expenses), (CatExp / TotalExp)*100", "Стратегический анализ структуры помощи и балансировка направлений."]
    ]
    add_data_table(tbl7_headers, tbl7_rows, [Cm(1.0), Cm(4.0), Cm(5.5), Cm(6.5)])

    # ==========================================
    # 6. 3 РУКОВОДСТВО ПРОГРАММИСТА
    # ==========================================
    add_h1("3 РУКОВОДСТВО ПРОГРАММИСТА")
    add_h2("3.1 Введение")

    add_p("Настоящее руководство программиста предназначено для специалистов в области программной инженерии, системных разработчиков и администраторов баз данных, осуществляющих развертывание, сопровождение, аудит кода, модернизацию или функциональное расширение программного комплекса «Информационная система Благотворительный фонд».")

    add_p("Документ содержит подробное описание многослойной архитектуры приложения, спецификацию типов данных и физической схемы базы данных Microsoft SQL Server, детальный разбор исходных текстов ключевых компонентов слоя доступа к данным (ADO.NET), а также архитектурную компоновку экранных форм графического интерфейса.")

    add_p("Для успешного сопровождения программного продукта разработчик должен владеть языком C# (стандарты C# 10–12), платформой .NET 8.0, принципами объектно-ориентированного проектирования (SOLID, Layered Architecture), языком структурированных запросов T-SQL, механизмом параметризованных запросов и транзакций технологии ADO.NET (библиотека Microsoft.Data.SqlClient), а также технологией построения десктопных интерфейсов Windows Forms.")

    add_p("Для сборки проекта из исходных текстов, отладки и запуска необходим следующий инструментарий: операционная система Microsoft Windows 10/11; комплект разработчика .NET 8.0 SDK (версия 8.0.200 или новее); среда разработки Visual Studio 2022 или JetBrains Rider; сервер СУБД Microsoft SQL Server (LocalDB / Express / Standard 2019/2022); среда администрирования SQL Server Management Studio (SSMS).")

    add_h2("3.2 Общие сведения и архитектура решения")

    add_p("В основу архитектуры программного комплекса положен классический многослойный шаблон проектирования (Layered Architecture). Применение данного архитектурного подхода обеспечивает слабую связанность (Loose Coupling) между подсистемой пользовательского интерфейса, бизнес-логикой и механизмами взаимодействия с СУБД, что значительно повышает тестируемость, расширяемость и простоту сопровождения программного кода.")

    add_p("Программное решение структурировано на следующие четыре логических слоя:")
    add_p("1. Слой сущностей и моделей данных (Domain Models & DTOs) — не содержит зависимостей от других слоев и внешних библиотек. Включает базовые классы предметной области (Category, Donor, Project, Donation, Recipient, FundExpense), а также специализированные объекты передачи данных (Data Transfer Objects), предназначенные для возврата агрегированных результатов аналитических запросов (DonationSummaryDto, ActiveProjectDto, TopDonorDto, FundBalanceDto, CategoryAidDto).", indent=0, space_after=3)
    add_p("2. Слой доступа к данным (Data Access Layer / Repositories) — инкапсулирует всю логику прямого взаимодействия с Microsoft SQL Server. Содержит статический вспомогательный класс DatabaseHelper, отвечающий за централизованное хранение строки подключения, создание соединений и безопасное исполнение команд, а также специализированные репозитории: CategoryRepository, DonorRepository, ProjectRepository, DonationRepository, RecipientRepository, FundExpenseRepository и AnalyticsRepository. Каждый репозиторий реализует интерфейсы доступа к данным, транслируя бизнес-операции в оптимизированные параметризованные запросы T-SQL.", indent=0, space_after=3)
    add_p("3. Слой сервисов и прикладной логики (Services / Business Logic) — отвечает за координацию бизнес-правил, сложную валидацию введенных данных перед сохранением в БД, вычисление производных финансовых показателей (например, сальдо фонда и процент выполнения планов сборов), а также за форматирование и экспорт аналитических выборок в файлы CSV через компонент ExportService.", indent=0, space_after=3)
    add_p("4. Слой представления (Presentation Layer / Windows Forms) — графический интерфейс пользователя. Построен на концепции Single-Form Navigation (форма MainForm). Форма содержит боковую панель навигации (Sidebar Navigation), заголовочную панель со статусом связи с СУБД и контейнерную область панели страниц. Для создания и изменения записей применяются специализированные модальные диалоги (DonorEditForm, ProjectEditForm, DonationEditForm, RecipientEditForm, ExpenseEditForm, ConnectionSettingsForm), использующие двухстороннее связывание данных и строгую валидацию полей перед закрытием диалога.", indent=0, space_after=6)

    add_p("Принципиальным инженерным решением проекта явился осознанный отказ от тяжеловесных объектно-реляционных отображателей (Entity Framework Core / NHibernate) в пользу технологии чистого низкоуровневого ADO.NET на основе библиотеки Microsoft.Data.SqlClient. Данный выбор обеспечивает: максимальное быстродействие при потоковом чтении данных через SqlDataReader; абсолютный детерминированный контроль T-SQL запросов без скрытых накладных расходов; явное и надежное управление транзакциями через SqlTransaction; гарантированную защиту от SQL Injection за счет строгой типизации параметров SqlParameter.")

    add_p("Управление сетевыми соединениями в приложении организовано с использованием пула соединений (Connection Pooling), встроенного в ADO.NET. Все операции с базой данных оборачиваются в конструкции using (var connection = DatabaseHelper.GetConnection()), что гарантирует своевременное возвращение дескриптора соединения в пул даже при возникновении критических исключений во время исполнения запроса.")

    add_h2("3.3 Описание типов данных и схемы БД")

    add_p("База данных CharityFund спроектирована в реляционной парадигме и находится в третьей нормальной форме (3NF). Все атрибуты являются атомарными (1NF), каждый неключевой атрибут функционально полно зависит от первичного ключа (2NF) и между неключевыми атрибутами отсутствуют транзитивные зависимости (3NF). Ссылочная целостность данных строго обеспечивается на уровне СУБД с помощью внешних ключей FOREIGN KEY.")

    add_p("Ниже, в таблицах 1–6, представлено подробное описание структуры всех шести таблиц базы данных с указанием наименований полей, типов данных SQL Server, модификаторов ограничений и смыслового назначения.")

    add_table_caption("Таблица 1 — Структура таблицы Categories (Категории программ и помощи)")
    tbl1_headers = ["Имя поля", "Тип данных SQL", "Ограничения", "Назначение поля"]
    tbl1_rows = [
        ["Id", "INT", "PK, IDENTITY(1,1)", "Уникальный суррогатный идентификатор категории"],
        ["Name", "NVARCHAR(100)", "NOT NULL, UNIQUE", "Уникальное название направления (Медицина, Сироты и др.)"],
        ["Description", "NVARCHAR(500)", "NULL", "Подробное описание целей и задач данного направления"]
    ]
    add_data_table(tbl1_headers, tbl1_rows, [Cm(2.5), Cm(3.5), Cm(4.5), Cm(6.5)])

    add_table_caption("Таблица 2 — Структура таблицы Donors (Доноры / Благотворители)")
    tbl2_headers = ["Имя поля", "Тип данных SQL", "Ограничения", "Назначение поля"]
    tbl2_rows = [
        ["Id", "INT", "PK, IDENTITY(1,1)", "Уникальный идентификатор благотворителя"],
        ["FullName", "NVARCHAR(150)", "NOT NULL", "ФИО гражданина или полное наименование юридического лица"],
        ["DonorType", "NVARCHAR(50)", "NOT NULL", "Тип донора (Физическое лицо, Юрлицо, Анонимный)"],
        ["ContactPerson", "NVARCHAR(100)", "NULL", "Контактное лицо (для предприятий и организаций)"],
        ["Phone", "NVARCHAR(50)", "NULL", "Номер телефона для оперативной связи и уведомлений"],
        ["Email", "NVARCHAR(100)", "NULL", "Адрес электронной почты для отправки отчетов фонда"],
        ["Address", "NVARCHAR(200)", "NULL", "Почтовый/юридический адрес благотворителя"],
        ["CreatedAt", "DATETIME2", "NOT NULL, DEFAULT GETDATE()", "Дата и время первичной регистрации в системе фонда"]
    ]
    add_data_table(tbl2_headers, tbl2_rows, [Cm(2.5), Cm(3.5), Cm(4.5), Cm(6.5)])

    add_table_caption("Таблица 3 — Структура таблицы Projects (Благотворительные проекты / сборы)")
    tbl3_headers = ["Имя поля", "Тип данных SQL", "Ограничения", "Назначение поля"]
    tbl3_rows = [
        ["Id", "INT", "PK, IDENTITY(1,1)", "Уникальный идентификатор благотворительного проекта"],
        ["Title", "NVARCHAR(150)", "NOT NULL", "Публичное название проекта или целевой программы сбора"],
        ["CategoryId", "INT", "FK -> Categories(Id)", "Внешний ключ привязки к категории помощи"],
        ["TargetAmount", "DECIMAL(18,2)", "CHECK (TargetAmount > 0)", "Целевой бюджет проекта, требуемый к сбору (руб.)"],
        ["CurrentAmount", "DECIMAL(18,2)", "DEFAULT 0, CHECK (>= 0)", "Фактически аккумулированная сумма пожертвований (руб.)"],
        ["StartDate", "DATE", "NOT NULL", "Дата официального старта сбора средств по проекту"],
        ["EndDate", "DATE", "NULL", "Плановая дата завершения проекта (если ограничена)"],
        ["Status", "NVARCHAR(50)", "NOT NULL", "Текущий статус (Активен, Завершен, Приостановлен)"],
        ["Description", "NVARCHAR(1000)", "NULL", "Детальное обоснование целевого назначения сбора"]
    ]
    add_data_table(tbl3_headers, tbl3_rows, [Cm(2.5), Cm(3.5), Cm(4.5), Cm(6.5)])

    add_table_caption("Таблица 4 — Структура таблицы Donations (Пожертвования)")
    tbl4_headers = ["Имя поля", "Тип данных SQL", "Ограничения", "Назначение поля"]
    tbl4_rows = [
        ["Id", "INT", "PK, IDENTITY(1,1)", "Уникальный идентификатор платежной проводки"],
        ["DonorId", "INT", "FK -> Donors(Id)", "Внешний ключ благотворителя, совершившего взнос"],
        ["ProjectId", "INT", "FK -> Projects(Id)", "Внешний ключ целевого проекта сбора средств"],
        ["Amount", "DECIMAL(18,2)", "CHECK (Amount > 0)", "Сумма поступившего пожертвования в рублях"],
        ["DonationDate", "DATETIME2", "DEFAULT GETDATE()", "Точные дата и время фиксации финансовой транзакции"],
        ["PaymentMethod", "NVARCHAR(50)", "NOT NULL", "Способ оплаты (Банковский перевод, Карта, Наличные)"],
        ["Notes", "NVARCHAR(500)", "NULL", "Назначение платежа, комментарий благотворителя"]
    ]
    add_data_table(tbl4_headers, tbl4_rows, [Cm(2.5), Cm(3.5), Cm(4.5), Cm(6.5)])

    add_table_caption("Таблица 5 — Структура таблицы Recipients (Получатели благотворительной помощи)")
    tbl5_headers = ["Имя поля", "Тип данных SQL", "Ограничения", "Назначение поля"]
    tbl5_rows = [
        ["Id", "INT", "PK, IDENTITY(1,1)", "Уникальный идентификатор заявления получателя"],
        ["FullName", "NVARCHAR(150)", "NOT NULL", "ФИО гражданина или официальное название учреждения"],
        ["CategoryId", "INT", "FK -> Categories(Id)", "Профильное направление требуемой социальной помощи"],
        ["Phone", "NVARCHAR(50)", "NULL", "Контактный номер телефона для связи"],
        ["Address", "NVARCHAR(200)", "NULL", "Адрес регистрации / фактического проживания"],
        ["NeedDescription", "NVARCHAR(1000)", "NOT NULL", "Подробное описание трудной ситуации, медицинский диагноз"],
        ["Status", "NVARCHAR(50)", "NOT NULL", "Статус заявки (На рассмотрении, Одобрено, Оказана, Отклонена)"],
        ["CreatedAt", "DATETIME2", "DEFAULT GETDATE()", "Дата первичной подачи и регистрации обращения в фонде"]
    ]
    add_data_table(tbl5_headers, tbl5_rows, [Cm(2.5), Cm(3.5), Cm(4.5), Cm(6.5)])

    add_table_caption("Таблица 6 — Структура таблицы FundExpenses (Расходы фонда / Выдача помощи)")
    tbl6_headers = ["Имя поля", "Тип данных SQL", "Ограничения", "Назначение поля"]
    tbl6_rows = [
        ["Id", "INT", "PK, IDENTITY(1,1)", "Уникальный номер расходного финансового ордера"],
        ["ProjectId", "INT", "FK -> Projects(Id)", "Внешний ключ проекта, с баланса которого списаны средства"],
        ["RecipientId", "INT", "FK -> Recipients(Id), NULL", "Внешний ключ получателя (если помощь адресная)"],
        ["Amount", "DECIMAL(18,2)", "CHECK (Amount > 0)", "Сумма расходованных средств фонда в рублях"],
        ["ExpenseDate", "DATETIME2", "DEFAULT GETDATE()", "Дата и время фактического осуществления списания"],
        ["Purpose", "NVARCHAR(250)", "NOT NULL", "Статья расхода (закупка медикаментов, оплата операции)"],
        ["ResponsiblePerson", "NVARCHAR(100)", "NOT NULL", "ФИО сотрудника фонда, санкционировавшего списание"]
    ]
    add_data_table(tbl6_headers, tbl6_rows, [Cm(2.5), Cm(3.5), Cm(4.5), Cm(6.5)])

    add_p("В доменном слое приложения каждой реляционной таблице соответствует строго типизированный C#-класс сущности. Например, класс Project содержит свойства Id, Title, CategoryId, TargetAmount, CurrentAmount, StartDate, EndDate, Status, Description, а также вычисляемое свойство ProgressPercentage ((CurrentAmount / TargetAmount) * 100). Для передачи результатов аналитических запросов разработаны специализированные DTO-классы с компактными свойствами только для чтения, что исключает избыточные аллокации памяти.")

    add_h2("3.4 Описание исходных текстов программного продукта")

    add_p("Проект организован в соответствии со строгой модульной структурой каталогов. В директории DataAccess сосредоточены классы взаимодействия с СУБД, в Models — структуры данных, в Services — прикладная логика экспорта и проверок, а в UI — формы интерфейса Windows Forms.")

    add_p("Центральным компонентом подсистемы доступа к данным является класс DatabaseHelper, представленный в листинге 1. Класс отвечает за чтение строки подключения, открытие пулированных соединений и безопасное выполнение SQL-команд.")

    add_listing_caption("Листинг 1 — Класс DatabaseHelper (DataAccess/DatabaseHelper.cs)")
    code_listing1 = [
        "public static class DatabaseHelper",
        "{",
        "    public static string ConnectionString { get; set; } = ",
        "        @\"Server=(localdb)\\mssqllocaldb;Database=CharityFund;Integrated Security=True;TrustServerCertificate=True;\";",
        "",
        "    public static SqlConnection GetConnection()",
        "    {",
        "        var connection = new SqlConnection(ConnectionString);",
        "        connection.Open();",
        "        return connection;",
        "    }",
        "",
        "    public static DataTable ExecuteDataTable(string sql, params SqlParameter[] parameters)",
        "    {",
        "        using var conn = GetConnection();",
        "        using var cmd = new SqlCommand(sql, conn);",
        "        if (parameters != null && parameters.Length > 0)",
        "            cmd.Parameters.AddRange(parameters);",
        "",
        "        using var adapter = new SqlDataAdapter(cmd);",
        "        var table = new DataTable();",
        "        adapter.Fill(table);",
        "        return table;",
        "    }",
        "}"
    ]
    add_code_listing(code_listing1)

    add_p("Особое значение для обеспечения финансовой целостности имеет транзакционная регистрация входящего пожертвования в классе DonationRepository (листинг 2). Вставка записи пожертвования и увеличение фактически собранной суммы проекта CurrentAmount выполняются в рамках единой транзакции SqlTransaction.")

    add_listing_caption("Листинг 2 — Транзакционная проводка пожертвования (DataAccess/DonationRepository.cs)")
    code_listing2 = [
        "public class DonationRepository",
        "{",
        "    public bool RegisterDonation(Donation donation, out string errorMsg)",
        "    {",
        "        errorMsg = string.Empty;",
        "        using var conn = DatabaseHelper.GetConnection();",
        "        using var transaction = conn.BeginTransaction(IsolationLevel.ReadCommitted);",
        "        try",
        "        {",
        "            const string insertSql = @\"INSERT INTO Donations (DonorId, ProjectId, Amount, DonationDate, PaymentMethod, Notes)",
        "                                       VALUES (@DonorId, @ProjectId, @Amount, @DonationDate, @PaymentMethod, @Notes);\";",
        "            using (var cmdInsert = new SqlCommand(insertSql, conn, transaction))",
        "            {",
        "                cmdInsert.Parameters.Add(new SqlParameter(\"@DonorId\", SqlDbType.Int) { Value = donation.DonorId });",
        "                cmdInsert.Parameters.Add(new SqlParameter(\"@ProjectId\", SqlDbType.Int) { Value = donation.ProjectId });",
        "                cmdInsert.Parameters.Add(new SqlParameter(\"@Amount\", SqlDbType.Decimal) { Value = donation.Amount });",
        "                cmdInsert.Parameters.Add(new SqlParameter(\"@DonationDate\", SqlDbType.DateTime2) { Value = donation.DonationDate });",
        "                cmdInsert.Parameters.Add(new SqlParameter(\"@PaymentMethod\", SqlDbType.NVarChar, 50) { Value = donation.PaymentMethod });",
        "                cmdInsert.Parameters.Add(new SqlParameter(\"@Notes\", SqlDbType.NVarChar, 500) { Value = (object)donation.Notes ?? DBNull.Value });",
        "                cmdInsert.ExecuteNonQuery();",
        "            }",
        "            const string updateSql = @\"UPDATE Projects SET CurrentAmount = CurrentAmount + @Amount WHERE Id = @ProjectId;\";",
        "            using (var cmdUpdate = new SqlCommand(updateSql, conn, transaction))",
        "            {",
        "                cmdUpdate.Parameters.Add(new SqlParameter(\"@Amount\", SqlDbType.Decimal) { Value = donation.Amount });",
        "                cmdUpdate.Parameters.Add(new SqlParameter(\"@ProjectId\", SqlDbType.Int) { Value = donation.ProjectId });",
        "                cmdUpdate.ExecuteNonQuery();",
        "            }",
        "            transaction.Commit();",
        "            return true;",
        "        }",
        "        catch (Exception ex) { transaction.Rollback(); errorMsg = ex.Message; return false; }",
        "    }",
        "}"
    ]
    add_code_listing(code_listing2)

    add_p("В листинге 3 представлена реализация аналитического репозитория AnalyticsRepository, содержащего регламентные методы формирования комплексных аналитических срезов.")

    add_listing_caption("Листинг 3 — Аналитический репозиторий (DataAccess/AnalyticsRepository.cs)")
    code_listing3 = [
        "public class AnalyticsRepository",
        "{",
        "    // Запрос 1: Сумма пожертвований с группировкой по проектам и фильтром по датам",
        "    public DataTable GetDonationsSummary(DateTime? fromDate = null, DateTime? toDate = null)",
        "    {",
        "        string sql = @\"SELECT p.Title AS [Проект], c.Name AS [Категория],",
        "                              COUNT(d.Id) AS [Число взносов], ISNULL(SUM(d.Amount), 0) AS [Собрано (руб.)],",
        "                              ISNULL(AVG(d.Amount), 0) AS [Средний чек (руб.)]",
        "                       FROM Projects p INNER JOIN Categories c ON p.CategoryId = c.Id",
        "                       LEFT JOIN Donations d ON p.Id = d.ProjectId",
        "                       WHERE (@FromDate IS NULL OR d.DonationDate >= @FromDate)",
        "                         AND (@ToDate IS NULL OR d.DonationDate <= @ToDate)",
        "                       GROUP BY p.Title, c.Name ORDER BY [Собрано (руб.)] DESC;\";",
        "        var p1 = new SqlParameter(\"@FromDate\", SqlDbType.DateTime2) { Value = (object)fromDate ?? DBNull.Value };",
        "        var p2 = new SqlParameter(\"@ToDate\", SqlDbType.DateTime2) { Value = (object)toDate ?? DBNull.Value };",
        "        return DatabaseHelper.ExecuteDataTable(sql, p1, p2);",
        "    }",
        "",
        "    // Запрос 2: Активные проекты с вычислением прогресса",
        "    public DataTable GetActiveProjectsProgress()",
        "    {",
        "        const string sql = @\"SELECT Title AS [Проект], TargetAmount AS [Бюджет (руб.)],",
        "                                   CurrentAmount AS [Собрано (руб.)],",
        "                                   CAST(ROUND((CurrentAmount / TargetAmount) * 100, 2) AS DECIMAL(5,2)) AS [Прогресс %],",
        "                                   (TargetAmount - CurrentAmount) AS [Осталось собрать (руб.)]",
        "                            FROM Projects WHERE Status = N'Активен' ORDER BY [Прогресс %] DESC;\";",
        "        return DatabaseHelper.ExecuteDataTable(sql);",
        "    }",
        "",
        "    // Запрос 3: Рейтинг крупнейших доноров",
        "    public DataTable GetTopDonors(int topCount = 10)",
        "    {",
        "        string sql = $@\"SELECT TOP ({topCount}) d.FullName AS [Благотворитель], d.DonorType AS [Тип],",
        "                               COUNT(dn.Id) AS [Число взносов], SUM(dn.Amount) AS [Всего (руб.)],",
        "                               MAX(dn.DonationDate) AS [Последний взнос]",
        "                        FROM Donors d INNER JOIN Donations dn ON d.Id = dn.DonorId",
        "                        GROUP BY d.Id, d.FullName, d.DonorType ORDER BY [Всего (руб.)] DESC;\";",
        "        return DatabaseHelper.ExecuteDataTable(sql);",
        "    }",
        "}"
    ]
    add_code_listing(code_listing3)

    add_h2("3.5 Вид программного продукта")

    add_p("Пользовательский интерфейс информационной системы «Благотворительный фонд» спроектирован по эргономичному принципу единого навигационного окна (Single Main Window with Side Panel Navigation). Данный подход избавляет пользователя от необходимости управлять десятками хаотично открывающихся дочерних окон и обеспечивает целостное восприятие рабочего пространства фонда.")

    add_p("Визуальное оформление выдержано в строгой корпоративной цветовой гамме, традиционной для благотворительных и финансовых систем: основной акцентный цвет — глубокий морской синий (#1F497D / SteelBlue), вспомогательный цвет фона бокового меню — темно-серый сланец (#2C3E50), фон рабочей области — нейтральный светло-серый (#F8F9FA), цвет активных кнопок и статусов — изумрудный (#28A745) и янтарный (#FFC107). В качестве основного системного экранного шрифта применен Segoe UI с размерами 9–11 pt для таблиц и списков и 14–18 pt для заголовков и карточек метрик.")

    add_p("Главное окно программы MainForm композиционно разделено на три основные функциональные зоны:")
    add_p("1. Верхняя информационная панель (Header Panel) — содержит стилизованную эмблему фонда, официальное наименование организации «Благотворительный фонд «Надежда и Созидание»», информацию о текущем авторизованном пользователе и визуальный светодиодный индикатор статуса соединения с Microsoft SQL Server (зеленый круг — связь стабильна, красный — ошибка соединения).", indent=0, space_after=3)
    add_p("2. Левая навигационная панель (Sidebar Navigation) — вертикальное меню с пиктограммами и текстовыми подписями разделов: «🏠 Дашборд фонда», «🏢 Благотворители (Доноры)», «📋 Проекты и программы», «💰 Журнал пожертвований», «👥 Получатели помощи», «💳 Расходы и выдачи», «📊 Аналитические выборки» и «⚙️ Параметры подключения к БД». Нажатие на кнопку меню плавно переключает отображаемый контент в центральной рабочей зоне.", indent=0, space_after=3)
    add_p("3. Центральная рабочая зона (Main Content Area) — многостраничный контейнер, в котором динамически активируется соответствующий раздел: «Дашборд» с KPI-плитками («Привлечено средств», «Профинансировано», «Сальдо», «Активных сборов»); «Доноры» с поиском по ФИО/телефону и фильтром типа; «Проекты» со шкалой прогресса и статусами; «Пожертвования» с фильтром по датам и проектам; «Получатели» с картотекой обращений; «Расходы» со статьями затрат; «Аналитические выборки» с пятью кнопками регламентных запросов и экспортом в CSV; «Параметры подключения к БД» для проверки соединения с сервером.", indent=0, space_after=6)

    add_p("Для операций добавления и изменения данных реализованы модальные диалоговые формы: DonorEditForm, ProjectEditForm, DonationEditForm, RecipientEditForm и ExpenseEditForm. Каждая форма содержит строгую проверку корректности ввода (валидация обязательных полей, проверка числовых форматов, дат и телефонов с выводом понятных пользователю подсказок).")

    # ==========================================
    # 7. 4 РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ
    # ==========================================
    add_h1("4 РУКОВОДСТВО ПОЛЬЗОВАТЕЛЯ")
    add_h2("4.1 Введение")

    add_p("Настоящее руководство пользователя предназначено для сотрудников и руководящего состава благотворительного фонда — координаторов благотворительных программ, специалистов по взаимодействию с благотворителями, финансовых менеджеров, бухгалтеров и администраторов системы.")

    add_p("Документ содержит исчерпывающее описание системных требований к аппаратному и программному обеспечению, порядок первоначального развертывания базы данных и настройки приложения, а также детальные сквозные сценарии выполнения ежедневных операционных и аналитических задач фонда.")

    add_p("Программный комплекс обеспечивает удобный и безопасный учет всех сторон благотворительной деятельности, гарантируя сохранение конфиденциальности персональных данных доноров и благополучателей и прозрачность распределения пожертвований.")

    add_h2("4.2 Системные требования")

    add_p("Для надежной и бесперебойной эксплуатации программного комплекса «Информационная система Благотворительный фонд» рабочие станции пользователей и сервер базы данных должны удовлетворять системным требованиям, приведенным в таблице 8.")

    add_table_caption("Таблица 8 — Системные требования к программно-аппаратному обеспечению")
    tbl8_headers = ["Компонент системы", "Минимальные требования", "Рекомендуемые требования"]
    tbl8_rows = [
        ["Процессор (CPU)", "2 ядра, x64, с тактовой частотой от 1.8 ГГц", "4 ядра и более, x64, с частотой от 2.5 ГГц"],
        ["Оперативная память (RAM)", "4 ГБ оперативной памяти", "8 ГБ оперативной памяти и более"],
        ["Накопитель (HDD/SSD)", "Свободное пространство от 500 МБ", "Высокоскоростной SSD, от 2 ГБ свободного места"],
        ["Видеоподсистема и монитор", "Разрешение экрана не менее 1280x720", "Разрешение экрана Full HD (1920x1080)"],
        ["Сетевое подключение", "Локальная сеть от 10 Мбит/с (при удаленной БД)", "Локальная сеть 100/1000 Мбит/с"],
        ["Операционная система", "Microsoft Windows 10 (версия 1809+)", "Microsoft Windows 10 / 11 (64-разрядная)"],
        ["Среда выполнения", ".NET Desktop Runtime 8.0 (x64)", ".NET Desktop Runtime 8.0 (последний релиз)"],
        ["СУБД", "MS SQL Server 2019 LocalDB / Express", "MS SQL Server 2019/2022 Standard / Enterprise"]
    ]
    add_data_table(tbl8_headers, tbl8_rows, [Cm(3.5), Cm(6.5), Cm(6.5)])

    add_h2("4.3 Запуск и подробные сценарии работы со сквозными примерами")

    add_p("Подготовительный этап развертывания. Перед первым запуском программного комплекса администратор системы развертывает базу данных путем выполнения скрипта CharityFund_CreateDB.sql в SQL Server Management Studio. Скрипт создает базу данных CharityFund, структуру таблиц, первичные и внешние ключи, проверочные ограничения и начальный демонстрационный набор категорий, проектов и доноров.")

    add_p("Ниже приведены подробные пошаговые сценарии типовых рабочих процессов сотрудника фонда со сквозными числовыми данными.")

    add_p("Сценарий 1. Первый запуск системы и проверка связи с базой данных.", bold=True)
    add_p("Пользователь запускает исполняемый файл CharityFundApp.exe. При старте приложение автоматически устанавливает связь с СУБД по строке подключения по умолчанию. При успешном соединении индикатор в правом верхнем углу главного окна окрашивается в зеленый цвет со статусом «Подключено: CharityFund на (localdb)\\mssqllocaldb», и загружаются актуальные показатели дашборда фонда. При необходимости изменения параметров сервера пользователь переходит во вкладку «Параметры подключения к БД», вводит сетевой адрес, выбирает тип проверки подлинности, нажимает «Проверить соединение» и «Сохранить и применить».")

    add_p("Сценарий 2. Регистрация нового корпоративного благотворителя.", bold=True)
    add_p("В боковом меню выбирается раздел «🏢 Благотворители (Доноры)» и нажимается кнопка «Добавить донора». В модальной форме заполняются поля: «ФИО / Наименование организации» — ООО «Тираспольтранс»; «Тип благотворителя» — «Юридическое лицо»; «Контактное лицо» — Смирнов Алексей Петрович (генеральный директор); «Телефон» — +373 777 12345; «Email» — info@tirastrans.com; «Фактический адрес» — г. Тирасполь, ул. Ленина, д. 45, оф. 12. После нажатия кнопки «Сохранить» программа валидирует данные, выполняет параметризованную вставку в таблицу Donors и отображает новую запись в таблице.")

    add_p("Сценарий 3. Открытие нового благотворительного проекта (целевого сбора).", bold=True)
    add_p("В разделе «📋 Проекты и программы» нажимается кнопка «Создать проект». Заполняются реквизиты: наименование — «Срочная кардиохирургическая операция для Максима К.»; категория — «Медицина и лечение»; целевая сумма сбора (TargetAmount) — 350 000,00 руб.; дата старта — 01.06.2026; плановое окончание — 31.08.2026; статус — «Активен»; описание — «Целевой сбор на проведение высокотехнологичной операции на сердце ребенку 7 лет». При нажатии кнопки «Сохранить» создается проект с начальной суммой сбора 0 руб. и прогрессом 0%.")

    add_p("Сценарий 4. Прием и проведение пожертвования.", bold=True)
    add_p("В разделе «💰 Журнал пожертвований» нажимается «Зарегистрировать пожертвование». В форме выбираются: донор — ООО «Тираспольтранс»; проект — «Срочная кардиохирургическая операция для Максима К.»; сумма — 100 000,00 руб.; метод оплаты — «Банковский перевод»; примечание — «Платежное поручение №412 от 05.06.2026». Нажатие кнопки «Провести пожертвование» инициирует транзакцию SqlTransaction: фиксируется запись в Donations и обновляется собранная сумма проекта до 100 000,00 руб. (прогресс достигает 28.57%).")

    add_p("Сценарий 5. Регистрация заявления нуждающегося гражданина и одобрение помощи.", bold=True)
    add_p("В разделе «👥 Получатели помощи» нажимается кнопка «Новое обращение». Заполняются поля: заявитель — Кузнецова Елена Сергеевна; категория — «Медицина и лечение»; телефон — +373 778 98765; адрес — г. Бендеры, ул. Суворова, д. 18, кв. 4; описание ситуации — «Врожденный порок сердца у сына Максима, пакет медицинских документов прикреплен к делу №104»; статус — «На рассмотрении». После проверки документов попечительским советом сотрудник открывает карточку и переводит статус в положение «Одобрено».")

    add_p("Сценарий 6. Оформление целевого расхода фонда.", bold=True)
    add_p("В разделе «💳 Расходы и выдачи» нажимается «Оформить расход». В диалоге указываются: проект — «Срочная кардиохирургическая операция для Максима К.» (баланс проекта: 100 000 руб.); получатель — Кузнецова Елена Сергеевна; сумма расхода — 75 000,00 руб.; статья расхода — «Оплата счета №89/К кардиохирургического центра за диагностику и подготовку к операции»; ответственное лицо — координатор Васильева О.И. Программа проверяет лимит средств (75 000 <= 100 000) и фиксирует операцию списания в таблице FundExpenses.")

    add_p("Сценарий 7. Формирование 5 обязательных аналитических отчетов.", bold=True)
    add_p("В разделе «📊 Аналитические выборки» пользователю доступны 5 регламентных отчетов: 1) «Сумма пожертвований» — расчет совокупных поступлений по проектам за выбранный диапазон дат с подсчетом транзакций и среднего взноса; 2) «Активные проекты» — мониторинг действующих сборов с вычислением процентов выполнения и оставшихся сумм; 3) «Крупнейшие доноры» — рейтинг ТОП-10 благотворителей по общей сумме взносов с датой последнего взноса; 4) «Расходы фонда» — структура списаний по проектам и статьям затрат; 5) «Помощь по категориям» — распределение сумм помощи, числа получателей и долей категорий в совокупном бюджете фонда.")

    add_p("Сценарий 8. Экспорт аналитических данных в формат CSV.", bold=True)
    add_p("Сформировав любой аналитический срез, сотрудник нажимает кнопку «📥 Экспорт в CSV». В окне проводника выбирается путь сохранения (например, Otchet_Kategorii_2026.csv). Сервис ExportService формирует структурированный CSV-файл с кодировкой UTF-8 BOM и корректным экранированием символов, готовый для моментального открытия в Microsoft Excel и построения презентационных диаграмм.")

    add_p("Сценарий 9. Обработка нештатных ситуаций и защита от ошибок оператора.", bold=True)
    add_p("Система предотвращает типичные ошибки оператора: при вводе отрицательной или нулевой суммы кнопка проводки блокируется с подсказкой «Сумма должна быть строго больше нуля»; при попытке списания суммы, превышающей собранный баланс проекта, операция отклоняется сообщением «Превышен доступный баланс проекта»; при временном разрыве связи с SQL Server приложение перехватывает SqlException, сохраняет введенные данные и переводит индикатор в красный цвет.")

    # ==========================================
    # 8. ЗАКЛЮЧЕНИЕ
    # ==========================================
    add_h1("ЗАКЛЮЧЕНИЕ")

    add_p("В ходе прохождения технологической (проектно-технологической) практики был успешно разработан полнофункциональный программный комплекс «Информационная система Благотворительный фонд», предназначенный для комплексной автоматизации операционного, финансового и аналитического учета в некоммерческих благотворительных организациях.")

    add_p("В процессе выполнения индивидуального задания практики были в полном объеме решены все поставленные теоретические и практические задачи:")
    add_p("— проведен глубокий системный анализ предметной области, выделены ключевые бизнес-процессы фонда и формализованы требования к информационному обеспечению;", indent=0, space_after=3)
    add_p("— спроектирована и развернута в СУБД Microsoft SQL Server реляционная база данных в третьей нормальной форме (3NF), включающая шесть взаимосвязанных таблиц с необходимыми ограничениями ссылочной целостности (PRIMARY KEY, FOREIGN KEY), проверочными правилами (CHECK Amount > 0) и индексами;", indent=0, space_after=3)
    add_p("— обоснована и практически реализована многослойная архитектура приложения (Layered Architecture), обеспечивающая четкое разделение ответственности между сущностями предметной области, компонентами доступа к данным, сервисами бизнес-логики и пользовательским интерфейсом;", indent=0, space_after=3)
    add_p("— разработан высокопроизводительный слой доступа к данным на основе чистого низкоуровневого ADO.NET (Microsoft.Data.SqlClient) с использованием пула соединений, стопроцентной параметризации SQL-команд против атак SQL Injection и атомарных транзакций SqlTransaction;", indent=0, space_after=3)
    add_p("— реализованы пять обязательных аналитических выборок, охватывающих все аспекты деятельности фонда: динамику сборов, мониторинг выполнения планов активных проектов, рейтинг благотворителей, сводный баланс и постатейные затраты, а также распределение ресурсов по профильным категориям;", indent=0, space_after=3)
    add_p("— создан модуль экспорта данных в формат CSV с поддержкой стандарта RFC 4180 и кодировки UTF-8 BOM, обеспечивающий интеграцию с популярными офисными табличными процессорами;", indent=0, space_after=3)
    add_p("— разработан интуитивно понятный графический интерфейс пользователя на базе Windows Forms (.NET 8.0) с боковой навигационной панелью, информационными плитками дашборда, фильтрами, поиском и модальными диалогами редактирования с надежной валидацией входных данных;", indent=0, space_after=3)
    add_p("— проведено комплексное модульное и сквозное пользовательское тестирование системы, подтвердившее ее стабильность, отказоустойчивость и полное соответствие функциональным требованиям.", indent=0, space_after=6)

    add_p("Разработанный программный комплекс представляет собой готовое практическое решение, способное повысить прозрачность и оперативность управления благотворительной организацией. В качестве направлений дальнейшего развития проекта планируется реализация веб-кабинета благотворителя на базе ASP.NET Core Blazor, интеграция с онлайн-платежными шлюзами и внедрение модуля автоматической генерации отчетов в формате PDF.")

    # ==========================================
    # 9. СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ
    # ==========================================
    add_h1("СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ")

    lit_sources = [
        "1. Рихтер, Дж. CLR via C#. Программирование на платформе Microsoft .NET Framework 4.5 на языке C# / Дж. Рихтер. — 4-е изд. — СПб.: Питер, 2019. — 896 с.",
        "2. Троелсен, Э., Джепикс, Ф. Язык C# 10 и платформа .NET 6: основные принципы и практики программирования / Э. Троелсен, Ф. Джепикс. — 11-е изд. — М.: Диалектика, 2023. — 1392 с.",
        "3. Фаулер, М. Шаблоны корпоративных приложений: пер. с англ. / М. Фаулер. — М.: Издательский дом «Вильямс», 2018. — 544 с.",
        "4. Мартин, Р. Чистая архитектура. Искусство разработки программного обеспечения / Р. Мартин. — СПб.: Питер, 2021. — 352 с.",
        "5. Дейт, К. Дж. Введение в системы баз данных / К. Дж. Дейт. — 8-е изд. — М.: Издательский дом «Вильямс», 2019. — 1328 с.",
        "6. Кузнецов, С. Д. Основы баз данных: учебное пособие / С. Д. Кузнецов. — 2-е изд. — М.: Интернет-Университет Информационных Технологий; БИНОМ. Лаборатория знаний, 2020. — 488 с.",
        "7. Бен-Ган, И. Microsoft SQL Server 2019. Высокопроизводительный T-SQL / И. Бен-Ган. — СПб.: Питер, 2021. — 576 с.",
        "8. Microsoft. Документация по платформе .NET и Windows Forms [Электронный ресурс]. — Режим доступа: https://learn.microsoft.com/ru-ru/dotnet/desktop/winforms. — Дата доступа: 01.07.2026.",
        "9. Microsoft. Обзор пространства имен Microsoft.Data.SqlClient [Электронный ресурс]. — Режим доступа: https://learn.microsoft.com/ru-ru/sql/connect/ado-net/introduction-microsoft-data-sqlclient-namespace. — Дата доступа: 01.07.2026.",
        "10. Microsoft. Документация по СУБД Microsoft SQL Server [Электронный ресурс]. — Режим доступа: https://learn.microsoft.com/ru-ru/sql/sql-server. — Дата доступа: 01.07.2026.",
        "11. ГОСТ 19.401-78. Единая система программной документации. Текст программы. Требования к содержанию и оформлению. — М.: Стандартинформ, 2010. — 4 с.",
        "12. ГОСТ 19.505-79. Единая система программной документации. Руководство оператора. Требования к содержанию и оформлению. — М.: Стандартинформ, 2010. — 5 с."
    ]

    for lit in lit_sources:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf = p.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(3)
        pf.line_spacing = 1.2
        pf.first_line_indent = Pt(0)
        r = p.add_run(lit)
        r.font.name = 'Times New Roman'
        r.font.size = Pt(14)

    # Save document
    output_filename = "Отчет_Практика_Благотворительный_Фонд.docx"
    output_path = os.path.abspath(output_filename)
    doc.save(output_path)
    print(f"Document initially generated at: {output_path}")

    # Use Word COM to inspect exact heading page numbers and update TOC
    try:
        import win32com.client
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False
        wdoc = word.Documents.Open(output_path)
        
        total_pages = wdoc.ComputeStatistics(2) # wdStatisticPages
        print(f"Word rendered total page count: {total_pages}")

        # Scan paragraphs after page 2 to find headings
        heading_pages = {}
        for p in wdoc.Paragraphs:
            pg = p.Range.Information(3) # wdActiveEndPageNumber
            if pg > 2: # must be after TOC
                raw_text = p.Range.Text.strip()
                if '\t' not in p.Range.Text:
                    for title, key in toc_items:
                        if (raw_text == title or raw_text.startswith(title)) and key not in heading_pages:
                            heading_pages[key] = pg
                            print(f"Matched '{title}' -> Page {pg}")

        wdoc.Close(False)
        word.Quit()

        # Update TOC in python-docx
        if heading_pages:
            print("Updating TOC with exact page numbers...")
            for key, page_num in heading_pages.items():
                if key in toc_paragraphs:
                    p, r_num = toc_paragraphs[key]
                    r_num.text = str(page_num)

            doc.save(output_path)
            print(f"Successfully finalized and saved: {output_path}")

    except Exception as e:
        print(f"Word COM update exception: {e}")

    return output_path

if __name__ == "__main__":
    generate_report()
