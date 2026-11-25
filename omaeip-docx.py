from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import subprocess
import os


def clear_screen():
    subprocess.run("cls" if os.name == "nt" else "clear", shell=True)


def create_report():
    document_name = "OmAEiP"
    post = "ДОЛЖНОСТЬ ПРЕПОДАВАТЕЛЯ"
    teacher_name = "ФИО ПРЕПОДАВАТЕЛЯ"
    report_title = "ТЕМА"
    text = "ТЕКСТ"

    print("----- Докладная -----\n")

    document_name = input("Введите название файла: ")
    post = input("Введите должность преподавателя: ")
    teacher_name = input("Введите ФИО преподавателя: ")
    report_title = input("Введите тему докладной: ")
    text = input("Введите текст: ")

    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(14)
    style.paragraph_format.line_spacing = 1.5

    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.5)

    header_text = [
        "Декану факультета",
        "Информационных технологий и Права",
        "Князевой О.Н.",
        "от специалиста деканата факультета",
        "Информационных технологий и Права",
        f"{teacher_name}.",
    ]

    for line in header_text:
        p = doc.add_paragraph(line)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(0)

    doc.add_paragraph()

    p_title = doc.add_paragraph("ДОКЛАДНАЯ ЗАПИСКА")
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.runs[0].bold = True
    p_title.paragraph_format.space_after = Pt(0)

    p_subtitle = doc.add_paragraph(f"{report_title}")
    p_subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_subtitle.paragraph_format.space_after = Pt(0)
    doc.add_paragraph()

    def add_body_paragraph(text):
        p = doc.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.25)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.line_spacing = 1.5
        return p

    add_body_paragraph(f"{text}")

    doc.add_paragraph()

    p_sign = doc.add_paragraph()
    p_sign.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sign.paragraph_format.left_indent = Cm(1.25)
    p_sign.paragraph_format.space_after = Pt(0)

    p_sign.add_run(f"{post} _________________________ {teacher_name}.")

    doc.save(f"{document_name}.docx")
    print(f"Документ {document_name}.docx успешно создан.")


def ShowMenu():
    print("----- Список операций -----\n")
    print("1. Докладная")
    print("2. Генерация шаблона")
    print("\n0. Выход\n")
    print("---------------------------\n")


while True:
    ShowMenu()
    operation = int(input("Введите операцию: "))
    if operation == 1:
        clear_screen()
        create_report()
        break
    if operation == 2:
        create_report()
        break
    if operation == 0:
        break
