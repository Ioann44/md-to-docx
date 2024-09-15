from enum import Enum
import functools
import logging
import os
from pathlib import Path

import markdown2
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt
from docx.styles.style import BaseStyle, ParagraphStyle
from docx.text.parfmt import ParagraphFormat
from tqdm import tqdm
from docx.shared import RGBColor
import docx.oxml.ns


logger = logging.getLogger()
logger.setLevel(logging.DEBUG)


class Style(Enum):
    CODE = "STD_code"
    H1 = "STD_h1"
    H2 = "STD_h2"
    H3 = "STD_h3"
    H4 = "STD_h4"
    IMAGE = "STD_image"
    IMAGE_TEXT = "STD_image_text"
    LIST_M = "STD_list_m"
    LIST_N = "STD_list_n"
    NORMAL = "STD_normal"


def init_docx(docx_file):
    docx_template_path = Path(__file__).parent.joinpath("resources/style_template.docx")
    doc = Document(docx_template_path.resolve().__str__())
    doc._body.clear_content()
    return doc


def md_to_docx(md_file, docx_file):
    with open(md_file, "r", encoding="utf-8") as file:
        md_content = file.read()
    html_content = markdown2.markdown(md_content)
    doc = init_docx("styled_output.docx")

    def add_paragraph(text: str, style: Style, tag_len: int = 2):
        doc.add_paragraph(text[tag_len + 2 : -tag_len - 3], style.value)

    ordered_list = False
    # Разбиваем HTML на блоки для обработки
    for line in tqdm(html_content.splitlines(), ncols=80):
        if line.startswith("<h1>"):
            add_paragraph(line, Style.H1)
        elif line.startswith("<h2>"):
            add_paragraph(line, Style.H2)
        elif line.startswith("<h3>"):
            add_paragraph(line, Style.H3)
        elif line.startswith("<h4>"):
            add_paragraph(line, Style.H4)
        elif line.startswith("<ul>"):
            ordered_list = False
        elif line.startswith("<ol>"):
            ordered_list = True
        elif line.startswith("<li>"):
            add_paragraph(line, Style.LIST_N if ordered_list else Style.LIST_M)
        elif line.startswith("<p>"):
            add_paragraph(line, Style.NORMAL, 1)

    # Сохраняем документ .docx
    doc.save(docx_file)


if __name__ == "__main__":
    # Пример использования
    md_to_docx("data/input.md", "data/output.docx")
