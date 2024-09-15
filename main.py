import markdown2
from docx import Document
from tqdm import tqdm


def md_to_docx(md_file, docx_file):
    # Открываем .md файл и читаем его содержимое
    with open(md_file, "r", encoding="utf-8") as file:
        md_content = file.read()

    # Преобразуем содержимое .md в HTML (это промежуточный этап)
    html_content = markdown2.markdown(md_content)

    # Создаем новый документ .docx
    doc = Document()

    # Разбиваем HTML на блоки для обработки
    in_list = False
    for line in tqdm(html_content.splitlines(), ncols=80):
        if line.startswith("<h1>"):
            doc.add_heading(line[4:-5], level=1)
        elif line.startswith("<h2>"):
            doc.add_heading(line[4:-5], level=2)
        elif line.startswith("<h3>"):
            doc.add_heading(line[4:-5], level=3)
        elif line.startswith("<ul>") or in_list:
            # Если начинается маркированный список
            in_list = True
            while in_list:
                line = line.strip()
                if line.startswith("<li>") or line.startswith("<ul>"):
                    doc.add_paragraph(line[4:-5], style="ListBullet")
                if "</ul>" in line or "</li>" in line:
                    in_list = False
                else:
                    break
        elif line.startswith("<p>"):
            # Добавляем новый абзац
            doc.add_paragraph(line[3:-4])

    # Сохраняем документ .docx
    doc.save(docx_file)


if __name__ == "__main__":
    # Пример использования
    md_to_docx("input.md", "output.docx")
