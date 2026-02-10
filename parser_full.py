import pandas as pd
import re
from pathlib import Path
from pdf2image import convert_from_path
import pytesseract

PDF_PATH = "input.pdf"
MD_PATH = "intermediate.md"
XLS_PATH = "result.xlsx"


def pdf_to_md(pdf_path, md_path):
    import pytesseract
    from pdf2image import convert_from_path
    from pathlib import Path

# Необходима программа Tesseract-OCR. Прописать путь к .exe файлу.
    pytesseract.pytesseract.tesseract_cmd = (
        r"C:\Program Files\Tesseract-OCR\tesseract.exe"
    )

# Необходимо скачать программу poppler и прописать путь к её каталогам.
    images = convert_from_path(
        pdf_path,
        dpi=300,
        poppler_path=r"C:\poppler\Library\bin"
    )

    pages_text = []
    for img in images:
        text = pytesseract.image_to_string(img, lang="rus")
        if text.strip():
            pages_text.append(text)

    Path(md_path).write_text(
        "\n\n".join(pages_text),
        encoding="utf-8"
    )



def md_to_xls(md_path, xls_path):
    text = Path(md_path).read_text(encoding="utf-8")

    part = None
    current_letter = None
    q_num = 0
    q_id = None
    q_lines = []
    rows = []

    def flush():
        nonlocal q_id, q_lines
        if q_id and q_lines:
            rows.append((part, q_id, " ".join(q_lines)))
        q_lines = []

    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue

        m = re.search(r'часть\s*(\d+)', line, re.IGNORECASE)
        if m:
            flush()
            part = int(m.group(1))
            current_letter = None
            q_num = 0
            q_id = None
            continue

        m = re.search(r'\[\s*([всbc])\s*([0-9]+)\s*\|', line, re.IGNORECASE)
        if m:
            flush()
            letter = m.group(1).upper().replace("B", "В").replace("C", "С")
            q_num = int(m.group(2))
            current_letter = letter
            q_id = f"{letter}{q_num}"

            after = line.split("|", 1)
            q_lines = [after[1].strip()] if len(after) > 1 else []
            continue

        if (
            q_id is None
            or line.startswith("*")
            or (
                current_letter
                and re.match(r'^[А-ЯA-Z]', line)
                and not line.lower().startswith("рис")
            )
        ):
            flush()
            if current_letter is None:
                current_letter = "В" if part == 2 else "С"
            q_num += 1
            q_id = f"{current_letter}{q_num}"
            q_lines = [line.lstrip("* ").strip()]
            continue

        q_lines.append(line)

    flush()

    print("ROWS FOUND:", len(rows))

    df = pd.DataFrame(
        rows,
        columns=["Часть", "Номер вопроса", "Вопрос"]
    )
    df.to_excel(xls_path, index=False)




#MAIN
if __name__ == "__main__":
    pdf_to_md(PDF_PATH, MD_PATH)
    md_to_xls(MD_PATH, XLS_PATH)
    print("Готово")

