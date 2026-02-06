import pdfplumber
import pandas as pd
import re
from pathlib import Path

PDF_PATH = "input.pdf"
MD_PATH = "intermediate.md"
XLS_PATH = "result.xlsx"


def pdf_to_md(pdf_path, md_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n\n".join(
            page.extract_text()
            for page in pdf.pages
            if page.extract_text()
        )
    Path(md_path).write_text(text, encoding="utf-8")


def md_to_xls(md_path, xls_path):
    text = Path(md_path).read_text(encoding="utf-8")

    part = None
    q_id = None
    q_lines = []
    figures = []
    rows = []

    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue

        m = re.search(r'ЧАСТЬ\s+(\d+)', line)
        if m:
            if q_id:
                rows.append((part, q_id, " ".join(q_lines), "; ".join(figures)))
            part = int(m.group(1))
            q_id = None
            q_lines, figures = [], []
            continue

        m = re.match(r'^([A-ZА-Я]\d+)\s+(.*)', line)
        if m:
            if q_id:
                rows.append((part, q_id, " ".join(q_lines), "; ".join(figures)))
            q_id = m.group(1)
            q_lines = [m.group(2)]
            figures = []
            continue

        if line.startswith("![](") or line.startswith("Рис"):
            figures.append(line)
            continue

        if q_id:
            q_lines.append(line)

    if q_id:
        rows.append((part, q_id, " ".join(q_lines), "; ".join(figures)))

    df = pd.DataFrame(
        rows,
        columns=["Часть", "Номер вопроса", "Вопрос", "Рисунок"]
    )
    df.to_excel(xls_path, index=False)


if __name__ == "__main__":
    if not Path(MD_PATH).exists():
        pdf_to_md(PDF_PATH, MD_PATH)
    md_to_xls(MD_PATH, XLS_PATH)
    print("Готово")
