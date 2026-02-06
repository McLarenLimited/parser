import pdfplumber
import pandas as pd
import re
from pathlib import Path
import tempfile
from typing import Optional, List, Dict, Any
import logging


# Настройка логирования

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
logger = logging.getLogger(__name__)

PDF_PATH = "input.pdf"
MD_PATH = "intermediate.md"
XLS_PATH = "result.xlsx"



# Парсер вопросов

class QuestionParser:
    """
    Парсер вопросов из MD / текстового представления
    """

    def __init__(self) -> None:
        self.part_pattern = re.compile(r'ЧАСТЬ\s+(\d+)', re.IGNORECASE)
        self.question_pattern = re.compile(
            r'^([A-ZА-Я]{1}\d{1,3})[\s\.\)]*\s+(.*)'
        )
        self.figure_pattern = re.compile(r'^Рис\.?\s*.+', re.IGNORECASE)

        self.reset()

    def reset(self) -> None:
        """Сброс состояния (важно для повторного использования)"""
        self.current_part: Optional[int] = None
        self.current_question_id: Optional[str] = None
        self.current_question_lines: List[str] = []
        self.current_figures: List[str] = []
        self.questions: List[Dict[str, Any]] = []

    def flush_question(self) -> None:
        """Сохранить текущий вопрос"""
        if not self.current_question_id:
            return

        question_text = " ".join(self.current_question_lines).strip()
        figures = "; ".join(self.current_figures) if self.current_figures else None

        self.questions.append({
            "Часть": self.current_part,
            "Номер вопроса": self.current_question_id,
            "Вопрос": question_text,
            "Рисунок": figures
        })

        self.current_question_id = None
        self.current_question_lines = []
        self.current_figures = []

    def parse_line(self, line: str) -> None:
        line = line.strip()
        if not line:
            return

        # Игнорируем LaTeX-заголовки
        if line.startswith("\\title") or line.startswith("\\section"):
            return

        # Определение части
        part_match = self.part_pattern.search(line)
        if part_match:
            self.flush_question()
            self.current_part = int(part_match.group(1))
            logger.debug(f"Найдена часть: {self.current_part}")
            return

        # Начало нового вопроса
        q_match = self.question_pattern.match(line)
        if q_match:
            self.flush_question()
            self.current_question_id = q_match.group(1)
            self.current_question_lines.append(q_match.group(2))
            return

        # Картинка
        if line.startswith("![]("):
            if self.current_question_id:
                self.current_figures.append(line)
            return

        # Подпись к рисунку
        if self.figure_pattern.match(line):
            if self.current_question_id:
                self.current_figures.append(line)
            return

        # Продолжение текста вопроса
        if self.current_question_id:
            self.current_question_lines.append(line)

    def parse_text(self, text: str) -> List[Dict[str, Any]]:
        self.reset()

        for line in text.splitlines():
            self.parse_line(line)

        self.flush_question()
        return self.questions



# PDF to MD
def pdf_to_md(pdf_path: str, md_path: str) -> None:
    pdf_path = Path(pdf_path)
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF файл не найден: {pdf_path}")

    pages_text: List[str] = []

    try:
        with pdfplumber.open(pdf_path) as pdf:
            for i, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text:
                    pages_text.append(f"\n\n## PAGE {i}\n\n{text}")
                logger.info(f"Обработана страница {i}")
    except Exception as e:
        raise RuntimeError(f"Ошибка чтения PDF: {e}")

    with tempfile.NamedTemporaryFile(
        mode="w",
        encoding="utf-8",
        delete=False,
        suffix=".md"
    ) as tmp:
        tmp.write("\n".join(pages_text))
        temp_path = tmp.name

    Path(temp_path).replace(md_path)
    logger.info(f"MD файл сохранён: {md_path}")


# MD to XLS
def md_to_xls(md_path: str, xls_path: str) -> pd.DataFrame:
    md_path = Path(md_path)
    if not md_path.exists():
        raise FileNotFoundError(f"MD файл не найден: {md_path}")

    # Чтение с попытками разных кодировок
    try:
        text = md_path.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        for enc in ("cp1251", "iso-8859-1"):
            try:
                text = md_path.read_text(encoding=enc)
                break
            except UnicodeDecodeError:
                continue
        else:
            raise ValueError("Не удалось определить кодировку MD файла")

    parser = QuestionParser()
    questions = parser.parse_text(text)

    if not questions:
        logger.warning("Вопросы не найдены")

    df = pd.DataFrame(questions)

    with tempfile.NamedTemporaryFile(
        suffix=".xlsx",
        delete=False
    ) as tmp:
        df.to_excel(tmp.name, index=False)
        Path(tmp.name).replace(xls_path)

    logger.info(f"Excel файл сохранён: {xls_path}")
    return df



# Тоxка входа
def main() -> None:
    try:
        if not Path(MD_PATH).exists():
            logger.info("MD не найден — конвертация PDF → MD")
            pdf_to_md(PDF_PATH, MD_PATH)
        else:
            logger.info("MD файл уже существует")

        logger.info("Парсинг MD → XLS")
        df = md_to_xls(MD_PATH, XLS_PATH)

        logger.info(f"Обработано вопросов: {len(df)}")
        if not df.empty:
            logger.info(f"Части: {sorted(df['Часть'].dropna().unique())}")

        print("Успешно")

    except FileNotFoundError as e:
        logger.error(e)
        print("Файл не найден")
    except Exception as e:
        logger.exception("Критическая ошибка")
        print(f"Ошибка: {e}")


if __name__ == "__main__":
    main()
