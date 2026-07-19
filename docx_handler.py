import os
import json
import re
import copy
import logging
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_LINE_SPACING, WD_COLOR_INDEX
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn

# Налаштування логера (щоб не тягнути зовнішній logger_config, використовуємо стандартний)
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("docx_handler")


class DocxHandler:
    def __init__(self, client_data: dict, const_path: str = "const.json"):
        # Тепер ми отримуємо дані клієнта напряму з Flask (без читання з диска)
        self.client_data = client_data

        # Константу можемо залишити з файлу, якщо він є
        self.const_path = const_path
        self.const_data = self._load_json(const_path) if os.path.exists(const_path) else {"web": "creditkasa.com.ua"}

    def _load_json(self, path: str) -> dict:
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"Не вдалося завантажити {path}: {e}. Використовуються стандартні налаштування.")
            return {}

    def _clean_and_paint_black(self, doc):
        header_pattern = re.compile(r"^[IІVХX]+\.\s+[А-ЯЩЬЮЯЄІЇҐA-Z]", re.IGNORECASE)

        for paragraph in doc.paragraphs:
            p_text = paragraph.text.strip()
            is_header = bool(header_pattern.match(p_text))

            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0, 0, 0)
                if is_header:
                    run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
                else:
                    run.font.highlight_color = None

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        p_text = paragraph.text.strip()
                        is_header = bool(header_pattern.match(p_text))

                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(0, 0, 0)
                            if is_header:
                                run.font.highlight_color = WD_COLOR_INDEX.GRAY_25
                            else:
                                run.font.highlight_color = None

    def _format_rate_comma(self, rate_str: str) -> str:
        try:
            val = float(rate_str.replace(',', '.'))
            return f"{val:.2f}".replace('.', ',') + " %"
        except ValueError:
            return f"{rate_str} %"

    def _convert_rate_to_words(self, rate_str: str) -> tuple:
        try:
            val = float(rate_str.replace(",", "."))
            if val == 1.0 or val == 1:
                return "1.00%", "одна ціла, нуль сотих процентів"
            elif val == 0.74:
                return "0.74%", "нуль цілих, сімдесят чотири сотих процентів"
            elif val == 1.5:
                return "1.50%", "одна ціла, п'ятдесят сотих процентів"

            parts = f"{val:.2f}".split('.')
            whole, frac = int(parts[0]), int(parts[1])
            num_words = {0: "нуль", 1: "одна", 2: "дві", 3: "три", 4: "чотири", 5: "п'ять", 6: "шість", 7: "сім",
                         8: "вісім", 9: "дев'ять"}
            frac_words = {0: "нуль", 50: "п'ятдесят", 74: "сімдесят чотири"}

            w_txt = num_words.get(whole, str(whole))
            f_txt = frac_words.get(frac, str(frac))
            w_suf = "ціла" if whole == 1 else "цілих"

            return f"{val:.2f}%", f"{w_txt} {w_suf}, {f_txt} сотих процентів"
        except Exception:
            return f"{rate_str}%", "вказана кількість процентів"

    def _replace_in_p(self, paragraph, pattern_str, replacement):
        pattern = re.compile(pattern_str, re.IGNORECASE)
        replaced = False

        for run in paragraph.runs:
            if pattern.search(run.text):
                run.text = pattern.sub(replacement, run.text)
                replaced = True

        if not replaced:
            full_txt = "".join(r.text for r in paragraph.runs)
            if pattern.search(full_txt):
                fmts = [(r.font.name, r.font.size, r.bold, r.italic) for r in paragraph.runs]
                paragraph.text = pattern.sub(replacement, full_txt)
                if paragraph.runs and fmts:
                    paragraph.runs[0].font.name = fmts[0][0]
                    paragraph.runs[0].font.size = fmts[0][1]
                    paragraph.runs[0].bold = fmts[0][2]
                    paragraph.runs[0].italic = fmts[0][3]

    def _remove_p(self, paragraph):
        element = paragraph._element
        if element.getparent() is not None:
            element.getparent().remove(element)

    def _insert_cloned_p_before(self, target_p, text, clone_from=None):
        ref_p = clone_from if clone_from else target_p
        new_p_element = copy.deepcopy(ref_p._element)
        target_p._element.addprevious(new_p_element)
        new_p = Paragraph(new_p_element, target_p._parent)

        for r in new_p.runs:
            r.text = ""

        match = re.match(r"^(\s*[-*•]\s*)", ref_p.text)
        prefix = match.group(1) if match else ""

        final_text = f"{prefix}{text}"

        if new_p.runs:
            new_p.runs[0].text = final_text
            new_p.runs[0].bold = False
            new_p.runs[0].italic = False
        else:
            r = new_p.add_run(final_text)
            r.bold = False
            r.italic = False
        return new_p

    def _insert_p_aligned_no_bullet(self, p_ref, text):
        new_p_element = copy.deepcopy(p_ref._element)
        p_ref._element.addnext(new_p_element)
        new_p = Paragraph(new_p_element, p_ref._parent)

        pPr = new_p._p.get_or_add_pPr()

        if pPr.numPr is not None:
            pPr.remove(pPr.numPr)

        ind = pPr.find(qn('w:ind'))
        if ind is not None:
            for attr in ['hanging', 'firstLine']:
                if qn(f'w:{attr}') in ind.attrib:
                    del ind.attrib[qn(f'w:{attr}')]
        else:
            new_p.paragraph_format.left_indent = p_ref.paragraph_format.left_indent
            new_p.paragraph_format.first_line_indent = Pt(0)

        new_p.paragraph_format.space_before = Pt(0)
        new_p.paragraph_format.space_after = Pt(0)
        new_p.paragraph_format.line_spacing = 1.0
        new_p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE

        for r in new_p.runs:
            r.text = ""

        if new_p.runs:
            new_p.runs[0].text = text
            new_p.runs[0].bold = False
            new_p.runs[0].italic = False
        else:
            r = new_p.add_run(text)
            r.bold = False
            r.italic = False

        return new_p

    def process_file(self, input_path: str, output_path: str) -> str:
        logger.info(f"Початок обробки DOCX файлу: {input_path}")

        if not os.path.exists(input_path):
            raise FileNotFoundError(f"Вхідний DOCX не знайдено: {input_path}")

        doc = Document(input_path)
        web_domain = self.const_data.get("web", "creditkasa.com.ua")
        otp = self.client_data.get("otp", "")

        for paragraph in doc.paragraphs:
            self._replace_in_p(paragraph, r"creditkasa\.(com\.)?ua", web_domain)

        if otp:
            for paragraph in doc.paragraphs:
                if "на виконання зазначених вимог" in paragraph.text.lower() or "ідентифікатор" in paragraph.text.lower():
                    self._replace_in_p(paragraph, r"(ідентифікатор\s+)[A-Za-zА-Яа-я0-9іІїЇєЄґҐ]+", rf"\g<1>{otp}")

        fee_val = str(self.client_data.get("fee", "0"))
        std_val = str(self.client_data.get("standard_rate", "0"))
        red_val = str(self.client_data.get("reduced_rate", "0"))
        promo_val = str(self.client_data.get("promo_rate", "0"))
        pref_val = str(self.client_data.get("preferential_rate", "0"))

        start_idx, end_idx = -1, -1
        for i, p in enumerate(doc.paragraphs):
            txt = p.text.lower()
            if "на наступних умовах:" in txt:
                start_idx = i + 1
            elif start_idx != -1 and "* базовий період" in txt and "проміжки часу" in txt:
                end_idx = i
                break

        if start_idx != -1 and end_idx != -1:
            cond_paragraphs = doc.paragraphs[start_idx:end_idx]

            p_fee = p_red = p_promo = p_pref = p_std_title = None
            p_std_subs = []

            for p in cond_paragraphs:
                txt = p.text.lower()
                if p_std_title is not None:
                    p_std_subs.append(p)
                elif "комісія за видачу" in txt:
                    p_fee = p
                elif "промо-ставка" in txt or "(промо-ставка)" in txt:
                    p_promo = p
                elif "пільгова" in txt:
                    p_pref = p
                elif "знижена" in txt:
                    p_red = p
                elif "стандартна % ставка" in txt:
                    p_std_title = p

            p_bullet_ref = p_fee if p_fee else (p_red if p_red else None)

            if fee_val and fee_val not in ["0", ""]:
                if p_fee: self._replace_in_p(p_fee, r"\d+(?:[.,]\d+)?\s*%", self._format_rate_comma(fee_val))
            elif p_fee:
                self._remove_p(p_fee)

            if red_val and red_val not in ["0", ""]:
                if p_red: self._replace_in_p(p_red, r"\d+(?:[.,]\d+)?\s*%", self._format_rate_comma(red_val))
            elif p_red:
                self._remove_p(p_red)

            if promo_val and promo_val not in ["0", ""]:
                if p_promo:
                    self._replace_in_p(p_promo, r"\d+(?:[.,]\d+)?\s*%", self._format_rate_comma(promo_val))
                elif p_std_title and p_bullet_ref:
                    self._insert_cloned_p_before(p_std_title,
                                                 f"знижена % ставка (Промо-ставка) – {self._format_rate_comma(promo_val)} в день;",
                                                 clone_from=p_bullet_ref)
            elif p_promo:
                self._remove_p(p_promo)

            if pref_val and pref_val not in ["0", ""]:
                if p_pref:
                    self._replace_in_p(p_pref, r"\d+(?:[.,]\d+)?\s*%", self._format_rate_comma(pref_val))
                elif p_std_title and p_bullet_ref:
                    self._insert_cloned_p_before(p_std_title,
                                                 f"пільгова % ставка – {self._format_rate_comma(pref_val)} в день;",
                                                 clone_from=p_bullet_ref)
            elif p_pref:
                self._remove_p(p_pref)

            if p_std_title and p_bullet_ref:
                new_std_title = self._insert_cloned_p_before(p_std_title, "стандартна % ставка:",
                                                             clone_from=p_bullet_ref)

                self._remove_p(p_std_title)

                for p in p_std_subs:
                    self._remove_p(p)

                std_pct, std_words = self._convert_rate_to_words(std_val)
                txt1 = f"*- {std_pct} ({std_words}) за кожен день користування Кредитом, яка застосовується протягом перших 180 (ста вісімдесяти) календарних днів з дати укладення цього Договору. В цей період можливе використання Позичальником права користування Кредитом за Промо-ставкою та/або Зниженою та/або Пільговою процентною ставкою."
                txt2 = "*-0.74% (нуль цілих, сімдесят чотири сотих процентів) за кожен день користування Кредитом, яка застосовується у період починаючи з 181 (ста вісімдесят першого) календарного дня дії Договору і до закінчення строку дії цього Договору або до дати фактичного повернення всієї суми Кредиту (до тієї із зазначених дат, яка настане раніше)."

                self._insert_p_aligned_no_bullet(new_std_title, txt2)
                self._insert_p_aligned_no_bullet(new_std_title, txt1)

        self._clean_and_paint_black(doc)

        doc.save(output_path)
        logger.info(f"Оброблений файл збережено: {output_path}")
        return output_path


# =====================================================================
# Це функція-обгортка, яку викликає Flask (app.py)
# Вона бере словник з даними з фронтенду, створює твій ідеальний клас
# і запускає обробку.
# =====================================================================
def process_document(input_path: str, output_path: str, data: dict):
    # Якщо потрібно, щоб const.json підтягувався, просто поклади його поруч з app.py
    handler = DocxHandler(client_data=data, const_path="const.json")
    handler.process_file(input_path, output_path)