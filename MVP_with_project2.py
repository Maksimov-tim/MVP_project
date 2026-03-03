import easyocr
import re
import os
from typing import List, Dict, Optional, Tuple
from spellchecker import SpellChecker
from dataclasses import dataclass
from PIL import Image, ImageEnhance, ImageFilter
import numpy as np

# Проверка наличия openpyxl для сохранения в Excel
try:
    from openpyxl import Workbook
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("⚠️ Модуль openpyxl не установлен. Установите: pip install openpyxl")
    print("   Программа продолжит работу без сохранения в Excel.\n")

# База известных месторождений (можно расширить или загружать из файла)
KNOWN_FIELDS = [
    'ловинское', 'когалымское', 'самотлорское', 'приобское',
    'федоровское', 'мамонтовское', 'сургутское', 'ромашкинское',
    'васюганское', 'красноленинское', 'салымское', 'таймырское'
]

@dataclass
class ExtractedRecord:
    filename: str
    box_number: Optional[str] = None
    field: Optional[str] = None
    well: Optional[str] = None
    depth_interval: Optional[str] = None
    errors: List[str] = None

    def __post_init__(self):
        if self.errors is None:
            self.errors = []

    def to_dict(self):
        return {
            'Файл': self.filename,
            'Номер ящика': self.box_number or '',
            'Месторождение': self.field or '',
            'Номер скважины': self.well or '',
            'Интервал отбора (м)': self.depth_interval or '',
            'Ошибки': '; '.join(self.errors)
        }


class CoreLabelExtractor:
    def __init__(self, languages: List[str] = ['ru']):
        self.reader = easyocr.Reader(languages)

        # База месторождений (для финальной проверки)
        self.known_fields = KNOWN_FIELDS

        # Словарь для исправления опечаток (только буквенные слова)
        self.keyword_dict = self._build_keyword_dictionary()
        self.spell = SpellChecker(language=None)
        self.spell.word_frequency.load_words(self.keyword_dict)

        # Регулярные выражения для извлечения данных
        self.patterns = {
            'field': [
                r'(?:месторождение|м-ние|площадь|месторожд|место)\s*[:\-]?\s*([а-яА-ЯёЁ\s\-]+)',
                r'([А-Я][а-я]+ское)\b'
            ],
            'well': [
                r'(?:скважина|скв|№скв|скваж)\s*[:\-]?\s*№?\s*(\d{3,5})',
                r'скв\.?\s*№\s*(\d{3,5})',
                r'\b(\d{4,5})\b'
            ],
            'depth': [
                r'(?:интервал|глубина|инт|глуб|отбор)\s*[:\-]?\s*([\d\.,]+\s*[—–-]\s*[\d\.,]+)\s*[мm]',
                r'([\d\.,]+\s*[—–-]\s*[\d\.,]+)\s*(?:м|m)',
                r'\b(\d+[.,]?\d*\s*[—–-]\s*\d+[.,]?\d*)\b'
            ],
            'box': [
                r'(?:коробка|ящик|кор)\s*[:\-]?\s*№?\s*(\d+)',
                r'№\s*(\d{2,3})\s*$',
                r'\b(\d{2,3})\b(?!\s*[мm])'
            ]
        }

    def _build_keyword_dictionary(self) -> List[str]:
        """Создаёт список правильных слов (без дефисов)"""
        keywords = [
            'месторождение', 'площадь', 'месторожд', 'место',
            'скважина', 'скв', 'скваж', '№скв',
            'интервал', 'глубина', 'инт', 'глуб', 'отбор',
            'коробка', 'ящик', 'кор',
            'метры', 'метр', 'м',
            'мние',  # возможная замена для "м-ние"
        ]
        # Добавляем названия месторождений из глобальной константы
        keywords.extend(KNOWN_FIELDS)
        return [kw.lower() for kw in keywords]

    def preprocess_image(self, image_path: str) -> np.ndarray:
        """Улучшает качество изображения перед распознаванием (ваша функция)"""
        img = Image.open(image_path)

        if img.mode == 'RGBA':
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[-1])
            img = background
        elif img.mode != 'RGB':
            img = img.convert('RGB')

        img = img.convert('L')
        enhancer = ImageEnhance.Contrast(img)
        img = enhancer.enhance(2.0)
        img = img.filter(ImageFilter.SHARPEN)
        img = img.convert('RGB')

        return np.array(img)

    def correct_words(self, text: str) -> str:
        """Исправляет опечатки во всех словах, используя словарь"""
        tokens = re.findall(r'[а-яА-ЯёЁ]+|[^а-яА-ЯёЁ]+', text)
        corrected_tokens = []
        for token in tokens:
            if re.match(r'^[а-яА-ЯёЁ]+$', token):
                try:
                    corrected = self.spell.correction(token)
                    corrected_tokens.append(corrected if corrected else token)
                except:
                    corrected_tokens.append(token)
            else:
                corrected_tokens.append(token)
        return ''.join(corrected_tokens)

    def match_field_name(self, candidate: str) -> Optional[str]:
        """Сопоставляет название месторождения с базой (после исправления)"""
        candidate = candidate.strip().lower()
        if candidate in self.known_fields:
            return candidate.capitalize()
        # Дополнительная проверка на вхождение (если исправление не помогло)
        for known in self.known_fields:
            if known in candidate or candidate in known:
                return known.capitalize()
        return None

    def extract_data(self, corrected_text: str) -> Dict[str, Optional[str]]:
        """Извлекает поля из исправленного текста"""
        data = {'field': None, 'well': None, 'depth': None, 'box': None}
        
        # 1. Пытаемся извлечь по регулярным выражениям
        for key, patterns in self.patterns.items():
            for pattern in patterns:
                match = re.search(pattern, corrected_text, re.IGNORECASE)
                if match:
                    value = match.group(1).strip()
                    value = re.sub(r'[^\w\s\.,\-]', '', value)
                    data[key] = value
                    break
        
        # 2. Если месторождение не найдено, попробуем найти прямое вхождение известного названия
        if not data['field']:
            text_lower = corrected_text.lower()
            for known in self.known_fields:
                if known in text_lower:
                    data['field'] = known.capitalize()
                    break
        
        return data

    def validate_data(self, data: Dict[str, Optional[str]]) -> Tuple[Dict[str, Optional[str]], List[str]]:
        """Проверяет и нормализует данные"""
        validated = {}
        errors = []

        # Месторождение
        field_val = data.get('field')
        if field_val:
            matched = self.match_field_name(field_val)
            if matched:
                validated['field'] = matched
            else:
                validated['field'] = None
                errors.append(f"Месторождение '{field_val}' не найдено в базе")
        else:
            validated['field'] = None
            errors.append("Месторождение не найдено")

        # Скважина
        well_val = data.get('well')
        if well_val:
            digits = re.findall(r'\d+', well_val)
            if digits:
                best = max(digits, key=len)
                if 3 <= len(best) <= 5:
                    validated['well'] = best
                else:
                    validated['well'] = None
                    errors.append(f"Номер скважины '{best}' вне диапазона (3-5 цифр)")
            else:
                validated['well'] = None
                errors.append("Номер скважины не содержит цифр")
        else:
            validated['well'] = None
            errors.append("Скважина не найдена")

        # Глубина
        depth_val = data.get('depth')
        if depth_val:
            cleaned = re.sub(r'\s+', '', depth_val).replace(',', '.')
            match = re.search(r'(\d+[.]?\d*)[—–-](\d+[.]?\d*)', cleaned)
            if match:
                validated['depth'] = f"{match.group(1)}-{match.group(2)}"
            else:
                validated['depth'] = None
                errors.append(f"Интервал глубины '{depth_val}' не соответствует формату")
        else:
            validated['depth'] = None
            errors.append("Интервал глубины не найден")

        # Ящик
        box_val = data.get('box')
        if box_val:
            digits = re.findall(r'\d+', box_val)
            if digits:
                for d in reversed(digits):
                    if 2 <= len(d) <= 3:
                        validated['box'] = d
                        break
                else:
                    validated['box'] = None
                    errors.append(f"Номер ящика '{box_val}' не содержит 2-3-значного числа")
            else:
                validated['box'] = None
                errors.append("Номер ящика не содержит цифр")
        else:
            validated['box'] = None
            errors.append("Номер ящика не найден")

        return validated, errors

    def process_image(self, image_path: str) -> List[ExtractedRecord]:
        filename = os.path.basename(image_path)
        records = []

        try:
            processed_img = self.preprocess_image(image_path)
            ocr_results = self.reader.readtext(processed_img, detail=1, paragraph=False)

            if not ocr_results:
                ocr_results = self.reader.readtext(image_path, detail=1, paragraph=False)

            if not ocr_results:
                err = ExtractedRecord(filename=filename)
                err.errors.append("Текст не распознан")
                return [err]

            all_text = ' '.join([text for _, text, _ in ocr_results])

            # Исправляем опечатки во всех словах
            corrected_text = self.correct_words(all_text)

            # Извлекаем данные
            extracted = self.extract_data(corrected_text)
            validated, errors = self.validate_data(extracted)

            record = ExtractedRecord(filename=filename)
            record.field = validated.get('field')
            record.well = validated.get('well')
            record.depth_interval = validated.get('depth')
            record.box_number = validated.get('box')
            record.errors = errors
            records.append(record)

        except Exception as e:
            err = ExtractedRecord(filename=filename)
            err.errors.append(f"Ошибка обработки: {str(e)}")
            records.append(err)

        return records

    def process_folder(self, folder: str = ".") -> List[ExtractedRecord]:
        images = [f for f in os.listdir(folder) 
                  if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp'))]
        all_records = []
        for img in images:
            print(f"Обработка: {img}")
            records = self.process_image(os.path.join(folder, img))
            all_records.extend(records)
        return all_records


def save_to_excel(records: List[ExtractedRecord], out_file: str = "core_labels.xlsx"):
    """Сохраняет данные в Excel с автоматической шириной столбцов"""
    if not OPENPYXL_AVAILABLE:
        print("⚠️ openpyxl не установлен. Excel-файл не создан.")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "Бирки керна"

    headers = ['Файл', 'Номер ящика', 'Месторождение', 'Номер скважины', 'Интервал отбора (м)', 'Ошибки']
    ws.append(headers)

    for r in records:
        ws.append([
            r.filename,
            r.box_number or '',
            r.field or '',
            r.well or '',
            r.depth_interval or '',
            '; '.join(r.errors)
        ])

    # Автоматическая ширина столбцов
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 80)
        ws.column_dimensions[col_letter].width = adjusted_width

    wb.save(out_file)
    print(f"✅ Excel сохранён: {out_file}")


def print_statistics(records: List[ExtractedRecord]):
    total = len(records)
    fields = sum(1 for r in records if r.field)
    wells = sum(1 for r in records if r.well)
    depths = sum(1 for r in records if r.depth_interval)
    boxes = sum(1 for r in records if r.box_number)
    no_errors = sum(1 for r in records if not r.errors)

    print("\n" + "="*60)
    print("СТАТИСТИКА")
    print("="*60)
    print(f"Всего записей: {total}")
    print(f"Месторождений: {fields} ({fields/total*100:.1f}%)")
    print(f"Скважин: {wells} ({wells/total*100:.1f}%)")
    print(f"Интервалов: {depths} ({depths/total*100:.1f}%)")
    print(f"Ящиков: {boxes} ({boxes/total*100:.1f}%)")
    print(f"Без ошибок: {no_errors} ({no_errors/total*100:.1f}%)")


def main():
    print("="*60)
    print("РАСПОЗНАВАНИЕ БИРОК КЕРНА")
    print("="*60)

    extractor = CoreLabelExtractor(languages=['ru'])
    records = extractor.process_folder(".")

    if records:
        save_to_excel(records, "core_labels.xlsx")
        print_statistics(records)

        print("\nПервые 5 записей:")
        for i, r in enumerate(records[:5]):
            print(f"{i+1}. {r.filename}: ящ.{r.box_number}, м.{r.field}, скв.{r.well}, {r.depth_interval}")
            if r.errors:
                print(f"   Ошибки: {', '.join(r.errors)}")
    else:
        print("Нет изображений для обработки")


if __name__ == "__main__":
    main()