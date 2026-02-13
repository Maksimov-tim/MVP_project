import easyocr
from PIL import Image, ImageEnhance, ImageFilter
import numpy as np
from spellchecker import SpellChecker

simvols = ["(",")","[","]",":"]
spell = SpellChecker(language="ru")
spell.word_frequency.load_words(["Месторождение","Коробка"])
inf = ""

def preprocess_image(image_path):
    """Улучшает качество изображения перед распознаванием"""
    img = Image.open(image_path)
    
    # Удаляем альфа-канал, если он есть (для PNG изображений)
    if img.mode == 'RGBA':
        # Создаем белый фон
        background = Image.new('RGB', img.size, (255, 255, 255))
        # Накладываем изображение на белый фон
        background.paste(img, mask=img.split()[-1])  # используем альфа-канал как маску
        img = background
    elif img.mode != 'RGB':
        img = img.convert('RGB')
    
    # Конвертируем в оттенки серого
    img = img.convert('L')
    
    # Увеличение контраста (теперь img в режиме 'L', что поддерживается)
    enhancer = ImageEnhance.Contrast(img)
    img = enhancer.enhance(2.0)
    
    # Увеличение резкости
    img = img.filter(ImageFilter.SHARPEN)
    
    # Конвертируем обратно в RGB для EasyOCR (хотя EasyOCR сам конвертирует в grayscale)
    img = img.convert('RGB')
    
    return np.array(img)

# Создаем объект ридера, указываем язык (русский и английский)
reader = easyocr.Reader(['ru'])

# Путь к изображению
image_path = '/home/toor/Документы/Максимов Тимофей/birka/test_2_1.png'

# Читаем текст с изображения
result = reader.readtext(preprocess_image(image_path))

# Выводим результат
for detection in result:
    print(detection[1]) # detection[1] - это распознанный текст. [0 - координаты (?) прямоугольника с текстом, 1 - текст, 2 - степень увренности]
    buff  = detection[1].split(" ")
    for word in buff:
        for simvol in simvols:
            if word is None:
                continue
            
            word = word.replace(simvol, "")

        if len(spell.unknown([str(word)])) <= 0:
            result_word = spell.correction(word)
        else:
            result_word = None

        if result_word is not None:
            inf += result_word
            print(f"Пробразованное слово: {inf}")
            
        else:
            print(f"Не удалось распознать {word}")

        print(f"Список: {buff}")

    inf = ""