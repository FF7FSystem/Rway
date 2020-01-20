from requests import get
import json
from PIL import Image
import pytesseract
import base64
from io import BytesIO
import re
import time

def benchmark(func):
    def wrapper(*args, **kwargs):
        t = time.time()
        res = func(*args, **kwargs)
        print(f'Функция {func.__name__} выполнилась за время {time.time() - t:.2f}s')
        return res
    return wrapper

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"

@benchmark
def urlcont(url):
    """
    Загрузка картинки с телефоном в формате base64 по урлу
    :param url: Адрес расположения картинки
    :return:    Словарь содержащий картинку
    """
    try:
        req=get(url)
    except Exception as e:
        return e
    if req.status_code == 200:
        try:
            return json.loads(req.text)
        except Exception as e:
            return e

@benchmark
def phone(data):
    """
    Фнкция принимает словарь, выбирает данные относящиеся к картинке (регуляркой),
    перекодирует данные из base64 в (объект байтового типа), отдает библиотеке PIL.Image, распознает библиотекой tesseract
    :param data:    Словарь содержащий картинку
    :return:        Номер телефона в строковом виде
    """
    result=re.findall(r'(?<=data:image\/png;base64,).+',data['anonymImage64'])
    if result:
        img=Image.open(BytesIO(base64.b64decode(result[0])))
        text=pytesseract.image_to_string(img,lang="eng")
    return text

def main(url):
    data=urlcont(url)
    if isinstance(data,dict):
        phones=phone(data)
    else:
        phones='Загразка картинки закончилась с ошибкой: '+str(data)
    return phones

if __name__ == '__main__':
    print(main(r'https://www.avito.ru/items/phone/1622573389?pkey=2ce7f3b14adf6b14c0b9702ac0ccbfaa&vsrc=r&searchHash=ff47c499af980b66b5df065c6f56b0a5ee0cf712'))
