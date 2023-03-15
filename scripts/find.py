from re import sub
from bs4 import BeautifulSoup
from os.path import join
from typing import Union, Dict, List


def get_description(bs: BeautifulSoup) -> Dict:
    """
    Получение описаний

    :param bs: XML-файл
    :type bs: BeautifulSoup
    :return: Возвращает словарь с описаниями
    :rtype: dict
    """

    info_smeta = {
        'объект': '',
        'составил': '',
        'согласовано': '',
        'заказчик': '',
        'принял': '',
        'проверил': '',
        'утверждаю': ''
    }

    properties = bs.find(name='Properties') or None

    if properties:
        description = properties.get('Description') or None
        if description:
            info_smeta['объект'] = description

    gs_doc_signatures = bs.find(name='GsDocSignatures') or None

    if gs_doc_signatures:

        if gs_doc_signatures.find(name='Item', attrs={"Caption": "Составил"}):
            r = gs_doc_signatures.find(name='Item', attrs={"Caption": "Составил"})
            info_smeta['cоставил'] = r.get('Value')

        if gs_doc_signatures.find(name='Item', attrs={"Caption": "Согласовано"}):
            r = gs_doc_signatures.find(name='Item', attrs={"Caption": "Согласовано"})
            info_smeta["согласовано"] = r.get('Value')

        if gs_doc_signatures.find(name='Item', attrs={"Caption": "Заказчик"}):
            r = gs_doc_signatures.find(name='Item', attrs={"Caption": "Заказчик"})
            info_smeta["заказчик"] = r.get('Value')

        if gs_doc_signatures.find(name='Item', attrs={"Caption": "Принял"}):
            r = gs_doc_signatures.find(name='Item', attrs={"Caption": "Принял"})
            info_smeta["принял"] = r.get('Value')

        if gs_doc_signatures.find(name='Item', attrs={"Caption": "Проверил"}):
            r = gs_doc_signatures.find(name='Item', attrs={"Caption": "Проверил"})
            info_smeta["проверил"] = r.get('Value')

        if gs_doc_signatures.find(name='Item', attrs={"Caption": "Утверждаю"}):
            r = gs_doc_signatures.find(name='Item', attrs={"Caption": "Утверждаю"})
            info_smeta["утверждаю"] = r.get('Value')

    return info_smeta


def get_table(bs: BeautifulSoup) -> Union[List, Dict]:
    """
    Получение таблицы

    :param bs: XML-файл
    :type bs: BeautifulSoup
    :return: Возвращает словарь либо список таблицы
    :rtype:dict, list
    """
    count = 1

    table = list()

    for chapter in bs.find_all(name='Chapter'):  # Итерация по разделам
        text = f'Раздел {chapter.get("SysID")} {chapter.get("Caption")}'  # Сохранение наиминования раздела
        table.append({text: list()})

        for position in chapter.find_all(name='Position'):  # Итерация по позизациям текущего раздела

            if not position.get('Code'):
                continue  # Пропустить если код отсутствует

            quantity = position.find(name=['Quantity']) or None
            if quantity:
                result = quantity.get('Result') or '0'
                fx = quantity.get('Fx')
                fx = fx if '*' in fx or '+' in fx or '/' in fx or '-' in fx else None
            else:
                result = '0'
                fx = None

            result = float(sub(',', '.', result))
            units = position.get('Units') or ''

            if units != '':
                number = units.strip().split(' ')
                if number[0].isdigit():
                    result *= int(number[0])
                    units = sub(number[0], '', units)

            pos = dict(
                number=position.get('Number'),  # Номер по пункту
                number_pp=count,  # Номер по смете
                caption=position.get('Caption'),  # Описание
                units=units,  # Ед. изм.
                result=round(result, 3),  # Количество
                fx=fx  # Формула
            )
            count += 1
            table[-1][text].append(pos)

    return table


def find(name_smeta: str) -> Union[List, Dict]:
    """
    Поиск данных для формирования вывода формы в excell

    :param name_smeta: Имя сметы без расширения
    :type name_smeta:str
    :return: Найденые данные
    :rtype:list, dict
    """

    path = join('download', f'{name_smeta}.xml')

    with open(path, 'r', encoding='windows-1251') as f:
        file = f.read()
        file = sub('windows-1251', 'utf-8', file)

    bs = BeautifulSoup(file, 'xml')

    table = get_table(bs)

    info_smeta = get_description(bs)

    info_smeta['table'] = table

    return info_smeta
