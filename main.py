import os

import pandas as pd
import xml.etree.ElementTree as et


def add_year(date):
    return pd.to_datetime(date) + pd.offsets.DateOffset(years=1) - pd.offsets.DateOffset(days=1)


def data_preparing(source_excel_file):
    to_drop = ['Дата изменения', 'Владелец СИ', 'Рег. номер типа СИ', 'Заводской №/ Буквенно-цифровое обозначение',
               'Разряд эталона', 'Статус']
    table = pd.read_excel(source_excel_file, header=2).sort_values(by=['Дата поверки']).reset_index(drop=True)
    table = table.drop(columns=to_drop)

    # Форматирование даты
    table['Дата поверки'] = pd.to_datetime(table['Дата поверки'], dayfirst=True).dt.strftime('%Y-%m-%d')

    # Форматирование пригодности
    table['Пригодность'] = table['Пригодность'].map({'Да': '1', 'Нет': '0'})

    # Разбивка ФИО
    author = table['Автор записи'].str.split(' ', expand=True).rename({0: 'Фамилия', 1: 'Имя', 2: 'Отчество'}, axis=1)

    # Изъятие номера поверки
    document = table['Документ'].str.split('/', expand=True).drop(columns=[0, 1]).rename({2: 'Номер поверки'}, axis=1)

    # Создание даты окончания поверки
    verification_end = pd.Series(table['Дата поверки'].map(add_year), name='Дата окончания')

    # Объединение результатов
    table = pd.concat([table, verification_end, author, document], axis=1)

    # Дополнительный год для некоторых приборов
    additional_year = ['Фотометры фотоэлектрические']
    table.loc[table['Тип СИ'].isin(additional_year), 'Дата окончания'] = \
        pd.to_datetime(table.loc[table['Тип СИ'] == 'Фотометры фотоэлектрические', 'Дата окончания']) + \
        pd.offsets.DateOffset(years=1)
    table['Дата окончания'] = pd.to_datetime(table['Дата окончания']).dt.date

    table = table.drop(columns=['Автор записи', 'Документ'])
    table = table.sort_values(by=['Дата поверки']).reset_index(drop=True)

    # Перетасовка столбцов
    table = table[['Номер поверки', 'Дата поверки', 'Дата окончания', 'Тип СИ', 'Фамилия', 'Имя', 'Отчество', 'СНИЛС']]
    table = table.rename({'Номер поверки': 'NumberVerification', 'Дата поверки': 'DateVerification',
                          'Дата окончания': 'DateEndVerification', 'Тип СИ': 'TypeMeasuringInstrument',
                          'Фамилия': 'Last', 'Имя': 'First', 'Отчество': 'Middle', 'СНИЛС': 'SNILS'}, axis=1)

    # print(table.head(10))
    return table


def parse_xml(result_file):
    temp_file = 'temp.xml'
    tree = et.parse(temp_file)
    root = tree.getroot()
    verification_data = et.SubElement(root, 'VerificationMeasuringInstrumentData')
    childs = root.findall('.//VerificationMeasuringInstrument')
    for child in childs:
        child_copy = child.__copy__()
        verification_data.append(child_copy)
        root.remove(child)
    for verification in root.findall('.//VerificationMeasuringInstrumentData/VerificationMeasuringInstrument'):
        name = et.Element('Name')
        verification.append(name)
        name.append(verification.find('.//Last').__copy__())
        name.append(verification.find('.//First').__copy__())
        name.append(verification.find('.//Middle').__copy__())
        verification.remove(verification.find('.//Last'))
        verification.remove(verification.find('.//First'))
        verification.remove(verification.find('.//Middle'))

        employee = et.Element('ApprovedEmployee')
        verification.append(employee)
        employee.append(verification.find('.//Name').__copy__())
        employee.append(verification.find('.//SNILS').__copy__())
        verification.remove(verification.find('.//Name'))
        verification.remove(verification.find('.//SNILS'))
    # print(childs)
    et.indent(tree, "    ", 0)
    tree.write(result_file, encoding='UTF-8')
    if os.path.isfile(temp_file):
        os.remove(temp_file)

    with open(result_file, 'r+', encoding='UTF-8') as file:
        lines = file.readlines()
        lines.insert(0, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
        file.seek(0)
        file.writelines(lines)


if __name__ == '__main__':
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)

    is_correct_filename = False
    while not is_correct_filename:
        try:
            source_excel_file = input('Enter source excel file name with extension: ')
            table = data_preparing(source_excel_file)
            is_correct_filename = True
        except FileNotFoundError:
            print('ERROR: Filename isn\'t correct')


    table.to_xml('temp.xml', root_name='Message', row_name='VerificationMeasuringInstrument', index=False)
    result_xml = 'result.xml'
    parse_xml(result_xml)

    print(f'Your {result_xml} file is ready :)')
