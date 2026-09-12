#!/usr/bin/env python3
"""
Скрипт для создания тестовых данных (Excel файл и шаблон документа)
для демонстрации работы программы генерации документов.
"""

import pandas as pd
from datetime import datetime, timedelta
import random


def create_test_excel():
    """Создание тестового Excel файла с данными о задолженностях."""
    
    # Данные для примера
    companies = ["ООО Ромашка", "ЗАО Вектор", "ИП Петров", "ООО ТехноСтрой", "АО МеталлТрейд"]
    clients = ["ПАО Сбербанк", "ООО ГазпромНефть", "ЗАО АльфаБанк", "ИП Сидоров", "ООО СтройМатериалы"]
    
    data = {
        'Компания': [],
        'Клиент': [],
        'Номер договора': [],
        'Дата': [],
        'Сумма задолженности': [],
        'Валюта': [],
        'Тип задолженности': [],
        'Примечание': []
    }
    
    # Генерация 10 записей
    for i in range(10):
        data['Компания'].append(random.choice(companies))
        data['Клиент'].append(random.choice(clients))
        data['Номер договора'].append(f"ДГ-{random.randint(100, 999)}/{2024}")
        
        # Дата в диапазоне последних 6 месяцев
        days_ago = random.randint(1, 180)
        date = datetime.now() - timedelta(days=days_ago)
        data['Дата'].append(date.strftime('%d.%m.%Y'))
        
        # Сумма задолженности от 10 000 до 5 000 000
        data['Сумма задолженности'].append(random.randint(10000, 5000000))
        data['Валюта'].append('RUB')
        data['Тип задолженности'].append(random.choice(['Дебиторская', 'Кредиторская']))
        data['Примечание'].append(f"Запись #{i+1}")
    
    # Создание DataFrame и сохранение в Excel
    df = pd.DataFrame(data)
    filename = 'test_debts.xlsx'
    df.to_excel(filename, index=False)
    
    print(f"✅ Тестовый Excel файл создан: {filename}")
    print(f"📊 Количество записей: {len(df)}")
    print(f"\n📋 Структура таблицы:")
    print(df.to_string())
    
    return filename


def create_test_template():
    """Создание тестового шаблона документа."""
    
    template_content = """АКТ СВЕРКИ ВЗАИМНЫХ РАСЧЁТОВ
№ {{contract_number}} от {{date}}

г. Москва                                                                        «___» ________ 20__ г.

{{company}}, именуемое в дальнейшем «Компания», в лице Генерального директора, 
действующего на основании Устава, с одной стороны, и 

{{client}}, именуемое в дальнейшем «Клиент», в лице Генерального директора, 
действующего на основании Устава, с другой стороны,

составили настоящий акт о нижеследующем:

1. По состоянию на {{current_date}} задолженность Компании перед Клиентом составляет: 
   {{debt_amount}} {{currency}}.

2. По состоянию на {{current_date}} задолженность Клиента перед Компанией составляет: 
   {{total_debt}} {{currency}}.

3. Разница между взаимными обязательствами сторон составляет: 
   {{debt_amount}} {{currency}}.

4. Стороны подтверждают, что все расчеты произведены в соответствии с договором 
   № {{contract_number}} от {{date}}.

5. Настоящий акт составлен в двух экземплярах, имеющих одинаковую юридическую силу, 
   по одному для каждой из сторон.

ПОДПИСИ СТОРОН:

От Компании: ___________________ /___________________/
            (подпись)           (ФИО)

М.П.

От Клиента: ___________________ /___________________/
           (подпись)           (ФИО)

М.П.


================================================================================

ПРИМЕЧАНИЕ: {{note}}
Тип задолженности: {{debt_type}}
"""
    
    filename = 'template.txt'
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(template_content)
    
    print(f"\n✅ Тестовый шаблон создан: {filename}")
    print("📄 Используемые метки:")
    print("   {{company}} - название компании")
    print("   {{client}} - название клиента")
    print("   {{contract_number}} - номер договора")
    print("   {{date}} - дата договора")
    print("   {{current_date}} - текущая дата (автоматически)")
    print("   {{debt_amount}} - сумма задолженности")
    print("   {{total_debt}} - общая сумма (автоматический расчет)")
    print("   {{currency}} - валюта")
    print("   {{note}} - примечание")
    print("   {{debt_type}} - тип задолженности")
    
    return filename


if __name__ == "__main__":
    print("=" * 60)
    print("СОЗДАНИЕ ТЕСТОВЫХ ДАННЫХ ДЛЯ ГЕНЕРАТОРА ДОКУМЕНТОВ")
    print("=" * 60)
    
    # Создание тестовых файлов
    excel_file = create_test_excel()
    template_file = create_test_template()
    
    print("\n" + "=" * 60)
    print("ТЕСТОВЫЕ ФАЙЛЫ ГОТОВЫ!")
    print("=" * 60)
    print(f"\nДля запуска программы выполните:")
    print(f"  python3 debt_document_generator.py")
    print(f"\nИспользуйте созданные файлы:")
    print(f"  1. Excel: {excel_file}")
    print(f"  2. Шаблон: {template_file}")
