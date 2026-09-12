#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автоматическая генерация документов о задолженностях
Программа анализирует шаблон документа, находит числовые значения и другие данные,
заменяет их на реальные данные из Excel таблицы
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import re
import os
from datetime import datetime
from pathlib import Path


class DebtDocumentAutomator:
    """Основной класс приложения"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор документов о задолженностях")
        self.root.geometry("1200x700")
        
        # Переменные
        self.excel_file = None
        self.template_file = None
        self.df = None
        self.template_text = ""
        
        self.create_widgets()
        
    def create_widgets(self):
        """Создание интерфейса"""
        
        # Верхняя панель с кнопками
        top_frame = ttk.Frame(self.root, padding="10")
        top_frame.pack(fill=tk.X)
        
        ttk.Button(top_frame, text="📂 Выбрать Excel файл", 
                  command=self.load_excel).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="📄 Выбрать шаблон документа", 
                  command=self.load_template).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="⚙️ Настроить сопоставление", 
                  command=self.setup_mapping).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="🚀 Сгенерировать документ", 
                  command=self.generate_document).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(top_frame, text="💾 Сохранить результат", 
                  command=self.save_result).pack(side=tk.LEFT, padx=5)
        
        # Статус бар
        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(top_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.RIGHT, fill=tk.X, expand=True, padx=5)
        
        # Основная область с вкладками
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Вкладка 1: Excel данные
        excel_frame = ttk.Frame(notebook)
        notebook.add(excel_frame, text="📊 Данные из Excel")
        
        self.excel_tree = ttk.Treeview(excel_frame, show='headings')
        excel_scrollbar = ttk.Scrollbar(excel_frame, orient=tk.VERTICAL, 
                                       command=self.excel_tree.yview)
        self.excel_tree.configure(yscrollcommand=excel_scrollbar.set)
        
        self.excel_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        excel_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Вкладка 2: Шаблон документа
        template_frame = ttk.Frame(notebook)
        notebook.add(template_frame, text="📝 Шаблон документа")
        
        self.template_text_widget = scrolledtext.ScrolledText(template_frame, 
                                                             wrap=tk.WORD, width=80, height=20)
        self.template_text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Вкладка 3: Результат
        result_frame = ttk.Frame(notebook)
        notebook.add(result_frame, text="✅ Результат")
        
        self.result_text_widget = scrolledtext.ScrolledText(result_frame, 
                                                           wrap=tk.WORD, width=80, height=20)
        self.result_text_widget.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Вкладка 4: Настройки сопоставления
        mapping_frame = ttk.Frame(notebook)
        notebook.add(mapping_frame, text="⚙️ Настройки")
        
        ttk.Label(mapping_frame, text="Настройте соответствие полей Excel и элементов шаблона:",
                 font=('Arial', 11, 'bold')).pack(pady=10)
        
        self.mapping_frame_inner = ttk.Frame(mapping_frame)
        self.mapping_frame_inner.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        ttk.Label(mapping_frame, text="💡 Программа автоматически распознает суммы, даты и номера в шаблоне",
                 font=('Arial', 9), foreground='blue').pack(pady=5)
        
    def load_excel(self):
        """Загрузка Excel файла"""
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл с задолженностями",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.excel_file = file_path
                self.df = pd.read_excel(file_path)
                
                # Отображение данных в таблице
                self.excel_tree.delete(*self.excel_tree['children'])
                self.excel_tree["columns"] = list(self.df.columns)
                
                for col in self.df.columns:
                    self.excel_tree.heading(col, text=col)
                    self.excel_tree.column(col, width=150, minwidth=100)
                
                for idx, row in self.df.iterrows():
                    values = [str(val) for val in row.values]
                    self.excel_tree.insert('', tk.END, values=values)
                
                self.status_var.set(f"✅ Загружено: {os.path.basename(file_path)} ({len(self.df)} записей)")
                messagebox.showinfo("Успех", f"Загружено {len(self.df)} записей из Excel")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить Excel файл:\n{str(e)}")
    
    def load_template(self):
        """Загрузка шаблона документа"""
        file_path = filedialog.askopenfilename(
            title="Выберите шаблон документа",
            filetypes=[("Text files", "*.txt *.docx *.rtf"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                self.template_file = file_path
                
                # Чтение файла
                if file_path.endswith('.docx'):
                    import docx
                    doc = docx.Document(file_path)
                    self.template_text = "\n".join([para.text for para in doc.paragraphs])
                else:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        self.template_text = f.read()
                
                self.template_text_widget.delete(1.0, tk.END)
                self.template_text_widget.insert(1.0, self.template_text)
                
                self.status_var.set(f"✅ Шаблон загружен: {os.path.basename(file_path)}")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить шаблон:\n{str(e)}")
    
    def setup_mapping(self):
        """Настройка сопоставления полей"""
        if self.df is None:
            messagebox.showwarning("Внимание", "Сначала загрузите Excel файл")
            return
        
        # Создаем диалоговое окно
        dialog = tk.Toplevel(self.root)
        dialog.title("Настройка сопоставления полей")
        dialog.geometry("600x400")
        
        ttk.Label(dialog, text="Соответствие полей Excel и данных в документе:",
                 font=('Arial', 11, 'bold')).pack(pady=10)
        
        # Создаем таблицу сопоставления
        mapping_frame = ttk.Frame(dialog)
        mapping_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        columns = ('field', 'description', 'example')
        tree = ttk.Treeview(mapping_frame, columns=columns, show='headings', height=15)
        
        tree.heading('field', text='Поле Excel')
        tree.heading('description', text='Описание')
        tree.heading('example', text='Пример значения')
        
        tree.column('field', width=200)
        tree.column('description', width=250)
        tree.column('example', width=150)
        
        scrollbar = ttk.Scrollbar(mapping_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Заполняем таблицу
        for col in self.df.columns:
            sample_value = str(self.df[col].iloc[0]) if len(self.df) > 0 else ""
            tree.insert('', tk.END, values=(col, self.get_field_description(col), sample_value))
        
        ttk.Button(dialog, text="Закрыть", command=dialog.destroy).pack(pady=10)
    
    def get_field_description(self, field_name):
        """Автоматическое описание поля на основе названия"""
        field_lower = field_name.lower()
        
        descriptions = {
            'company': 'Название компании',
            'client': 'Название клиента',
            'debt': 'Сумма задолженности',
            'amount': 'Сумма',
            'sum': 'Сумма',
            'total': 'Общая сумма',
            'contract': 'Номер договора',
            'agreement': 'Номер соглашения',
            'date': 'Дата',
            'upd': 'Номер УПД',
            'uppd': 'Номер УПД',
            'payment': 'Платежное поручение',
            'document': 'Документ',
            'number': 'Номер',
            'rub': 'Рубли',
            'kop': 'Копейки'
        }
        
        for key, desc in descriptions.items():
            if key in field_lower:
                return desc
        
        return 'Данные из таблицы'
    
    def extract_numbers_from_template(self, text):
        """Извлечение всех числовых значений из шаблона"""
        # Паттерны для поиска чисел
        patterns = [
            r'\d{1,3}(?:\s?\d{3})*(?:,\d{2})?',  # Числа с пробелами и копейками
            r'\d+\.\d{2}',  # Числа с точкой
            r'\d+,\d{2}',   # Числа с запятой
        ]
        
        numbers = []
        for pattern in patterns:
            matches = re.findall(pattern, text)
            numbers.extend(matches)
        
        return numbers
    
    def extract_dates_from_template(self, text):
        """Извлечение дат из шаблона"""
        date_patterns = [
            r'\d{2}\.\d{2}\.\d{4}',  # ДД.ММ.ГГГГ
            r'\d{2}/\d{2}/\d{4}',    # ДД/ММ/ГГГГ
        ]
        
        dates = []
        for pattern in date_patterns:
            matches = re.findall(pattern, text)
            dates.extend(matches)
        
        return dates
    
    def extract_contract_names(self, text):
        """Извлечение названий договоров"""
        contract_pattern = r'Договор[- ]?(\d+|[А-Яа-я]+)'
        matches = re.findall(contract_pattern, text)
        return matches
    
    def number_to_words(self, amount):
        """Преобразование числа в слова (рубли и копейки)"""
        try:
            # Очистка числа от пробелов и замена запятой на точку
            amount_str = str(amount).replace(' ', '').replace(',', '.')
            amount_float = float(amount_str)
            
            rubles = int(amount_float)
            kopecks = int(round((amount_float - rubles) * 100))
            
            # Простая реализация (можно расширить)
            rubles_words = self.number_to_words_rub(rubles)
            kopecks_words = self.kopecks_to_words(kopecks)
            
            return f"{rubles_words} {kopecks_words}"
            
        except:
            return str(amount)
    
    def number_to_words_rub(self, n):
        """Число прописью (рубли)"""
        if n == 0:
            return "ноль рублей"
        
        ones = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
        teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", 
                "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"]
        tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", 
               "семьдесят", "восемьдесят", "девяносто"]
        hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", 
                   "семьсот", "восемьсот", "девятьсот"]
        
        scales = ["", "тысяча", "миллион", "миллиард"]
        scale_multipliers = [1, 1000, 1000000, 1000000000]
        
        result = []
        scale_idx = 0
        
        while n > 0:
            chunk = n % 1000
            if chunk > 0:
                chunk_words = []
                
                h = chunk // 100
                t = (chunk % 100) // 10
                o = chunk % 10
                
                if h > 0:
                    chunk_words.append(hundreds[h])
                
                if 10 <= chunk % 100 <= 19:
                    chunk_words.append(teens[chunk % 10])
                else:
                    if t >= 2:
                        chunk_words.append(tens[t])
                    if o > 0:
                        chunk_words.append(ones[o])
                
                if scale_idx > 0:
                    if chunk == 1:
                        chunk_words.append(scales[scale_idx])
                    elif 2 <= chunk <= 4:
                        if scales[scale_idx] == "тысяча":
                            chunk_words.append("тысячи")
                        else:
                            chunk_words.append(scales[scale_idx])
                    else:
                        chunk_words.append(scales[scale_idx])
                
                result.append(" ".join(chunk_words))
            
            n //= 1000
            scale_idx += 1
        
        result.reverse()
        rub_text = " ".join(result)
        
        # Добавляем "рублей" с правильным окончанием
        last_digit = rubles % 10
        last_two_digits = rubles % 100
        
        if last_two_digits >= 11 and last_two_digits <= 19:
            rub_text += " рублей"
        elif last_digit == 1:
            rub_text += " рубль"
        elif 2 <= last_digit <= 4:
            rub_text += " рубля"
        else:
            rub_text += " рублей"
        
        return rub_text.capitalize()
    
    def kopecks_to_words(self, k):
        """Копейки прописью"""
        ones = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
        teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", 
                "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"]
        tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", 
               "семьдесят", "восемьдесят", "девяносто"]
        
        if k == 0:
            return "0 копеек"
        
        result = []
        t = k // 10
        o = k % 10
        
        if 10 <= k <= 19:
            result.append(teens[k - 10])
        else:
            if t >= 2:
                result.append(tens[t])
            if o > 0:
                result.append(ones[o])
        
        kop_text = " ".join(result)
        
        if 11 <= k <= 19:
            kop_text += " копеек"
        elif k % 10 == 1:
            kop_text += " копейка"
        elif 2 <= k % 10 <= 4:
            kop_text += " копейки"
        else:
            kop_text += " копеек"
        
        return kop_text
    
    def generate_document(self):
        """Генерация документа"""
        if self.df is None or self.template_text == "":
            messagebox.showwarning("Внимание", "Загрузите Excel файл и шаблон документа")
            return
        
        try:
            results = []
            
            # Обрабатываем каждую строку Excel
            for idx, row in self.df.iterrows():
                result_text = self.process_template_for_row(row)
                results.append(f"=== Документ #{idx + 1} ===\n{result_text}\n")
            
            final_result = "\n".join(results)
            
            self.result_text_widget.delete(1.0, tk.END)
            self.result_text_widget.insert(1.0, final_result)
            
            self.status_var.set(f"✅ Сгенерировано документов: {len(results)}")
            messagebox.showinfo("Успех", f"Сгенерировано {len(results)} документов")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сгенерировать документ:\n{str(e)}")
            import traceback
            traceback.print_exc()
    
    def process_template_for_row(self, row):
        """Обработка шаблона для одной строки данных"""
        result = self.template_text
        
        # Преобразуем row в словарь для удобного доступа
        data = row.to_dict()
        
        # 1. Замена явных совпадений по названиям колонок
        for col_name, value in data.items():
            if pd.notna(value):
                # Ищем упоминания названия колонки в тексте (в разных формах)
                col_lower = col_name.lower()
                
                # Прямая замена для простых случаев
                if f"[{col_name}]" in result:
                    result = result.replace(f"[{col_name}]", str(value))
                
                # Замена для названий компаний/клиентов
                if any(keyword in col_lower for keyword in ['company', 'organization', 'org', 'компания', 'организация']):
                    result = self.replace_company_names(result, str(value))
                
                if any(keyword in col_lower for keyword in ['client', 'customer', 'контрагент', 'клиент']):
                    result = self.replace_client_names(result, str(value))
        
        # 2. Автоматическая замена числовых значений
        result = self.replace_numeric_values(result, data)
        
        # 3. Замена дат
        result = self.replace_dates(result, data)
        
        # 4. Замена сумм прописью
        result = self.replace_amounts_in_words(result, data)
        
        return result
    
    def replace_company_names(self, text, company_name):
        """Замена названий компании"""
        patterns = [
            (r'Сторона-1', company_name),
            (r'ООО\s*"[^"]*"', company_name),
            (r'ООО\s*[\'"][^\'"]*[\'"]', company_name),
        ]
        
        for pattern, replacement in patterns:
            text = re.sub(pattern, replacement, text)
        
        return text
    
    def replace_client_names(self, text, client_name):
        """Замена названий клиента"""
        patterns = [
            (r'Сторона-2', client_name),
        ]
        
        for pattern, replacement in patterns:
            text = re.sub(pattern, replacement, text)
        
        return text
    
    def replace_numeric_values(self, text, data):
        """Автоматическая замена числовых значений"""
        # Извлекаем все числа из шаблона
        numbers_in_template = self.extract_numbers_from_template(text)
        
        # Ищем числовые поля в данных
        numeric_fields = {}
        for col, value in data.items():
            if pd.notna(value):
                try:
                    num_val = float(str(value).replace(' ', '').replace(',', '.'))
                    numeric_fields[col] = num_val
                except:
                    pass
        
        # Если нашли числовые поля, пробуем сопоставить их с числами в шаблоне
        if numeric_fields:
            # Сортируем числа по убыванию для корректной замены
            sorted_numbers = sorted(numbers_in_template, key=lambda x: float(x.replace(' ', '').replace(',', '.')) if x.replace(' ', '').replace(',', '.').replace('.', '').isdigit() else 0, reverse=True)
            
            # Берем основные суммы из данных
            debt_values = []
            for col, val in numeric_fields.items():
                if any(kw in col.lower() for kw in ['debt', 'sum', 'amount', 'total', 'задолженность', 'сумма']):
                    debt_values.append(val)
            
            if debt_values and sorted_numbers:
                # Заменяем найденные числа на реальные значения
                for i, num_str in enumerate(sorted_numbers[:len(debt_values)]):
                    if i < len(debt_values):
                        new_value = debt_values[i]
                        # Форматируем число
                        formatted = f"{new_value:,.2f}".replace(',', ' ').replace(',', ',')
                        text = text.replace(num_str, formatted, 1)
        
        return text
    
    def replace_dates(self, text, data):
        """Замена дат"""
        dates_in_template = self.extract_dates_from_template(text)
        
        date_fields = {}
        for col, value in data.items():
            if pd.notna(value) and any(kw in col.lower() for kw in ['date', 'дата', 'data']):
                date_fields[col] = value
        
        if date_fields and dates_in_template:
            for date_col, date_val in date_fields.items():
                if dates_in_template:
                    # Форматируем дату
                    try:
                        if isinstance(date_val, datetime):
                            formatted_date = date_val.strftime('%d.%m.%Y')
                        else:
                            formatted_date = str(date_val)
                        
                        text = text.replace(dates_in_template[0], formatted_date, 1)
                    except:
                        pass
        
        return text
    
    def replace_amounts_in_words(self, text, data):
        """Замена сумм прописью"""
        # Ищем паттерн: число, затем текст в скобках
        pattern = r'(\d{1,3}(?:\s?\d{3})*(?:,\d{2})?)\s*\(([^)]+)\)'
        
        matches = list(re.finditer(pattern, text))
        
        # Ищем числовые поля с суммами
        amount_fields = {}
        for col, value in data.items():
            if pd.notna(value):
                try:
                    num_val = float(str(value).replace(' ', '').replace(',', '.'))
                    if any(kw in col.lower() for kw in ['debt', 'sum', 'amount', 'total', 'задолженность', 'сумма']):
                        amount_fields[col] = num_val
                except:
                    pass
        
        if amount_fields:
            amounts = list(amount_fields.values())
            for i, match in enumerate(matches):
                if i < len(amounts):
                    number_str = match.group(1)
                    words_in_parens = match.group(2)
                    
                    # Проверяем, похоже ли это на сумму прописью
                    if any(word in words_in_parens.lower() for word in ['рубль', 'копейка', 'тысяч', 'миллион']):
                        new_amount = amounts[i]
                        amount_in_words = self.number_to_words(new_amount)
                        formatted_number = f"{new_amount:,.2f}".replace(',', ' ')
                        
                        old_text = f"{number_str} ({words_in_parens})"
                        new_text = f"{formatted_number} ({amount_in_words})"
                        
                        text = text.replace(old_text, new_text, 1)
        
        return text
    
    def save_result(self):
        """Сохранение результата"""
        result_text = self.result_text_widget.get(1.0, tk.END).strip()
        
        if not result_text:
            messagebox.showwarning("Внимание", "Нет результата для сохранения")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="Сохранить результат",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("Word documents", "*.docx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                if file_path.endswith('.docx'):
                    from docx import Document
                    doc = Document()
                    for line in result_text.split('\n'):
                        doc.add_paragraph(line)
                    doc.save(file_path)
                else:
                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(result_text)
                
                self.status_var.set(f"✅ Сохранено: {os.path.basename(file_path)}")
                messagebox.showinfo("Успех", f"Документ сохранен:\n{file_path}")
                
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")


def main():
    root = tk.Tk()
    app = DebtDocumentAutomator(root)
    root.mainloop()


if __name__ == "__main__":
    main()
