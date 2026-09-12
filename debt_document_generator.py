#!/usr/bin/env python3
"""
Программа для автоматического заполнения документа данными о задолженностях из Excel таблицы.
Минимальный интерфейс с использованием tkinter.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import re
from datetime import datetime
import os


class DebtDocumentGenerator:
    """Основной класс приложения для генерации документов с данными о задолженностях."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор документов по задолженностям")
        self.root.geometry("800x600")
        
        # Переменные для хранения данных
        self.excel_data = None
        self.template_content = ""
        self.output_path = ""
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Настройка пользовательского интерфейса."""
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Конфигурация сетки
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=1)
        
        # Заголовок
        title_label = ttk.Label(main_frame, text="Генератор документов по задолженностям", 
                               font=('Arial', 14, 'bold'))
        title_label.grid(row=0, column=0, pady=(0, 20))
        
        # Кнопки для выбора файлов
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=1, column=0, pady=10, sticky=tk.W)
        
        self.btn_excel = ttk.Button(btn_frame, text="📊 Выбрать Excel файл", 
                                   command=self._load_excel)
        self.btn_excel.grid(row=0, column=0, padx=5)
        
        self.btn_template = ttk.Button(btn_frame, text="📄 Выбрать шаблон документа", 
                                      command=self._load_template)
        self.btn_template.grid(row=0, column=1, padx=5)
        
        self.btn_generate = ttk.Button(btn_frame, text="✨ Сгенерировать документ", 
                                      command=self._generate_document, state=tk.DISABLED)
        self.btn_generate.grid(row=0, column=2, padx=5)
        
        # Статус бар
        self.status_var = tk.StringVar(value="Загрузите Excel файл и шаблон документа")
        status_label = ttk.Label(main_frame, textvariable=self.status_var, 
                                foreground="gray")
        status_label.grid(row=2, column=0, pady=10, sticky=tk.W)
        
        # Информация о загруженных файлах
        info_frame = ttk.LabelFrame(main_frame, text="Информация о файлах", padding="10")
        info_frame.grid(row=3, column=0, pady=10, sticky=(tk.W, tk.E))
        info_frame.columnconfigure(1, weight=1)
        
        ttk.Label(info_frame, text="Excel файл:").grid(row=0, column=0, sticky=tk.W)
        self.lbl_excel = ttk.Label(info_frame, text="Не выбран", foreground="gray")
        self.lbl_excel.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        ttk.Label(info_frame, text="Шаблон:").grid(row=1, column=0, sticky=tk.W)
        self.lbl_template = ttk.Label(info_frame, text="Не выбран", foreground="gray")
        self.lbl_template.grid(row=1, column=1, sticky=tk.W, padx=10)
        
        # Предпросмотр данных из Excel
        data_frame = ttk.LabelFrame(main_frame, text="Данные из Excel (предпросмотр)", padding="10")
        data_frame.grid(row=4, column=0, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)
        
        self.data_preview = scrolledtext.ScrolledText(data_frame, height=8, width=80)
        self.data_preview.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Область предпросмотра результата
        result_frame = ttk.LabelFrame(main_frame, text="Предпросмотр результата", padding="10")
        result_frame.grid(row=5, column=0, pady=10, sticky=(tk.W, tk.E, tk.N, tk.S))
        result_frame.columnconfigure(0, weight=1)
        result_frame.rowconfigure(0, weight=1)
        
        self.result_preview = scrolledtext.ScrolledText(result_frame, height=10, width=80)
        self.result_preview.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Инструкция
        instruction_text = """
ИНСТРУКЦИЯ ПО ИСПОЛЬЗОВАНИЮ:

1. Загрузите Excel файл с таблицей задолженностей
   - Таблица должна содержать столбцы: 'Компания', 'Клиент', 'Сумма задолженности', 
     'Дата', 'Номер договора' и другие необходимые данные
   
2. Загрузите шаблон документа (.txt или .docx)
   - В шаблоне используйте специальные метки для подстановки данных:
     {{company}} - название компании
     {{client}} - название клиента
     {{debt_amount}} - сумма задолженности
     {{date}} - дата
     {{contract_number}} - номер договора
     {{total_debt}} - общая сумма задолженности
   
3. Нажмите кнопку "Сгенерировать документ"
   - Программа автоматически найдет данные в таблице и подставит их в шаблон
   
4. Сохраните результат в текстовый файл
        """
        
        instr_frame = ttk.LabelFrame(main_frame, text="Инструкция", padding="10")
        instr_frame.grid(row=6, column=0, pady=10, sticky=(tk.W, tk.E))
        
        instr_text = scrolledtext.ScrolledText(instr_frame, height=6, width=80, wrap=tk.WORD)
        instr_text.grid(row=0, column=0, sticky=(tk.W, tk.E))
        instr_text.insert(tk.END, instruction_text)
        instr_text.config(state=tk.DISABLED)
    
    def _load_excel(self):
        """Загрузка Excel файла с данными о задолженностях."""
        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл с задолженностями",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Чтение Excel файла
            self.excel_data = pd.read_excel(file_path)
            
            # Обновление интерфейса
            self.lbl_excel.config(text=os.path.basename(file_path), foreground="green")
            self.status_var.set(f"Excel файл загружен: {len(self.excel_data)} записей")
            
            # Предпросмотр данных
            self.data_preview.delete(1.0, tk.END)
            preview_text = "Заголовки столбцов:\n"
            preview_text += ", ".join(self.excel_data.columns.tolist()) + "\n\n"
            preview_text += "Первые 5 записей:\n"
            preview_text += self.excel_data.head().to_string()
            self.data_preview.insert(tk.END, preview_text)
            
            # Проверка готовности к генерации
            self._check_ready_to_generate()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить Excel файл:\n{str(e)}")
            self.status_var.set("Ошибка загрузки Excel файла")
    
    def _load_template(self):
        """Загрузка файла шаблона документа."""
        file_path = filedialog.askopenfilename(
            title="Выберите шаблон документа",
            filetypes=[("Text files", "*.txt"), ("Word files", "*.docx"), ("All files", "*.*")]
        )
        
        if not file_path:
            return
        
        try:
            # Чтение файла шаблона
            if file_path.endswith('.docx'):
                try:
                    from docx import Document
                    doc = Document(file_path)
                    self.template_content = "\n".join([paragraph.text for paragraph in doc.paragraphs])
                except ImportError:
                    messagebox.showwarning("Внимание", 
                        "Для работы с .docx файлами установите библиотеку: pip install python-docx\n"
                        "Файл будет прочитан как текстовый.")
                    with open(file_path, 'r', encoding='utf-8') as f:
                        self.template_content = f.read()
            else:
                with open(file_path, 'r', encoding='utf-8') as f:
                    self.template_content = f.read()
            
            # Обновление интерфейса
            self.lbl_template.config(text=os.path.basename(file_path), foreground="green")
            self.status_var.set(f"Шаблон загружен: {len(self.template_content)} символов")
            
            # Предпросмотр шаблона
            self.result_preview.delete(1.0, tk.END)
            self.result_preview.insert(tk.END, "ШАБЛОН:\n\n" + self.template_content[:500])
            if len(self.template_content) > 500:
                self.result_preview.insert(tk.END, "\n... (показано первые 500 символов)")
            
            # Проверка готовности к генерации
            self._check_ready_to_generate()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить шаблон:\n{str(e)}")
            self.status_var.set("Ошибка загрузки шаблона")
    
    def _check_ready_to_generate(self):
        """Проверка готовности к генерации документа."""
        if self.excel_data is not None and self.template_content:
            self.btn_generate.config(state=tk.NORMAL)
            self.status_var.set("Готов к генерации документа")
    
    def _generate_document(self):
        """Генерация документа с подстановкой данных из Excel."""
        if self.excel_data is None or not self.template_content:
            messagebox.showwarning("Внимание", "Загрузите Excel файл и шаблон документа")
            return
        
        try:
            # Анализ шаблона и поиск меток
            placeholders = self._find_placeholders(self.template_content)
            
            if not placeholders:
                messagebox.showinfo("Информация", 
                    "В шаблоне не найдено меток для подстановки.\n"
                    "Используйте формат: {{company}}, {{client}}, {{debt_amount}} и т.д.")
                return
            
            # Генерация документа для каждой записи в Excel
            generated_docs = []
            
            for index, row in self.excel_data.iterrows():
                doc_content = self.template_content
                
                # Подстановка данных для каждой метки
                for placeholder in placeholders:
                    value = self._get_value_for_placeholder(placeholder, row, self.excel_data)
                    doc_content = doc_content.replace(placeholder, str(value))
                
                generated_docs.append(doc_content)
            
            # Объединение результатов (если записей несколько)
            if len(generated_docs) == 1:
                final_content = generated_docs[0]
            else:
                final_content = "\n\n".join([
                    f"=== Документ {i+1} ===\n\n{doc}" 
                    for i, doc in enumerate(generated_docs)
                ])
            
            # Предпросмотр результата
            self.result_preview.delete(1.0, tk.END)
            self.result_preview.insert(tk.END, final_content[:1000])
            if len(final_content) > 1000:
                self.result_preview.insert(tk.END, "\n\n... (показано первые 1000 символов)")
            
            # Сохранение результата
            self._save_document(final_content)
            
            self.status_var.set(f"Документ сгенерирован успешно! Обработано записей: {len(generated_docs)}")
            messagebox.showinfo("Успех", 
                f"Документ сгенерирован!\nОбработано записей: {len(generated_docs)}\n"
                f"Результат сохранен в выбранный файл.")
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сгенерировать документ:\n{str(e)}")
            self.status_var.set("Ошибка генерации документа")
    
    def _find_placeholders(self, text):
        """Поиск всех меток-заполнителей в тексте шаблона."""
        pattern = r'\{\{([^}]+)\}\}'
        matches = re.findall(pattern, text)
        return [f"{{{{{match}}}}}" for match in matches]
    
    def _get_value_for_placeholder(self, placeholder, row, full_data):
        """Получение значения для метки-заполнителя из данных Excel."""
        # Удаление скобок из имени метки
        key = placeholder.strip('{}').strip().lower()
        
        # Сопоставление с колонками Excel (регистронезависимое)
        for col in full_data.columns:
            if col.lower() == key or col.lower().replace(' ', '_') == key:
                return row[col]
        
        # Специальные обработки для некоторых меток
        if key == 'total_debt':
            # Сумма всех задолженностей
            debt_col = next((col for col in full_data.columns 
                           if 'сумма' in col.lower() or 'debt' in col.lower()), None)
            if debt_col:
                return full_data[debt_col].sum()
        
        if key == 'current_date':
            return datetime.now().strftime('%d.%m.%Y')
        
        # Если не найдено совпадение, возвращаем оригинальную метку
        return placeholder
    
    def _save_document(self, content):
        """Сохранение сгенерированного документа."""
        file_path = filedialog.asksaveasfilename(
            title="Сохранить сгенерированный документ",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialfile=f"document_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        )
        
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                self.output_path = file_path
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл:\n{str(e)}")


def main():
    """Точка входа в приложение."""
    root = tk.Tk()
    app = DebtDocumentGenerator(root)
    root.mainloop()


if __name__ == "__main__":
    main()
