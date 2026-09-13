import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import re
import os
from openpyxl import load_workbook
from PIL import Image, ImageTk, ImageDraw
import random

def number_to_words_rus(n):
    if n is None or n == '':
        return ""
    try:
        if isinstance(n, str):
            n = float(n.replace(',', '.'))
        n = float(n)
    except (ValueError, TypeError):
        return str(n)

    if n == 0:
        return "ноль"

    integer_part = int(abs(n))
    fractional_part = round((abs(n) - integer_part) * 100)

    words = []
    
    ones = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
    teens = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать", "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"]
    tens = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят", "семьдесят", "восемьдесят", "девяносто"]
    hundreds = ["", "сто", "двести", "триста", "четыреста", "пятьсот", "шестьсот", "семьсот", "восемьсот", "девятьсот"]
    
    thousands_words = ["", "тысяча", "тысячи", "тысяч"]
    millions_words = ["", "миллион", "миллиона", "миллионов"]

    def get_triplet(num, level=0):
        if num == 0:
            return ""
        
        h = num // 100
        t = (num % 100) // 10
        o = num % 10
        
        res = []
        if h > 0:
            res.append(hundreds[h])
        
        if t == 1:
            res.append(teens[o])
        else:
            if t > 0:
                res.append(tens[t])
            if o > 0:
                res.append(ones[o])
        
        return " ".join(res)

    result_str = ""
    
    millions = integer_part // 1000000
    rest = integer_part % 1000000
    
    if millions > 0:
        m_rem = millions % 10
        m_tens = millions % 100
        if m_tens > 10 and m_tens < 20:
            m_word = millions_words[3]
        elif m_rem == 1:
            m_word = millions_words[1]
        elif m_rem in [2, 3, 4]:
            m_word = millions_words[2]
        else:
            m_word = millions_words[3]
            
        result_str += f"{get_triplet(millions)} {m_word} "

    thousands = rest // 1000
    rest = rest % 1000
    
    if thousands > 0:
        t_rem = thousands % 10
        t_tens = thousands % 100
        if t_tens > 10 and t_tens < 20:
            t_word = thousands_words[3]
        elif t_rem == 1:
            t_word = thousands_words[1]
        elif t_rem in [2, 3, 4]:
            t_word = thousands_words[2]
        else:
            t_word = thousands_words[3]
            
        txt = get_triplet(thousands)
        if thousands == 1: txt = "одна"
        elif thousands == 2: txt = "две"
        
        result_str += f"{txt} {t_word} "

    if rest > 0:
        result_str += f"{get_triplet(rest)} "

    kopecks_text = ""
    if fractional_part > 0:
        kopecks_text = f"{fractional_part:02d} копеек"
    else:
        kopecks_text = "00 копеек"

    final_text = result_str.strip().capitalize() + f" рублей {kopecks_text}"
    
    return final_text

def create_flower_bg(width, height):
    # Создаем изображение с основным цветом (фисташковый)
    img = Image.new('RGBA', (width, height), color=(0, 0, 0, 0)) # Прозрачный фон для начала
    draw = ImageDraw.Draw(img)
    
    # Нежные цвета: желтый, розовый, белый, кремовый
    colors = [
        (255, 250, 205, 200), # LemonChiffon (полупрозрачный)
        (255, 182, 193, 180), # LightPink (полупрозрачный)
        (255, 255, 255, 210), # White (полупрозрачный)
        (255, 253, 208, 190)  # Cream
    ] 
    
    random.seed(42) # Чтобы цветы не двигались при каждом запуске
    
    for _ in range(70): # Количество цветов
        x = random.randint(30, width - 30)
        y = random.randint(30, height - 30)
        size = random.randint(15, 25) # Размер цветка
        color = random.choice(colors)
        
        # Рисуем 5 лепестков эллипсами вокруг центра
        # Лепесток 1 (Верх)
        draw.ellipse([x - size//2, y - size*1.2, x + size//2, y - size//2], fill=color)
        # Лепесток 2 (Низ)
        draw.ellipse([x - size//2, y + size//2, x + size//2, y + size*1.2], fill=color)
        # Лепесток 3 (Лево)
        draw.ellipse([x - size*1.2, y - size//2, x - size//2, y + size//2], fill=color)
        # Лепесток 4 (Право)
        draw.ellipse([x + size//2, y - size//2, x + size*1.2, y + size//2], fill=color)
        # Лепесток 5 (Центральный для объема)
        draw.ellipse([x - size//3, y - size//3, x + size//3, y + size//3], fill=color)
        
        # Серединка (более яркая желтая)
        center_size = size // 4
        draw.ellipse([x - center_size, y - center_size, x + center_size, y + center_size], fill=(255, 255, 224, 255))

    # Создаем финальное изображение с фисташковым фоном и накладываем цветы
    base = Image.new('RGB', (width, height), color="#C1E1C1")
    base.paste(img, (0,0), img) # Накладываем цветы с учетом прозрачности
    return base

class DebtGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор документов")
        self.root.geometry("900x750")
        
        self.bg_color = "#C1E1C1"
        self.btn_color = "#4CAF50"
        self.btn_fg = "white"
        
        self.root.configure(bg=self.bg_color)

        # Создаем фон с цветами
        self.bg_image = create_flower_bg(900, 750)
        self.photo_img = ImageTk.PhotoImage(self.bg_image)
        
        bg_label = tk.Label(root, image=self.photo_img, bg=self.bg_color)
        bg_label.place(x=0, y=0, relwidth=1, relheight=1)
        
        self.df = None
        self.wb = None

        # Контекстное меню
        self.context_menu = tk.Menu(root, tearoff=0)
        self.context_menu.add_command(label="Копировать", command=self.copy_text)
        self.context_menu.add_command(label="Вставить", command=self.paste_text)
        self.context_menu.add_command(label="Выделить всё", command=self.select_all_text)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Удалить", command=self.delete_text)
        
        instruction_text = (
            "ИНСТРУКЦИЯ:\n"
            "1. Выберите Excel файл (.xlsx). Убедитесь, что первые строки блоков закрашены ЦВЕТОМ.\n"
            "2. Укажите номер ПОСЛЕДНЕЙ строки данных.\n"
            "3. Шаблон: Используйте [A12] для ячейки A12. Программа повторит шаблон для каждого цветного блока.\n"
            "   - [A12!]: Соберет A12, A13, A14... до следующего цветного или пустой ячейки (через запятую).\n"
            "   - [H12|words]: Преобразует число в текст прописью.\n"
            "   - Важно: Номер строки в метке (например, 12) задает НАЧАЛО поиска первого блока."
        )
        
        self.create_widgets(instruction_text)
        
    def create_widgets(self, instruction_text):
        # Заголовок (прозрачный фон)
        title_label = tk.Label(self.root, text="Автоматическое заполнение документов", 
                               font=("Arial", 16, "bold"), bg=self.bg_color, fg="#2E7D32")
        title_label.pack(pady=10)
        
        # --- Блок выбора файла (БЕЗ ФРЕЙМА, чтобы не было зеленой полосы) ---
        lbl_file = tk.Label(self.root, text="Excel файл (.xlsx):", bg=self.bg_color, font=("Arial", 10), fg="#1B5E20")
        lbl_file.pack(anchor="w", padx=40)
        
        self.file_path_var = tk.StringVar()
        self.file_entry = tk.Entry(self.root, textvariable=self.file_path_var, width=65, font=("Arial", 10))
        self.file_entry.pack(padx=40, pady=2)
        
        btn_browse = tk.Button(self.root, text="Обзор...", command=self.browse_file, 
                               bg="#2196F3", fg="white", font=("Arial", 9))
        btn_browse.pack(pady=5)
        
        # --- Блок последней строки (БЕЗ ФРЕЙМА) ---
        lbl_row = tk.Label(self.root, text="Номер последней строки:", bg=self.bg_color, font=("Arial", 10), fg="#1B5E20")
        lbl_row.pack(anchor="w", padx=40)
        
        self.last_row_var = tk.StringVar(value="100")
        self.row_entry = tk.Entry(self.root, textvariable=self.last_row_var, width=15, font=("Arial", 10))
        self.row_entry.pack(anchor="w", padx=40, pady=2)

        # --- Блок инструкции (Полупрозрачный) ---
        # Используем Label с переносом текста прямо на фоне, но для читаемости добавим легкую подложку
        instr_label = tk.Label(self.root, text=instruction_text, justify="left", 
                               bg="#FFFFFF", font=("Arial", 9), anchor="w", wraplength=800, padx=10, pady=10)
        # Делаем фон немного прозрачным через атрибут alpha (работает не везде) или просто светлым
        # Для надежности оставляем светло-зеленый, но без рамок
        instr_label.config(bg="#E8F5E9") 
        instr_label.pack(fill="x", padx=20, pady=10)

        # --- Блок шаблона (Полупрозрачный контейнер) ---
        # Используем Frame с очень светлым цветом, чтобы цветы чуть просвечивали, но текст читался
        template_frame = tk.Frame(self.root, bg="#F1F8E9", bd=0) 
        template_frame.pack(fill="both", expand=True, padx=20, pady=5)
        
        tk.Label(template_frame, text="Шаблон текста:", bg="#F1F8E9", font=("Arial", 10, "bold"), fg="#1B5E20").pack(anchor="w")
        
        # Поле ввода белое для контраста
        self.template_text = tk.Text(template_frame, height=12, font=("Consolas", 11), bg="#FFFFFF", fg="#000000")
        self.template_text.pack(fill="both", expand=True, pady=5, padx=2)
        
        # Горячие клавиши и ПКМ
        self.template_text.bind("<Control-v>", lambda e: self.template_text.event_generate("<<Paste>>"))
        self.template_text.bind("<Control-c>", lambda e: self.template_text.event_generate("<<Copy>>"))
        self.template_text.bind("<Control-a>", lambda e: self.template_text.tag_add("sel", "1.0", "end-1c") or "break")
        self.template_text.bind("<Delete>", self.on_delete_key)
        self.template_text.bind("<Button-3>", self.show_context_menu)
        
        # Кнопка генерации
        btn_generate = tk.Button(self.root, text="СГЕНЕРИРОВАТЬ ДОКУМЕНТ", command=self.generate_document,
                                 bg=self.btn_color, fg=self.btn_fg, font=("Arial", 12, "bold"), height=2)
        btn_generate.pack(fill="x", padx=20, pady=10)
        
        # Статус бар (полупрозрачный)
        self.status_var = tk.StringVar(value="Готов к работе")
        status_label = tk.Label(self.root, textvariable=self.status_var, bd=0, anchor="w", bg="#C1E1C1", fg="#1B5E20", font=("Arial", 9))
        status_label.pack(side="bottom", fill="x")

    # Методы контекстного меню
    def show_context_menu(self, event):
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()

    def copy_text(self):
        try:
            self.template_text.event_generate("<<Copy>>")
        except:
            pass

    def paste_text(self):
        try:
            self.template_text.event_generate("<<Paste>>")
        except:
            pass

    def select_all_text(self):
        self.template_text.tag_add("sel", "1.0", "end")
        self.template_text.mark_set("insert", "end")
        self.template_text.see("insert")

    def delete_text(self):
        try:
            if self.template_text.tag_ranges("sel"):
                self.template_text.delete("sel.first", "sel.last")
        except tk.TclError:
            pass

    def on_delete_key(self, event):
        try:
            if self.template_text.tag_ranges("sel"):
                self.template_text.delete("sel.first", "sel.last")
                return "break"
        except tk.TclError:
            pass
        return None

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.file_path_var.set(filename)
            self.status_var.set(f"Файл выбран: {os.path.basename(filename)}")

    def parse_template_meta(self, text):
        pattern = r'\[([A-Z])(\d+)(!)?(?:\|words)?\]'
        matches = re.findall(pattern, text, re.IGNORECASE)
        
        if not matches:
            return None, []
            
        cols = set()
        min_row = float('inf')
        
        for col_letter, row_str, has_exclaim in matches:
            cols.add(col_letter.upper())
            r = int(row_str)
            if r < min_row:
                min_row = r
                
        return min_row, list(cols)

    def generate_document(self):
        file_path = self.file_path_var.get()
        if not file_path:
            messagebox.showerror("Ошибка", "Выберите файл Excel.")
            return

        try:
            last_row_input = int(self.last_row_var.get())
        except ValueError:
            messagebox.showerror("Ошибка", "Номер строки должен быть числом.")
            return

        template_content = self.template_text.get("1.0", "end-1c")
        if not template_content.strip():
            messagebox.showwarning("Внимание", "Введите шаблон текста.")
            return

        meta = self.parse_template_meta(template_content)
        if not meta:
            messagebox.showerror("Ошибка", "Не найдены метки вида [A1], [B2!] и т.д.")
            return
            
        start_row_limit, needed_cols = meta

        try:
            self.status_var.set("Чтение файла и анализ цветов...")
            self.root.update()
            
            wb = load_workbook(file_path, data_only=True)
            ws = wb.active
            
            blocks = {}
            current_block_start = None
            
            for r in range(start_row_limit, last_row_input + 1):
                is_colored = False
                
                for col_letter in needed_cols:
                    col_idx = ord(col_letter) - ord('A') + 1
                    cell = ws.cell(row=r, column=col_idx)
                    fill = cell.fill
                    if fill and fill.fgColor and fill.fgColor.rgb and fill.fgColor.rgb != "00000000":
                         is_colored = True
                         break
                
                if is_colored:
                    current_block_start = r
                    blocks[current_block_start] = [r]
                else:
                    if current_block_start is not None:
                        blocks[current_block_start].append(r)
            
            if not blocks:
                messagebox.showwarning("Результат", f"Не найдено цветных строк (блоков) начиная с строки {start_row_limit}. Проверьте заливку.")
                self.status_var.set("Нет данных")
                return

            results = []
            
            self.status_var.set("Генерация текста...")
            self.root.update()

            for block_start_row, block_rows in blocks.items():
                is_block_valid = True
                all_marks = re.findall(r'\[([A-Z])(\d+)(!)?(\|words)?\]', template_content, re.IGNORECASE)
                
                for col_letter, row_str, has_exclaim, has_words in all_marks:
                    col_letter = col_letter.upper()
                    ref_row = int(row_str)
                    
                    if has_exclaim:
                        found_any = False
                        for r in block_rows:
                            cell_val = ws.cell(row=r, column=ord(col_letter)-ord('A')+1).value
                            if cell_val is not None and str(cell_val).strip() != "":
                                found_any = True
                                break
                        if not found_any:
                            is_block_valid = False
                            break
                            
                    else:
                        offset = ref_row - start_row_limit
                        target_row = block_start_row + offset
                        
                        if target_row > last_row_input:
                            is_block_valid = False
                            break
                        
                        val = ws.cell(row=target_row, column=ord(col_letter)-ord('A')+1).value
                        if val is None or str(val).strip() == "":
                            is_block_valid = False
                            break

                if not is_block_valid:
                    continue

                current_text = template_content
                
                for col_letter, row_str, has_exclaim, has_words in all_marks:
                    col_letter = col_letter.upper()
                    ref_row = int(row_str)
                    
                    if has_exclaim:
                        vals = []
                        for r in block_rows:
                            cell_val = ws.cell(row=r, column=ord(col_letter)-ord('A')+1).value
                            if cell_val is not None and str(cell_val).strip() != "":
                                vals.append(str(cell_val))
                        replacement = ", ".join(vals)
                    else:
                        offset = ref_row - start_row_limit
                        target_row = block_start_row + offset
                        
                        if target_row > last_row_input:
                            replacement = ""
                        else:
                            val = ws.cell(row=target_row, column=ord(col_letter)-ord('A')+1).value
                            if val is None: 
                                replacement = ""
                            elif has_words:
                                replacement = number_to_words_rus(val)
                            else:
                                replacement = str(val)

                    original_mark = f"[{col_letter}{row_str}{'!' if has_exclaim else ''}{'|words' if has_words else ''}]"
                    current_text = current_text.replace(original_mark, replacement)
                
                results.append(current_text)

            if not results:
                messagebox.showwarning("Результат", "Документ пуст.")
                return

            save_path = filedialog.asksaveasfilename(
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx"), ("Text File", "*.txt")],
                title="Сохранить результат"
            )
            
            if save_path:
                self.status_var.set("Сохранение...")
                self.root.update()
                
                if save_path.endswith(".docx"):
                    from docx import Document
                    doc = Document()
                    full_text = "\n\n".join(results)
                    paragraphs = full_text.split('\n\n')
                    for p in paragraphs:
                        if p.strip():
                            doc.add_paragraph(p)
                    doc.save(save_path)
                else:
                    with open(save_path, "w", encoding="utf-8") as f:
                        f.write("\n\n".join(results))
                
                messagebox.showinfo("Успех", f"Документ создан!\nСгенерировано блоков: {len(results)}")
                self.status_var.set("Готово")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать файл: {str(e)}")
            print(e)
            self.status_var.set("Ошибка")

if __name__ == "__main__":
    root = tk.Tk()
    app = DebtGeneratorApp(root)
    root.mainloop()