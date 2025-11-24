import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import pandas as pd
import time
import urllib.parse
import random
import math
from openpyxl.styles import Font, Alignment
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import sys
import os
import csv
import json

class AvitoParserGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Парсер объявлений Авито")
        
        # Получаем размеры экрана
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Устанавливаем окно на весь экран
        self.root.geometry(f"{screen_width}x{screen_height}")
        self.root.state('zoomed')
        
        self.root.configure(bg='#f5f5f5')
        
        self.is_running = False
        self.save_path = ""
        self.current_progress = 0
        self.max_links = 500
        
        self.setup_ui()
        
    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure('Main.TFrame', background='#f5f5f5')
        style.configure('Title.TLabel', background='#f5f5f5', font=('Arial', 16, 'bold'), foreground='#2c3e50')
        style.configure('Subtitle.TLabel', background='#f5f5f5', font=('Arial', 10), foreground='#7f8c8d')
        style.configure('TButton', font=('Arial', 10), padding=6)
        style.configure('Action.TButton', background='#27ae60', foreground='white')
        style.map('Action.TButton', background=[('active', '#219653')])
        style.configure('Stop.TButton', background='#e74c3c', foreground='white')
        style.map('Stop.TButton', background=[('active', '#c0392b')])
        style.configure('Custom.Horizontal.TProgressbar', troughcolor='#ecf0f1', background='#27ae60', thickness=20)
        
        main_frame = ttk.Frame(self.root, style='Main.TFrame', padding="25")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        header_frame = ttk.Frame(main_frame, style='Main.TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(header_frame, text="Парсер объявлений Авито", style='Title.TLabel')
        title_label.pack(anchor=tk.CENTER)
        
        subtitle_label = ttk.Label(header_frame, text="Сбор данных по ключевым словам", style='Subtitle.TLabel')
        subtitle_label.pack(anchor=tk.CENTER, pady=(5, 0))
        
        info_frame = ttk.LabelFrame(main_frame, text="Информация", padding=15)
        info_frame.pack(fill=tk.X, pady=10)
        
        info_text = "Программа собирает данные с сайта Авито по вашим ключевым словам.\nВыберите папку для сохранения, формат файла и количество записей."
        ttk.Label(info_frame, text=info_text, font=('Arial', 9), background='#f5f5f5', justify=tk.CENTER).pack()
        
        settings_frame = ttk.Frame(main_frame, style='Main.TFrame')
        settings_frame.pack(fill=tk.X, pady=15)
        
        ttk.Label(settings_frame, text="Ключевые слова:").grid(row=0, column=0, sticky=tk.W, pady=8)
        self.search_var = tk.StringVar(value="нутрициолог")
        self.search_entry = ttk.Entry(settings_frame, textvariable=self.search_var, font=('Arial', 11), width=30)
        self.search_entry.grid(row=0, column=1, sticky=tk.W, padx=(10, 0), pady=8)
        
        ttk.Label(settings_frame, text="Количество объявлений:").grid(row=1, column=0, sticky=tk.W, pady=8)
        self.max_links_var = tk.StringVar(value="500")
        self.max_links_entry = ttk.Entry(settings_frame, textvariable=self.max_links_var, font=('Arial', 11), width=10)
        self.max_links_entry.grid(row=1, column=1, sticky=tk.W, padx=(10, 0), pady=8)
        
        ttk.Label(settings_frame, text="Формат файла:").grid(row=2, column=0, sticky=tk.W, pady=8)
        self.format_var = tk.StringVar(value="excel")
        format_combo = ttk.Combobox(settings_frame, textvariable=self.format_var, 
                                   values=["excel", "csv", "json"], state="readonly", width=15)
        format_combo.grid(row=2, column=1, sticky=tk.W, padx=(10, 0), pady=8)
        
        ttk.Label(settings_frame, text="Папка сохранения:").grid(row=3, column=0, sticky=tk.W, pady=8)
        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(settings_frame, textvariable=self.path_var, state='readonly', font=('Arial', 10))
        self.path_entry.grid(row=3, column=1, sticky=tk.EW, padx=(10, 0), pady=8)
        
        ttk.Button(settings_frame, text="Обзор", command=self.select_folder, width=8).grid(row=3, column=2, padx=(5, 0), pady=8)
        
        settings_frame.columnconfigure(1, weight=1)
        
        progress_frame = ttk.LabelFrame(main_frame, text="Прогресс выполнения", padding=15)
        progress_frame.pack(fill=tk.X, pady=15)
        
        self.progress_label = ttk.Label(progress_frame, text="Готов к работе", font=('Arial', 10), background='#f5f5f5')
        self.progress_label.pack(anchor=tk.CENTER, pady=(0, 8))
        
        self.progress = ttk.Progressbar(progress_frame, style='Custom.Horizontal.TProgressbar', mode='determinate', maximum=100)
        self.progress.pack(fill=tk.X)
        
        self.percentage_label = ttk.Label(progress_frame, text="0%", font=('Arial', 10, 'bold'), background='#f5f5f5')
        self.percentage_label.pack(anchor=tk.CENTER, pady=(5, 0))
        
        control_frame = ttk.Frame(main_frame, style='Main.TFrame')
        control_frame.pack(fill=tk.X, pady=10)
        
        self.start_button = ttk.Button(control_frame, text="Начать сбор данных", command=self.toggle_parser, style='Action.TButton')
        self.start_button.pack(anchor=tk.CENTER)
        
    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.save_path = folder
            self.path_var.set(folder)
            
    def toggle_parser(self):
        if not self.is_running:
            self.start_parser()
        else:
            self.stop_parser()
            
    def start_parser(self):
        if not self.save_path:
            messagebox.showerror("Ошибка", "Выберите папку для сохранения результатов")
            return
            
        if not self.search_var.get().strip():
            messagebox.showerror("Ошибка", "Введите ключевые слова для поиска")
            return
            
        try:
            self.max_links = int(self.max_links_var.get())
            if self.max_links <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Ошибка", "Введите корректное число записей")
            return
            
        self.is_running = True
        self.current_progress = 0
        self.start_button.configure(text="Остановить", style='Stop.TButton')
        self.progress['value'] = 0
        self.percentage_label['text'] = "0%"
        self.progress_label['text'] = f"Запуск поиска по запросу: {self.search_var.get()}"
        
        thread = threading.Thread(target=self.run_parser)
        thread.daemon = True
        thread.start()
        
    def stop_parser(self):
        self.is_running = False
        self.start_button.configure(text="Начать сбор данных", style='Action.TButton')
        self.progress_label['text'] = "Сбор данных остановлен"
        
    def run_parser(self):
        try:
            data = get_avito_data_selenium(
                self.search_var.get(), 
                self.max_links, 
                self.save_path, 
                self.format_var.get(),
                self.update_progress
            )
            if data:
                self.update_progress(f"Завершено! Собрано {len(data)} объявлений", 100)
            else:
                self.update_progress("Сбор завершен без данных", 100)
        except Exception as e:
            self.update_progress(f"Ошибка: {str(e)}", 0)
        finally:
            self.is_running = False
            self.root.after(0, self.stop_parser)
            
    def update_progress(self, message, progress=None):
        if not self.is_running:
            return
            
        def update():
            self.progress_label['text'] = message
            if progress is not None:
                self.current_progress = progress
                self.progress['value'] = progress
                self.percentage_label['text'] = f"{int(progress)}%"
        
        self.root.after(0, update)

def get_avito_data_selenium(search_query, max_links=500, save_path="", file_format="excel", callback=None):
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1200,800")
    
    driver = webdriver.Chrome(options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    try:
        product_data = []
        page = 1
        captcha_count = 0
        consecutive_empty_pages = 0
        
        items_per_page = 50
        max_pages = math.ceil(max_links / items_per_page) + 5
        
        if callback:
            callback("Инициализация парсера...", 5)
        
        encoded_query = urllib.parse.quote(search_query)
        
        while len(product_data) < max_links and page <= max_pages and captcha_count < 5:
            if page == 1:
                url = f"https://www.avito.ru/all?q={encoded_query}"
            else:
                url = f"https://www.avito.ru/all?p={page}&q={encoded_query}"
            
            if callback:
                progress = min(5 + (page * 90 / max_pages), 95)
                callback(f"Поиск '{search_query}', страница {page}", progress)
            
            driver.get(url)
            time.sleep(3)
            
            if check_captcha_improved(driver):
                captcha_count += 1
                if callback:
                    callback(f"Обнаружена капча (попытка {captcha_count}/5)", 0)
                
                time.sleep(10)
                
                retry_count = 0
                while retry_count < 30:
                    if not check_captcha_improved(driver):
                        if callback:
                            callback("Капча решена, продолжаем...", progress)
                        break
                    time.sleep(10)
                    retry_count += 1
                continue
            else:
                captcha_count = 0
            
            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "[data-marker='item']"))
                )
            except TimeoutException:
                if callback:
                    callback("Не удалось загрузить список объявлений", progress)
                if check_captcha_improved(driver):
                    captcha_count += 1
                    continue
                else:
                    consecutive_empty_pages += 1
                    if consecutive_empty_pages >= 2:
                        if callback:
                            callback("Завершение: пустые страницы", 100)
                        break
                    else:
                        page += 1
                        continue
            
            items = driver.find_elements(By.CSS_SELECTOR, "[data-marker='item']")
            
            if not items or len(items) < 5:
                consecutive_empty_pages += 1
                if consecutive_empty_pages >= 2:
                    if callback:
                        callback("Завершение: объявления не найдены", 100)
                    break
                else:
                    page += 1
                    continue
            else:
                consecutive_empty_pages = 0
            
            new_items_count = 0
            for item in items:
                if len(product_data) >= max_links:
                    break
                
                try:
                    link_element = item.find_element(By.CSS_SELECTOR, "a[data-marker='item-title']")
                    href = link_element.get_attribute('href')
                    
                    if not href:
                        continue
                    
                    if any(data['Ссылка'] == href for data in product_data):
                        continue
                    
                    title = link_element.get_attribute('title') or link_element.text
                    
                    price = "0"
                    try:
                        price_element = item.find_element(By.CSS_SELECTOR, "[data-marker='item-price']")
                        price_text = price_element.text.strip()
                        price = clean_price(price_text)
                    except:
                        pass
                    
                    seller_name = "Не указано"
                    try:
                        # Правильный селектор для продавца
                        seller_element = item.find_element(By.CSS_SELECTOR, "p.styles-module-root-PY1ie.styles-module-size_m-w6vzl.styles-module-size_m_dense-HvBLt.styles-module-size_m_compensated-a0qNK.styles-module-size_m-DKJW6.styles-module-ellipsis-HCaiF.styles-module-ellipsis_oneLine-VXBA3.styles-module-size_dense-u0sRJ.stylesMarningNormal-module-root-OE0X2.stylesMarningNormal-module-paragraph-m-dense-mYuSK")
                        seller_name = seller_element.text.strip()
                    except:
                        try:
                            seller_elements = item.find_elements(By.CSS_SELECTOR, "[class*='styles-module-root']")
                            for elem in seller_elements:
                                text = elem.text.strip()
                                if text and len(text) > 1 and len(text) < 50:
                                    seller_name = text
                                    break
                        except:
                            pass
                    
                    rating = "0"
                    try:
                        # Правильный селектор для рейтинга
                        rating_element = item.find_element(By.CSS_SELECTOR, "div.styles-module-root-Sd1q7 span[data-marker='seller-info/score']")
                        rating_text = rating_element.text.strip()
                        if rating_text and any(char.isdigit() for char in rating_text):
                            rating = rating_text.replace(',', '.')
                    except:
                        pass
                    
                    product_record = {
                        'Название': title,
                        'Продавец': seller_name,
                        'Цена': price,
                        'Рейтинг': rating,
                        'Ссылка': href
                    }
                    
                    product_data.append(product_record)
                    new_items_count += 1
                    
                    current_progress = min(5 + (len(product_data) * 90 / max_links), 95)
                    if callback and len(product_data) % 10 == 0:
                        callback(f"Найдено {len(product_data)} объявлений", current_progress)
                    
                except Exception:
                    continue
            
            should_continue = (
                len(product_data) < max_links and 
                new_items_count > 0 and 
                page < max_pages and
                consecutive_empty_pages == 0
            )
            
            if should_continue:
                page += 1
                delay = random.uniform(2, 4)
                time.sleep(delay)
                
                if page % 5 == 0:
                    long_pause = random.uniform(8, 12)
                    time.sleep(long_pause)
            else:
                break
        
        if callback:
            callback(f"Сбор завершен! Найдено: {len(product_data)} объявлений", 100)
        
        save_avito_data(product_data, search_query, save_path, file_format, callback)
        return product_data
        
    except Exception as e:
        if callback:
            callback(f"Ошибка выполнения: {str(e)}", 0)
        if 'product_data' in locals():
            save_avito_data(product_data, search_query, save_path, file_format, callback)
        return []
    finally:
        driver.quit()

def save_avito_data(product_data, search_query, save_path, file_format="excel", callback=None):
    if not product_data:
        if callback:
            callback("Нет данных для сохранения", 100)
        return
    
    if callback:
        callback(f"Сохранение в {file_format.upper()}...", 95)
    
    df = pd.DataFrame(product_data)
    
    df['Цена_число'] = pd.to_numeric(df['Цена'], errors='coerce').fillna(0)
    df['Цена_отображение'] = df['Цена_число'].apply(lambda x: f"{x:,.0f} ₽" if x > 0 else "Не указана")
    
    # Очищаем запрос для имени файла
    clean_query = "".join(c for c in search_query if c.isalnum() or c in (' ', '-', '_')).rstrip()
    if not clean_query:
        clean_query = "search_results"
    else:
        clean_query = clean_query.replace(' ', '_')[:50]
    
    timestamp = time.strftime('%Y%m%d_%H%M')
    
    try:
        if file_format == "excel":
            filename = f"avito_{clean_query}_{timestamp}.xlsx"
            full_path = f"{save_path}/{filename}"
            
            with pd.ExcelWriter(full_path, engine='openpyxl') as writer:
                df[['Название', 'Продавец', 'Цена_отображение', 'Рейтинг', 'Ссылка']].to_excel(
                    writer, index=False, sheet_name='Объявления'
                )
                
                workbook = writer.book
                worksheet = writer.sheets['Объявления']
                
                column_widths = {
                    'A': 40,  # Название
                    'B': 25,  # Продавец
                    'C': 15,  # Цена
                    'D': 15,  # Рейтинг
                    'E': 50   # Ссылка
                }
                
                for col, width in column_widths.items():
                    worksheet.column_dimensions[col].width = width
                
                # ДОБАВЛЯЕМ АКТИВНЫЕ ССЫЛКИ В EXCEL
                for row in range(2, len(df) + 2):
                    cell = worksheet[f'E{row}']  # Колонка E - ссылки
                    url = df.iloc[row-2]['Ссылка']
                    if url and isinstance(url, str) and url.startswith('http'):
                        cell.hyperlink = url
                        cell.font = Font(color="0000FF", underline="single")
                        cell.alignment = Alignment(horizontal='left')
                
                # Форматируем ячейки
                for row in worksheet.iter_rows(min_row=2, max_row=len(df)+1, min_col=1, max_col=5):
                    for cell in row:
                        if cell.column_letter in ['A', 'B']:
                            cell.alignment = Alignment(wrap_text=True, vertical='top')
                        elif cell.column_letter in ['C', 'D']:
                            cell.alignment = Alignment(horizontal='right', vertical='top')
            
        elif file_format == "csv":
            filename = f"avito_{clean_query}_{timestamp}.csv"
            full_path = f"{save_path}/{filename}"
            
            # Сохраняем в CSV
            df[['Название', 'Продавец', 'Цена_отображение', 'Рейтинг', 'Ссылка']].to_csv(
                full_path, index=False, encoding='utf-8-sig', sep=';'
            )
            
        elif file_format == "json":
            filename = f"avito_{clean_query}_{timestamp}.json"
            full_path = f"{save_path}/{filename}"
            
            # Подготавливаем данные для JSON
            json_data = []
            for _, row in df.iterrows():
                item = {
                    'Название': row['Название'],
                    'Продавец': row['Продавец'],
                    'Цена': row['Цена_отображение'],
                    'Рейтинг': row['Рейтинг'],
                    'Ссылка': row['Ссылка']
                }
                json_data.append(item)
            
            # Сохраняем в JSON
            with open(full_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
        
        if callback:
            callback(f"Файл сохранен: {filename}", 100)
        
    except Exception as e:
        if callback:
            callback(f"Ошибка сохранения: {str(e)}", 100)

def check_captcha_improved(driver):
    try:
        captcha_selectors = [
            "img[src*='captcha']",
            "iframe[src*='recaptcha']", 
            "div[class*='captcha']",
            "div[class*='recaptcha']",
        ]
        
        for selector in captcha_selectors:
            elements = driver.find_elements(By.CSS_SELECTOR, selector)
            for element in elements:
                if element.is_displayed():
                    return True
        
        page_text = driver.page_source.lower()
        exact_captcha_phrases = [
            'введите текст с картинки',
            'введите код с картинки', 
            'проверка безопасности',
        ]
        
        for phrase in exact_captcha_phrases:
            if phrase in page_text:
                return True
        
        return False
        
    except:
        return False

def clean_price(price_str):
    if not price_str or price_str == "Не указана":
        return "0"
    
    try:
        cleaned = ''.join(c for c in str(price_str) if c.isdigit())
        return cleaned if cleaned else "0"
    except:
        return "0"

def main():
    root = tk.Tk()
    app = AvitoParserGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()