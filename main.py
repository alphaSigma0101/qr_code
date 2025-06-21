import sqlite3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from PIL import Image, ImageDraw, ImageFont, ImageWin
import threading
import json
import win32print
import win32ui
import os
import requests
from io import BytesIO
import base64
import win32com.client
from pyzbar.pyzbar import decode
import cv2

# Для генерации DataMatrix
try:
    from pylibdmtx.pylibdmtx import encode
except ImportError:
    encode = None

# Для печати (ESC/POS)
try:
    from escpos.printer import Usb
except ImportError:
    Usb = None


class DataMatrixPrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("DataMatrix Printer (TCS TDP-225)")
        self.root.geometry("900x800")
        self.root.resizable(False, False)
        
        self.current_images = []
        self.printer = None
        self.selected_printer_info = None
        self.db_lock = threading.Lock()
        self.logo_image = None
        self.brand_logo = None
        self.scanner_device = None
        
        self.setup_database()
        self.load_printer_settings()
        self.setup_ui()
        self.check_dependencies()
        self.load_brand_logo()
        self.load_logo()
        self.detect_scanner()
        
        if not self.selected_printer_info:
            self.setup_printer_selection()
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def detect_scanner(self):
        """Автоматическое обнаружение 2D-сканера"""
        try:
            wmi = win32com.client.GetObject("winmgmts:")
            devices = wmi.InstancesOf("Win32_PnPEntity")
            
            scanner_names = ["2D Scanner", "Barcode Scanner", "Symbol", "Honeywell", "Datalogic"]
            for device in devices:
                if device.Name:
                    for name in scanner_names:
                        if name in device.Name:
                            self.scanner_device = device.Name
                            messagebox.showinfo("Сканер обнаружен", 
                                              f"Найден 2D-сканер: {device.Name}")
                            return
            
            messagebox.showwarning("Сканер не найден", 
                                 "2D-сканер не обнаружен. Проверьте подключение.")
        except Exception as e:
            print(f"Ошибка обнаружения сканера: {e}")

    def load_brand_logo(self):
        """Загрузка официального логотипа 'Честный знак' (встроенный)"""
        try:
            # Base64-кодированное изображение логотипа (улучшенная версия)
            logo_base64 = """
            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
            5D8B4gLfzOQ/AeIC383kPwHiAj/M5D8B4gI/zeQ/AeICv8zkPwHiAr/N5D8B4gJ/zOQ/AeICf83k
            PwHiAv/M5D8B4gL/zeQ/AeICo5nJfwLEBcYyk/8EiAuMZyb/CRAXmMhM/hMgLjCZmfwnQFxgKjP5
            T4C4wHRm8p8AcYEwM/lPgLhAlJn8J0BcIM5M/hMgLpBkJv8JEBdIM5P/BIgLZJnJfwLEBfLM5D8B
            4gJFZvKfAHGBMjP5T4C4QJWZ/CdAXKDOTP4TIC7QZCb/CRAXaDGT/wSIC7SZyX8CxAXazeQ/AeIC
            HWbynwBxgU4z+U+AuECXmfwnQFyg20z+EyAu0GMm/wkQF1hgJv8JEBdYaCb/CRAXWGQm/wkQF1hs
            Jv8JEBdYYib/CRAXWGom/wkQF1hmJv8JEBdYbib/CRAXWGEm/wkQF1hpJv8JEBdYZSb/CRAXWG0m
            /wkQF1hjJv8JEBdYayb/CRAXWGcm/wkQF1hvJv8JEBfYYCb/CRAX2Ggm/wkQF9hkJv8JEBfYbCb/

            CRAX2GIm/wkQF9hqJv8JEBfYZib/CRAX2G4m/wkQF9hhJv8JEBfYaSb/CRAX2GUm/wkQF9htJv8J
            EBfYYyb/CRAX2Gsm/wkQF9hnJv8JEBfYbyb/CRAXOJCZ/CdAXOBgZvKfAHGBQ5nJfwLEBQ5nJv8J
            EBc4kpn8J0Bc4Ghm8p8AcYFjmcl/AsQFjmcm/wkQFziRmfwnQFzgZGbynwBxgVOZyX8CxAVOZyb/
            CRAXOJOZ/CdAXOBsZvKfAHGBc5nJfwLEBc5nJv8JEBfozUz+EyAu0JeZ/CdAXKA/M/lPgLjAhczk
            PwHiAhezkv8EiAtcykr+EyAucDkr+U+AuMCVrOQ/AeICV7OS/wSIC1zLSv4TIC5wPSv5T4C4wI2s
            5D8B4gI3s5L/BIgL3MpK/hMgLnA7K/lPgLjAnazkPwHiAnez8v8EiAvcy8r/EyAucD8r/0+AuMCD
            rPw/AeICD7Py/wSICzzKyv8TIC7wOCv/T4C4wJOs/D8B4gJPs/L/BIgLPMvK/xMgLvA8K/9PgLjA
            i6z8PwHiAi+z8v8EiAu8ysr/EyAu8Dor/0+AuMCbrPw/AeICb7Py/wSIC7zNyv8TIC7wLiv/T4C4
            wPus/D8B4gIfsvL/BIgLfMzK/xMgLvApK/9PgLjA56z8PwHiAl+y8v8EiAt8zcr/EyAu8C0r/0+A
            uMD3rPw/AeICP7Ly/wSIC/zMyv8TIC7wKyv/T4C4wO+s/D8B4gJ/svL/BIgL/M3K/xMgLjCSlf8n
            QFxgNCv/T4C4wFhW/p8AcYHxrPw/AeICE1n5fwLEBSaz8v8EiAtMZeX/CRAXmM7K/xMgLjCTlf8n
            QFxgNiv/T4C4QJCV/ydAXCDMyv8TIC4QZeX/CRAXiLPy/wSICyRZ+X8CxAXSrPw/AeICWVb+nwBx
            gTwr/0+AuECRlf8nQFygyMr/EyAuUGbl/wkQF6iy8v8EiAvUWfl/AsQFmqz8PwHiAk1W/p8AcYE2
            K/9PgLhAm5X/J0BcoMvK/xMgLtBl5f8JEBfosvL/BIgLdFn5fwLEBbqs/D8B4gJdVv6fAHGBLiv/

            T4C4QJeV/ydAXKDLyv8TIC7QZeX/CRAX6LLy/wSIC3RZ+X8CxAW6rPw/AeICXVb+nwBxgS4r/wf+
            A4j9XqVr3X3/AAAAAElFTkSuQmCC
            """
            logo_data = base64.b64decode(logo_base64.strip())
            self.brand_logo = Image.open(BytesIO(logo_data))
            # Масштабируем логотип до подходящего размера
            self.brand_logo = self.brand_logo.resize((40, 40), Image.Resampling.LANCZOS)
        except Exception as e:
            print(f"Ошибка загрузки логотипа: {e}")
            self.brand_logo = None

    def load_logo(self):
        """Создание текстового логотипа 'Честный знак' с графическим лого"""
        try:
            width = 224
            logo_width = width // 3 * 2
            logo_height = 50  # Увеличили высоту для графического логотипа
            
            logo_img = Image.new('RGB', (logo_width, logo_height), color='white')
            draw = ImageDraw.Draw(logo_img)
            
            try:
                font = ImageFont.truetype("arialbd.ttf", 16)
            except:
                try:
                    font = ImageFont.truetype("arial.ttf", 16)
                except:
                    font = ImageFont.load_default()
            
            text = "ЧЕСТНЫЙ ЗНАК"
            
            # Если есть графический логотип, добавляем его перед текстом
            if self.brand_logo:
                # Позиционируем графический логотип слева от текста
                logo_img.paste(self.brand_logo, (10, (logo_height - self.brand_logo.height) // 2))
                text_x = 10 + self.brand_logo.width + 10  # Отступ после логотипа
            else:
                text_x = (logo_width - draw.textlength(text, font=font)) // 2
            
            text_y = (logo_height - font.size) // 2
            draw.text((text_x, text_y), text, fill='black', font=font)
            
            underline_y = text_y + font.size + 2
            draw.line((text_x, underline_y, text_x + draw.textlength(text, font=font), underline_y), 
                     fill='black', width=1)
            
            self.logo_image = logo_img
        except Exception as e:
            print(f"Ошибка создания логотипа: {e}")
            self.logo_image = None

    def generate_datamatrix(self, code):
        """Генерация изображения DataMatrix с кодом и логотипом"""
        PAPER_WIDTH_PX = 224
        TARGET_DM_SIZE = 118
        
        if encode is None:
            image = Image.new('RGB', (PAPER_WIDTH_PX, 300), color='white')
            draw = ImageDraw.Draw(image)
            draw.text((10, 10), f"DataMatrix:\n{code}\nЧестный знак", fill='black')
            return image
            
        try:
            encoded = encode(code.encode('utf-8'))
            dm_image = Image.frombytes('RGB', (encoded.width, encoded.height), encoded.pixels)
            
            scale = TARGET_DM_SIZE / max(dm_image.size)
            new_size = (int(dm_image.size[0] * scale), int(dm_image.size[1] * scale))
            dm_image = dm_image.resize(new_size, Image.Resampling.LANCZOS)
            
            padding_top = 15
            padding_bottom = 10
            code_font_size = 16
            
            try:
                code_font = ImageFont.truetype("arialbd.ttf", code_font_size)
            except:
                try:
                    code_font = ImageFont.truetype("arial.ttf", code_font_size)
                except:
                    code_font = ImageFont.load_default()
            
            dummy_draw = ImageDraw.Draw(Image.new('RGB', (1, 1)))
            code_width = int(dummy_draw.textlength(code, font=code_font))
            
            max_width = max(dm_image.size[0], code_width, self.logo_image.width if self.logo_image else 0)
            if max_width > PAPER_WIDTH_PX:
                scale_factor = PAPER_WIDTH_PX / max_width
                dm_image = dm_image.resize(
                    (int(dm_image.size[0] * scale_factor), 
                     int(dm_image.size[1] * scale_factor)), 
                    Image.Resampling.LANCZOS
                )
                code_font_size = int(code_font_size * scale_factor)
                try:
                    code_font = ImageFont.truetype("arialbd.ttf", code_font_size)
                except:
                    try:
                        code_font = ImageFont.truetype("arial.ttf", code_font_size)
                    except:
                        code_font = ImageFont.load_default()
            
            code_height = code_font.getbbox(code)[3] + 5
            logo_height = self.logo_image.height if self.logo_image else 0
            
            total_height = (padding_top + dm_image.size[1] + padding_bottom + 
                           code_height + padding_bottom + logo_height)
            
            img = Image.new('RGB', (PAPER_WIDTH_PX, total_height), 'white')
            draw = ImageDraw.Draw(img)
            
            current_y = padding_top
            
            dm_x = (PAPER_WIDTH_PX - dm_image.size[0]) // 2
            img.paste(dm_image, (dm_x, current_y))
            current_y += dm_image.size[1] + padding_bottom
            
            code_x = (PAPER_WIDTH_PX - code_width) // 2
            draw.text((code_x, current_y), code, fill='black', font=code_font)
            current_y += code_height + padding_bottom
            
            if self.logo_image:
                logo_x = (PAPER_WIDTH_PX - self.logo_image.width) // 2
                img.paste(self.logo_image, (logo_x, current_y))
            
            return img
            
        except Exception as e:
            messagebox.showerror("Ошибка генерации", f"Ошибка: {str(e)}")
            return None

    # Остальные методы класса остаются без изменений
    # ...


if __name__ == "__main__":
    root = tk.Tk()
    app = DataMatrixPrinterApp(root)
    root.mainloop()

    def print_images(self, images, copies=1):
        if not self.selected_printer_info:
            messagebox.showwarning("Ошибка", "Принтер не выбран!")
            return
            
        try:
            printer_name = self.selected_printer_info.get("printer_name")
            
            if printer_name:
                for i in range(copies):
                    for img in images:
                        img = img.convert('1')
                        self.print_image_to_windows_printer(img, printer_name)
                        if i < copies - 1 or img != images[-1]:
                            self.print_feed(3)
            
            messagebox.showinfo("Успех", "Печать завершена")
        except Exception as e:
            messagebox.showerror("Ошибка печати", str(e))

    def print_feed(self, lines=3):
        """Подача пустых строк"""
        if self.selected_printer_info.get("printer_name"):
            hprinter = win32print.OpenPrinter(self.selected_printer_info["printer_name"])
            try:
                hdc = win32ui.CreateDC()
                hdc.CreatePrinterDC(self.selected_printer_info["printer_name"])
                hdc.StartDoc("Feed")
                hdc.StartPage()
                hdc.EndPage()
                hdc.EndDoc()
            finally:
                win32print.ClosePrinter(hprinter)

    def print_image_to_windows_printer(self, img, printer_name):
        """Печать изображения через Windows API"""
        try:
            hprinter = win32print.OpenPrinter(printer_name)
            try:
                hdc = win32ui.CreateDC()
                hdc.CreatePrinterDC(printer_name)
                hdc.StartDoc("DataMatrix Print")
                hdc.StartPage()
                
                dib = ImageWin.Dib(img)
                width, height = img.size
                printable_area = hdc.GetDeviceCaps(8), hdc.GetDeviceCaps(10)
                scale = min(printable_area[0] / width, printable_area[1] / height) * 0.9
                new_width = int(width * scale)
                new_height = int(height * scale)
                x_offset = (printable_area[0] - new_width) // 2
                y_offset = (printable_area[1] - new_height) // 2
                
                dib.draw(hdc.GetHandleOutput(), (x_offset, y_offset, x_offset + new_width, y_offset + new_height))
                
                hdc.EndPage()
                hdc.EndDoc()
            finally:
                win32print.ClosePrinter(hprinter)
        except Exception as e:
            raise Exception(f"Ошибка печати: {str(e)}")

    def generate_and_print(self):
        codes = self.get_codes_from_input()
        if not codes:
            return
            
        try:
            copies = int(self.copies_spin.get())
            if copies < 1 or copies > 100:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Ошибка", "Некорректное количество копий (1-100)")
            return
            
        self.current_images = []
        for code in codes:
            img = self.generate_datamatrix(code)
            if img:
                self.current_images.append(img)
            
        if self.current_images:
            self.save_to_database(codes)
            self.print_images(self.current_images, copies)

    def save_to_database(self, codes):
        conn = self.get_db_connection()
        try:
            with self.db_lock:
                cursor = conn.cursor()
                for code in codes:
                    cursor.execute("INSERT INTO codes (code) VALUES (?)", (code,))
                conn.commit()
                self.load_history()
        except Exception as e:
            messagebox.showerror("Ошибка БД", str(e))
        finally:
            conn.close()

    def get_db_connection(self):
        return sqlite3.connect('datamatrix_codes.db', check_same_thread=False)

    def load_history(self):
        conn = self.get_db_connection()
        try:
            cursor = conn.cursor()
            self.history_listbox.delete(0, tk.END)
            cursor.execute("SELECT code FROM codes ORDER BY print_time DESC LIMIT 100")
            rows = cursor.fetchall()
            for row in rows:
                self.history_listbox.insert(tk.END, row[0])
        finally:
            conn.close()

    def reprint_selected(self):
        selected = self.history_listbox.curselection()
        if not selected:
            messagebox.showwarning("Ошибка", "Выберите коды для повторной печати")
            return
            
        try:
            copies = int(self.copies_spin.get())
            if copies < 1 or copies > 100:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Ошибка", "Некорректное количество копий (1-100)")
            return
            
        codes = [self.history_listbox.get(i) for i in selected]
        self.current_images = []
        for code in codes:
            img = self.generate_datamatrix(code)
            if img:
                self.current_images.append(img)
            
        if self.current_images:
            self.print_images(self.current_images, copies)

    def export_history(self):
        """Экспорт выбранных кодов в PNG"""
        selected = self.history_listbox.curselection()
        if not selected:
            messagebox.showwarning("Ошибка", "Выберите коды для экспорта!")
            return

        codes = [self.history_listbox.get(i) for i in selected]
        folder_path = filedialog.askdirectory(title="Выберите папку для сохранения")
        if not folder_path:
            return

        success_count = 0
        for code in codes:
            try:
                image = self.generate_datamatrix(code)
                if not image:
                    continue
                
                filename = f"datamatrix_{code}.png".replace('\\', '_').replace('/', '_')
                filepath = os.path.join(folder_path, filename)
                image.save(filepath, "PNG")
                success_count += 1
                
            except Exception as e:
                print(f"Ошибка при сохранении {code}: {e}")

        messagebox.showinfo(
            "Готово",
            f"Сохранено {success_count} из {len(codes)} кодов в папку:\n{folder_path}"
        )


if __name__ == "__main__":
    root = tk.Tk()
    app = DataMatrixPrinterApp(root)
    root.mainloop()