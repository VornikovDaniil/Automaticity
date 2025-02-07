import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import shutil
import subprocess
import sys
import os
from threading import Thread
import math

def install_packages():
    try:
        import openpyxl
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])

def upload_file():
    filepath = filedialog.askopenfilename()
    if filepath:
        try:
            if not os.path.exists('roob'):
                os.makedirs('roob')
            destination = os.path.join('roob', 'Карты ОПР.docx')
            shutil.copy(filepath, destination)
            messagebox.showinfo("Успех", f"Файл успешно загружен и сохранен как {destination}")
            install_packages()
            progress_bar.start()
            thread = Thread(target=run_script)
            thread.start()
        except Exception as e:
            messagebox.showerror("Ошибка", f"Произошла ошибка: {e}")

def run_script():
    try:
        subprocess.run([sys.executable, 'work.py'])
        progress_bar.stop()
        messagebox.showinfo("Успех", "Скрипт успешно выполнен")
        show_generated_files()
    except Exception as e:
        progress_bar.stop()
        messagebox.showerror("Ошибка", f"Произошла ошибка при выполнении скрипта: {e}")

def show_generated_files():
    files_window = tk.Toplevel(root)
    files_window.title("Сгенерированные файлы")
    files_window.geometry("600x400")
    files_window.configure(bg="black")

    if not os.path.exists('gotfill'):
        messagebox.showinfo("Информация", "Папка gotfill не существует")
        return

    files = os.listdir('gotfill')
    if not files:
        messagebox.showinfo("Информация", "Нет файлов в папке gotfill")
        return

    download_all_button = ttk.Button(files_window, text="Скачать все", command=lambda: download_all_files(files))
    download_all_button.pack(pady=20, padx=20, fill=tk.X)

def download_all_files(files):
    folder_path = filedialog.askdirectory()
    if folder_path:
        for file in files:
            file_path = os.path.join('gotfill', file)
            shutil.copy(file_path, os.path.join(folder_path, file))
        messagebox.showinfo("Успех", f"Все файлы были сохранены в {folder_path}")

def update_line(canvas, line_id, angle_offset):
    width, height = int(canvas.cget('width')), int(canvas.cget('height'))
    angle = math.radians(angle_offset)
    x1, y1 = width / 2 + 300 * math.sin(angle), height / 2 + 300 * math.cos(angle)
    x2, y2 = width / 2 - 300 * math.sin(angle), height / 2 - 300 * math.cos(angle)
    canvas.coords(line_id, x1, y1, x2, y2)
    canvas.after(50, update_line, canvas, line_id, angle_offset + 2)

def create_moving_lines(root):
    canvas = tk.Canvas(root, width=600, height=400)
    canvas.place(x=0, y=0, relwidth=1, relheight=1)
    lines = []
    for _ in range(10):
        line = canvas.create_line(0, 0, 0, 0, fill="#ccc", width=2)
        lines.append(line)
    for i, line in enumerate(lines):
        update_line(canvas, line, i * 36)  # Different initial angles for variety

root = tk.Tk()
root.title("Загрузка файла и запуск скрипта")
root.geometry("600x400")
create_moving_lines(root)

style = ttk.Style()
style.configure("TButton", padding=6, relief="flat", background="#ccc")
style.configure("TFrame", background="transparent")
style.configure("TLabel", background="transparent", foreground="white")

frame = ttk.Frame(root, padding="20")
frame.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

title_label = ttk.Label(frame, text="Добро пожаловать!", font=("Helvetica", 16))
title_label.pack(pady=10)

upload_button = ttk.Button(frame, text="Загрузить файл", command=upload_file)
upload_button.pack(pady=20)

progress_bar = ttk.Progressbar(frame, mode="indeterminate")
progress_bar.pack(pady=20, fill=tk.X)

root.mainloop()
