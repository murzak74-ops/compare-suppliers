"""Tkinter based GUI for batch document processing."""

from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox
from typing import List

from exporter import export_to_excel
from ocr_utils import extract_text_from_file
from parser import parse_items


class Application(tk.Frame):
    """Simple GUI application."""

    def __init__(self, master: tk.Tk | None = None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.files: List[str] = []
        self.items = []
        self.create_widgets()

    def create_widgets(self) -> None:
        self.load_btn = tk.Button(self, text="Загрузить файлы", command=self.load_files)
        self.load_btn.pack(padx=5, pady=5)

        self.process_btn = tk.Button(self, text="Обработать", command=self.process_files)
        self.process_btn.pack(padx=5, pady=5)

        self.export_btn = tk.Button(self, text="Экспорт в Excel", command=self.export)
        self.export_btn.pack(padx=5, pady=5)

        self.file_list = tk.Listbox(self, width=50)
        self.file_list.pack(padx=5, pady=5)

    def load_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Выберите файлы",
            filetypes=[
                ("Изображения и PDF", "*.png *.jpg *.jpeg *.tif *.tiff *.pdf"),
                ("Все файлы", "*.*"),
            ],
        )
        if paths:
            self.files = list(paths)
            self.file_list.delete(0, tk.END)
            for p in self.files:
                self.file_list.insert(tk.END, p)

    def process_files(self) -> None:
        self.items = []
        for path in self.files:
            text = extract_text_from_file(path)
            self.items.extend(parse_items(text))
        messagebox.showinfo("Готово", f"Обработано позиций: {len(self.items)}")

    def export(self) -> None:
        if not self.items:
            messagebox.showwarning("Нет данных", "Сначала обработайте файлы.")
            return
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")]
        )
        if filename:
            export_to_excel(self.items, filename)
            messagebox.showinfo("Сохранено", f"Файл сохранён: {filename}")


def run_app() -> None:
    root = tk.Tk()
    root.title("Распознавание документов")
    app = Application(master=root)
    app.mainloop()
