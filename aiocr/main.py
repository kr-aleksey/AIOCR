import logging
import threading
import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

from openpyxl.utils.exceptions import InvalidFileException

from aiocr.loggers import TextHandler
from ocr import OpenAIOCR
from services import Workbook


logger = logging.getLogger('main')

class App(tk.Frame):
    def __init__(self, master, workbook):
        super().__init__(master)
        self.pack(padx=10, pady=10)

        self.workbook = workbook

        self.info_lbl = tk.Label(self, text='XLSX файл с изображениями:')
        self.info_lbl.pack()

        self.source_lbl = tk.Label(self, text='файл не выбран')
        self.source_lbl.pack()

        self.select_source_btn = tk.Button(self,
                                           text='Выберите файл',
                                           command=self.select_source_file)
        self.select_source_btn.pack()

        self.ocr_btn = tk.Button(self, text='OCR', command=self.ocr_images)
        self.ocr_btn.pack(pady=20)

        self.save_btn = tk.Button(self, text='Сохранить как...', command=self.save_as)
        self.save_btn.pack(pady=20)

        self.log_widget = ScrolledText(self, state='disabled', height=17)
        self.log_widget.configure(font='TkFixedFont')
        self.log_widget.pack(pady=20)

        logger.addHandler(TextHandler(self.log_widget))

    def select_source_file(self):
        path = filedialog.askopenfilename(filetypes=[("XLSX files", "*.xlsx")])
        try:
            self.workbook.load(path)
        except InvalidFileException:
            logger.error('Файл поврежден или это не xlsx файл.')
        self.source_lbl['text'] = path

    def run_ocr(self):
        self.workbook.ocr(OpenAIOCR())

    def ocr_images(self):
        thread = threading.Thread(target=self.run_ocr)
        thread.start()

    def save_as(self):
        path = filedialog.asksaveasfilename(
            confirmoverwrite=True,
            filetypes=[("XLSX files", "*.xlsx")],
        )
        self.workbook.save_as(path)


def main():
    logging.basicConfig(level=logging.INFO)

    workbook = Workbook()

    window = tk.Tk()
    window.title('AI OCR')
    window.geometry('600x500')

    app = App(window, workbook)
    app.mainloop()


if __name__ == '__main__':
    main()
