import tkinter
from pathlib import Path
from tkinter import filedialog
import openpyxl
from datetime import datetime

class Application(tkinter.Frame):
    def __init__(self, root=None):
        super().__init__(root, width=380, height=280,
                         borderwidth=1, relief='groove')
        self.root = root
        self.pack()
        self.pack_propagate(0)
        self.create_widgets()

    def create_widgets(self):
        # 閉じるボタン
        quit_btn = tkinter.Button(self)
        quit_btn['text'] = '閉じる'
        quit_btn['command'] = self.root.destroy
        quit_btn.pack(side='bottom')

        # テキストボックス (請求番号用)
        self.invoice_num_label = tkinter.Label(self, text="請求番号:")
        self.invoice_num_label.pack()
        self.invoice_num_box = tkinter.Entry(self)
        self.invoice_num_box['width'] = 20
        self.invoice_num_box.pack()

        # テキストボックス (請求先名用)
        self.client_name_label = tkinter.Label(self, text="請求先名:")
        self.client_name_label.pack()
        self.client_name_box = tkinter.Entry(self)
        self.client_name_box['width'] = 20
        self.client_name_box.pack()

        # 実行ボタン
        submit_btn = tkinter.Button(self)
        submit_btn['text'] = '請求書発行'
        submit_btn['command'] = self.save_data
        submit_btn.pack()

        # メッセージ出力
        self.message = tkinter.Message(self)
        self.message.pack()

    def save_data(self):
        file_name = '新規納品書兼請求書.xlsx'
        new_file_name = '編集済み納品書兼請求書.xlsx'
        wb = openpyxl.load_workbook(file_name)
        ws = wb.worksheets[0]
        
        # 日付を取得してフォーマット
        today = datetime.today()
        formatted_date = f"{today.year}年{today.month}月{today.day}日"
        
        ws['A1'].value = formatted_date
        ws['H2'].value = self.invoice_num_box.get()  # 請求番号
        ws['A8'].value = self.client_name_box.get() + "　様"  # 請求先名
        wb.save(new_file_name)
        self.message['text'] = '保存完了'

root = tkinter.Tk()
root.title('請求書発行アプリ')
root.geometry('400x300')
app = Application(root=root)
app.mainloop()
