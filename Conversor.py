import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import openpyxl
import re
import os

def txt_to_xlsx(txt_file, xlsx_file):
    wb = openpyxl.Workbook()
    ws = wb.active

    with open(txt_file, 'r', encoding='utf-8') as file:
        row_num = 1
        for line in file:
            cells = re.split('\s{2,}', line.strip())
            for col_num, cell in enumerate(cells, start=1):
                ws.cell(row=row_num, column=col_num, value=cell)
            row_num += 1

    wb.save(xlsx_file)
    messagebox.showinfo("Concluído", "O arquivo foi gerado com sucesso!")

def select_input_file():
    filename = filedialog.askopenfilename(
        title="Selecionar arquivo TXT", 
        filetypes=[("Arquivos de texto", "*.txt")]
    )
    if filename:
        txt_entry.delete(0, tk.END)
        txt_entry.insert(0, filename)


def select_output_file():
    default_filename = "output.xlsx"
    initial_directory = os.path.expanduser("~")
    filename = filedialog.asksaveasfilename(
        title="Salvar arquivo XLSX", 
        defaultextension=".xlsx", 
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=initial_directory,
        initialfile=default_filename
    )
    xlsx_entry.delete(0, tk.END)
    xlsx_entry.insert(0, filename)


def convert():
    txt_file = txt_entry.get()
    if not txt_file:
        messagebox.showerror("Erro", "Por favor, selecione um arquivo TXT.")
        return

    xlsx_file = xlsx_entry.get()
    if not xlsx_file:
        messagebox.showerror("Erro", "Por favor, selecione o local para salvar o arquivo XLSX.")
        return

    txt_to_xlsx(txt_file, xlsx_file)

# Configuração da janela principal
root = tk.Tk()
root.title("Conversor TXT para XLSX")

# Frame para os widgets
frame = tk.Frame(root)
frame.pack(padx=20, pady=20)

# Entrada para o nome do arquivo TXT
tk.Label(frame, text="Converta seu TXT em XLSX").grid(row=0, column=0, columnspan=2)
txt_entry = tk.Entry(frame, width=40)
txt_entry.grid(row=1, column=1, pady=20)

# Botão para selecionar arquivo TXT
tk.Button(frame, width=20, text="Selecionar seu TXT", command=select_input_file).grid(row=1, column=0)

# Entrada para o local de salvamento do arquivo XLSX
xlsx_entry = tk.Entry(frame, width=40)
xlsx_entry.grid(row=4, column=1, padx=10, pady=0)

# Botão para selecionar local de salvamento do arquivo XLSX
tk.Button(frame,  width=20, text="Local para salvar XLSX", command=select_output_file).grid(row=4, column=0)

# Botão de conversão
tk.Button(frame, text="Converter para XLSX", command=convert).grid(row=5, columnspan=2, pady=20)

root.mainloop()
