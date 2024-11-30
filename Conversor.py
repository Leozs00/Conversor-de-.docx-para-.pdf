import os
from docx2pdf import convert
import tkinter as tk
from tkinter import filedialog, messagebox

def select_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        folder_label.config(text=f"Pasta selecionada: {folder_path}")
        convert_files(folder_path)

def convert_files(folder_path):
    try:
        files = os.listdir(folder_path)
        docx_files = [f for f in files if f.endswith('.docx')]

        if not docx_files:
            messagebox.showinfo("Informação", "Não existem arquivos .docx para conversão na pasta selecionada.")
            return

        for file in docx_files:
            file_path = os.path.join(folder_path, file)
            convert(file_path)
            print(f'Conversão realizada com sucesso para o arquivo: {file}')

        messagebox.showinfo("Sucesso", "Todos os arquivos .docx foram convertidos para .pdf com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

# Interface gráfica
root = tk.Tk()
root.title("Conversor de DOCX para PDF")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=10, pady=10)

folder_label = tk.Label(frame, text="Nenhuma pasta selecionada")
folder_label.pack(pady=5)

select_button = tk.Button(frame, text="Selecionar Pasta", command=select_folder)
select_button.pack(pady=5)

root.mainloop()