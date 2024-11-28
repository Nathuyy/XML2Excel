import os
import tkinter as tkinter
from tkinter import filedialog, messagebox


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Converta XML para Excel")
        self.root.geometry("400x200")  
        
        self.title_label = tkinter.Label(root, text="XML para Excel", font=("Arial", 16))
        self.title_label.pack(pady=10)
        
        self.select_button = tkinter.Button(root, text="Selecionar Pasta", command=self.select_folder)
        self.select_button.pack(pady=10)
        
        self.convert_button = tkinter.Button(root, text="Converter para Excel", command=self.convert_to_excel)
        self.convert_button.pack(pady=10)
        
        self.folder_label = tkinter.Label(root, text="Nenhuma pasta selecionada", fg="gray", wraplength=400)
        self.folder_label.pack(pady=5)

    def select_folder(self):
        """Permite que o usuário selecione uma pasta."""
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.folder_path = folder_path
            self.folder_label.config(text=f"Pasta selecionada: {folder_path}")
        else:
            self.folder_label.config(text="Nenhuma pasta selecionada", fg="gray")

    def convert_to_excel(self):
        """Realiza a conversão dos arquivos XML para Excel."""
        if hasattr(self, 'folder_path'):  
            xml_files = [f for f in os.listdir(self.folder_path) if f.endswith('.xml')]
            
            if not xml_files:
                messagebox.showwarning("Aviso", "A pasta selecionada não contém arquivos XML.")
                return



if __name__ == "__main__":
    root = tkinter.Tk()
    app = MainWindow(root)
    root.mainloop()
