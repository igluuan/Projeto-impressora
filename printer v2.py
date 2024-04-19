import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
from openpyxl import *
import os

# Set Styles
ctk.set_default_color_theme("dark-blue")
ctk.set_appearance_mode("dark")

# Create window
class Window(ctk.CTk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Set widht and height for window
        self.geometry("350x350")
        # Set grid collumns
        self.grid_columnconfigure(0, weight=1)
        # Set max width
        self.maxsize(350, 350)
        # Set title
        self.title("Gerenciador de impressoras")
        # Set font window
        self.font = ctk.CTkFont(family='Poppins', size=14, weight="bold")

       
        # Entry/Input - Model
        self.boxInputModel = ctk.CTkEntry(self, placeholder_text="Modelo:", font=self.font)
        self.boxInputModel.grid(padx=20, pady=20, sticky="ew")

        # Entry/Input - Patrimony
        self.boxInputPatrimony = ctk.CTkEntry(self, placeholder_text="Patrimônio:", font=self.font)
        self.boxInputPatrimony.grid(padx=20, pady=20, sticky="ew")

        # List of options
        self.label = ctk.CTkLabel(self, text='Selecione a marca:', font=self.font)
        self.label.grid()
        self.options_brands = ["SAMSUNG", "OKIDATA", "HP", "RICOH","CANON"]
        self.cmBox = ctk.CTkOptionMenu(self, values=self.options_brands, font=self.font)
        self.cmBox.grid(padx=20, pady=10, sticky="ew")

        # Add button
        self.btn = ctk.CTkButton(self, fg_color="gray", text="Adicionar a planilha", command=self.btn_callback, font=self.font)
        self.btn.grid(padx=20, pady=20, sticky="ew")

        #  Add file
        self.btnAddFile = ctk.CTkButton(self, fg_color="gray", text="Upload Planilha", font=self.font, command=self.uploadWorkbook)
        self.btnAddFile.grid(padx=20, sticky="ew")

        
    # Functions
    def uploadWorkbook(self):
        filename = filedialog.askopenfilename(filetypes=[("Planilhas Excel", "*.xlsx"), ("Todos os arquivos", "*.*")])
        if filename:
            try:
                self.workbook = load_workbook(filename)
                filename_only = os.path.basename(filename)
                messagebox.showinfo("Upload bem-sucedido", f"'{filename_only}' carregada com sucesso!")
                if hasattr(self, 'labelFile'):
                    self.labelFile.destroy()
                # Label for file
                self.labelFile = ctk.CTkLabel(self, text=f'Planilha: {filename_only}', font=self.font)
                self.labelFile.grid(padx=20, pady=7, sticky="ew")
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar a planilha: {str(e)}")


    def btn_callback(self):
        brands = self.cmBox.get()
        patrimony = self.boxInputPatrimony.get()
        print(brands, patrimony, "Adicionado com sucesso!")
        self.boxInputModel.delete(0, ctk.END)
        self.boxInputPatrimony.delete(0, ctk.END)
        self.boxInputModel.focus()

    def run(self):
        self.mainloop()

if __name__ == "__main__":
    manager = Window()
    manager.run()        