import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
from analisar_folha import analisar_folha
from extrair_tabela_pdfplumber import gerar_excel

class FolhaPontoApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Folha de Ponto - Processador")
        self.root.geometry("600x300")
        self.root.configure(bg="#f2f2f2")
        self.root.resizable(False, False)

        self.file_path_var = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value = "Aguardando arquivo...")
        self.cancel_requested = False # Flag para cancelar o processamento

        self.create_widgets()
    
    def create_widgets(self):
        # Título
        title_label = tk.Label(
            self.root,
            text = "Relatório: Folha de Ponto",
            font = ("Segoe UI", 14, "bold"),
            fg="#333"
        )
        title_label.pack(pady=(20, 10))

        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=6)

        select_frame = tk.Frame(self.root)
        select_frame.pack(pady=10)

        # Botão para selecionar PDF
        select_button = ttk.Button(
            select_frame,
            text = "Selecionar PDF",
            command=self.select_file,
            width=18
        )
        select_button.pack(side="left", padx=(0, 10))

        # Campo que mostra o caminho do arquivo selecionado
        file_label = tk.Label(
            select_frame,
            textvariable=self.file_path_var,
            font=("Segoe UI", 9),
            bg="white",
            fg="#888",
            anchor="w",
            justify="left",
            width=45,
            wraplength=300   # 🔹 quebra de linha automática se o caminho for longo
        )
        file_label.pack(pady=5)

        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=10)

        # Botão de processamento
        process_button = ttk.Button(
            button_frame,
            text="Processar",
            command=self.start_processing,
            width=20
        )
        process_button.pack(side="left", padx=(0, 10))

        # Botão de cancelar
        cancel_button = ttk.Button(
            button_frame,
            text="Cancelar",
            command=self.cancel_processing,
            width=20
        )
        cancel_button.pack(pady=5)

        # Barra de progresso
        ttk.Style().configure("Custom.Horizontal.TProgressbar", thickness=12)
        self.progress_bar = ttk.Progressbar(
            self.root,
            variable=self.progress_var,
            maximum=100,
            length=400,
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(pady=(10, 5))

        # Texto de status
        self.status_label = tk.Label(
            self.root,
            textvariable = self.status_var,
            font = ("Segoe UI", 10),
            fg="#444"
        )
        self.status_label.pack(pady=(5, 10))

        

    def select_file(self):
        filetypes = (("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")),
        filename = filedialog.askopenfilename(
            title = "Selecione o arquivo PDF",
            filetypes=filetypes
        )
        if filename:
            if not filename.lower().endswith(".pdf"):
                messagebox.showerror('Erro', 'Por favor, selecione um arquivo PDF válido.')
                return
            self.file_path_var.set(filename)
            self.status_var.set("Arquivo selecionado. Pronto para processar.")

    def start_processing(self):
        if not self.file_path_var.get():
            messagebox.showerror('Erro', 'Por favor, selecione um arquivo PDF antes de processar.')
            return
        
        # Reincia variáveis de controle
        self.cancel_requested = False
        self.progress_var.set(0)
        self.status_var.set("Iniciando processamento...")

        # Cria uma thread separada para o processamento
        thread = threading.Thread(target = self.run_processing_pipeline)
        thread.start()

    def run_processing_pipeline(self):
        """Executa o processo real do PDF"""
        pdf_path = self.file_path_var.get()
        output_path = os.path.splitext(pdf_path)[0] + "_processado.xlsx"

        def progress_update(percent, message):
            self.progress_var.set(percent)
            self.status_var.set(message)
            self.root.update_idletasks()
            
        def is_cancelled():
            return self.cancel_requested
        
        try:
            # etapa 1: Gerar Excel a partir do PDF
            gerar_excel(
                pdf_path, 
                output_path, 
                progress_callback=progress_update, 
                cancel_flag=is_cancelled
            )
            if self.cancel_requested:
                return
            
            # etapa 2: Analisar a folha gerada
            progress_update(100, "Analisando folha de ponto...")
            analisar_folha(output_path)
            
            self.status_var.set("Processamento concluído!")
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{output_path}")

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            self.status_var.set("Erro durante o processamento.")

        

    def cancel_processing(self):
        """Permite cancelar o processamento."""
        self.cancel_requested = True
        self.status_var.set("Cancelando...")


if __name__ == "__main__":
    root = tk.Tk()
    app = FolhaPontoApp(root)
    root.mainloop()