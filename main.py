import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os, sys
from naoconformidade import analisar_conformidade
from extrair_tabela import gerar_excel
from verificacao import analisar_verificacao

def resource_path(relative_path):
    try:
        # Quando rodado no executável
        base_path = sys._MEIPASS
    except Exception:
        # Quando rodado no Python normal
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# ======================================================
# CLASSE PRINCIPAL - Interface e lógica do aplicativo
# ======================================================
class FolhaPontoApp:
    def __init__(self, root):
        # --- Configuração da janela principal ---
        self.root = root
        self.root.title("Processador de Folha de Ponto")
        self.root.geometry("620x380")
        self.root.configure(bg="#eef1f6")
        self.root.resizable(False, False)
        self.root.iconbitmap(resource_path("icone.ico"))

        # --- Variáveis de controle e status ---
        self.file_path_var = tk.StringVar()   # Caminho do arquivo selecionado
        self.progress_var = tk.DoubleVar()    # Valor da barra de progresso
        self.status_var = tk.StringVar(value="Aguardando arquivo...")  # Mensagem de status
        self.cancel_requested = False         # Flag para cancelamento do processamento
        self.tipo_relatorio = tk.StringVar(value="Relatório de verificação") # Variável do tipo de relatório

        # --- Criação dos componentes da interface ---
        self.create_widgets()

    # ======================================================
    # BLOCO DE CRIAÇÃO DOS WIDGETS (componentes visuais)
    # ======================================================
    def create_widgets(self):
        # ======= Estilos gerais =======
        style = ttk.Style()
        style.theme_use("clam")

        # Estilo dos botões
        style.configure(
            "TButton",
            font=("Segoe UI", 10, "bold"),
            padding=8,
            background="#034794",
            foreground="white",
            borderwidth=0,
            focusthickness=3,
            focuscolor="none"
        )
        style.map(
            "TButton",
            background=[("active", "#034794"), ("disabled", "#A9A9A9")]
        )

        # Estilo da barra de progresso
        style.configure(
            "Custom.Horizontal.TProgressbar",
            thickness=14,
            troughcolor="#d9d9d9",
            background="#034794",
            bordercolor="#d9d9d9"
        )

        # ======= Cabeçalho com título =======
        header_frame = tk.Frame(self.root, bg="#034794", height=60)
        header_frame.pack(fill="x")

        title_label = tk.Label(
            header_frame,
            text="Processador de Folha de Ponto \n COMANDO",
            font=("Segoe UI", 14, "bold"),
            fg="white",
            bg="#034794"
        )
        title_label.pack(pady=15)

        # ======= Área principal =======
        content_frame = tk.Frame(self.root, bg="#eef1f6")
        content_frame.pack(fill="both", expand=True, pady=15)

        # ----- Bloco de seleção de arquivo -----
        select_frame = tk.Frame(content_frame, bg="#eef1f6")
        select_frame.pack(pady=5)

        # ======= Seleção do tipo de relatório =======
        tipo_frame = tk.Frame(select_frame, bg="#eef1f6")
        tipo_frame.pack(pady=(0, 10))

        ttk.Label(
            tipo_frame,
            text="Tipo de Relatório:",
            font=("Segoe UI", 10),
            background="#eef1f6"
        ).pack(side="left", padx=(0, 6))

        # Combobox para escolher o tipo de relatório
        self.combo_tipo = ttk.Combobox(
            tipo_frame,
            textvariable=self.tipo_relatorio,
            values=["Relatório de verificação","Relatório de não conformidade"],
            state="readonly",
            width=25
        )
        self.combo_tipo.current(0)
        self.combo_tipo.pack(side="left")

        # ======= Botão para selecionar o PDF =======
        select_button = ttk.Button(
            select_frame,
            text="Selecionar PDF",
            command=self.select_file,
            width=20
        )
        select_button.pack(pady=8)

        # ======= Exibição do caminho do arquivo selecionado =======
        file_label = tk.Label(
            select_frame,
            textvariable=self.file_path_var,
            font=("Segoe UI", 9),
            bg="white",
            fg="#555",
            anchor="w",
            justify="left",
            relief="solid",
            bd=1,
            padx=6,
            width=55,
            wraplength=350
        )
        file_label.pack(pady=(5, 15))

        # ======= Botões de controle (Processar e Cancelar) =======
        button_frame = tk.Frame(content_frame, bg="#eef1f6")
        button_frame.pack(pady=5)

        process_button = ttk.Button(
            button_frame,
            text="▶ Processar",
            command=self.start_processing,
            width=20
        )
        process_button.pack(side="left", padx=(0, 10))

        cancel_button = ttk.Button(
            button_frame,
            text="✖ Cancelar",
            command=self.cancel_processing,
            width=20
        )
        cancel_button.pack(side="left")

        # ======= Barra de progresso =======
        self.progress_bar = ttk.Progressbar(
            content_frame,
            variable=self.progress_var,
            maximum=100,
            length=420,
            style="Custom.Horizontal.TProgressbar"
        )
        self.progress_bar.pack(pady=(20, 10))

        # ======= Status (mensagem inferior) =======
        self.status_label = tk.Label(
            content_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10, "italic"),
            fg="#333",
            bg="#eef1f6"
        )
        self.status_label.pack()

    # ======================================================
    # FUNÇÃO: Seleção do arquivo PDF
    # ======================================================
    def select_file(self):
        filetypes = (("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*"))
        filename = filedialog.askopenfilename(
            title="Selecione o arquivo PDF", filetypes=filetypes
        )
        if filename:
            if not filename.lower().endswith(".pdf"):
                messagebox.showerror("Erro", "Por favor, selecione um arquivo PDF válido.")
                return
            # Atualiza o caminho e o status
            self.file_path_var.set(filename)
            self.status_var.set("Arquivo selecionado. Pronto para processar.")

    # ======================================================
    # FUNÇÃO: Inicia o processamento (thread separada)
    # ======================================================
    def start_processing(self):
        if not self.file_path_var.get():
            messagebox.showerror("Erro", "Por favor, selecione um arquivo PDF antes de processar.")
            return

        # Reinicia controles
        self.cancel_requested = False
        self.progress_var.set(0)
        self.status_var.set("Iniciando processamento...")

        # Cria thread para não travar a interface
        thread = threading.Thread(target=self.run_processing_pipeline)
        thread.start()

    # ======================================================
    # FUNÇÃO PRINCIPAL DE PROCESSAMENTO DO PDF
    # ======================================================
    def run_processing_pipeline(self):
        pdf_path = self.file_path_var.get()
        if self.tipo_relatorio.get() == "Relatório de verificação":
            output_path = os.path.splitext(pdf_path)[0] + "_verificacao.xlsx"
        elif self.tipo_relatorio.get() == "Relatório de não conformidade":
            output_path = os.path.splitext(pdf_path)[0] + "_naoconformidade.xlsx"

        # Callback interno para atualizar progresso
        def progress_update(percent, message):
            self.progress_var.set(percent)
            self.status_var.set(message)
            self.root.update_idletasks()

        # Função que verifica se o usuário cancelou
        def is_cancelled():
            return self.cancel_requested

        try:
            # --- Etapa 1: Gerar Excel a partir do PDF ---
            gerar_excel(
                pdf_path, 
                output_path, 
                progress_callback=progress_update, 
                cancel_flag=is_cancelled
            )
            if self.cancel_requested:
                return

            # --- Etapa 2: Analisar a folha gerada ---
            tipo = self.tipo_relatorio.get()
            progress_update(100, "Aguarde. Analisando folha de ponto...")

            if tipo == "Relatório de não conformidade":
                analisar_conformidade(output_path)
            elif tipo == "Relatório de verificação":
                analisar_verificacao(output_path)
            else:
                raise ValueError("Tipo de relatório desconhecido")

            # --- Conclusão ---
            self.status_var.set("Processamento concluído!")
            messagebox.showinfo("Sucesso", f"Arquivo salvo em:\n{output_path}")

        except Exception as e:
            # Exibe mensagem de erro em caso de falha
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")
            self.status_var.set("Erro durante o processamento.")

    # ======================================================
    # FUNÇÃO: Cancelamento do processamento
    # ======================================================
    def cancel_processing(self):
        self.cancel_requested = True
        self.status_var.set("Cancelando...")

# ======================================================
# EXECUÇÃO DO APLICATIVO
# ======================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = FolhaPontoApp(root)
    root.mainloop()
