import os
import sys
import json
import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import customtkinter as ctk
from tkcalendar import DateEntry
from datetime import datetime, timedelta
import time # Importar para usar time.sleep

# --- Funções Auxiliares ---
def resource_path(relative_path):
    """
    Obtém o caminho absoluto para um recurso, funciona tanto no desenvolvimento
    quanto com executáveis gerados pelo PyInstaller.
    """
    try:
        # PyInstaller cria uma pasta temporária onde os arquivos ficam após empacotamento
        base_path = sys._MEIPASS
    except Exception:
        # Rodando como script python
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Classe Principal da Aplicação ---
class UserManagerApp:
    def __init__(self, master):
        self.master = master
        master.title("Controle de Usuários")

        # Configurações iniciais do CustomTkinter para o modo escuro global
        ctk.set_appearance_mode("dark") # Definido para dark mode desde o início
        ctk.set_default_color_theme("blue")

        # Define o caminho base para pastas de dados
        # Garante que o diretório base seja sempre o do executável/script
        self.base_dir = self._get_base_directory()
        print(f"Base Directory: {self.base_dir}") # DEBUG: Imprime o diretório base
        self.pasta_equipes = os.path.join(self.base_dir, "equipes")
        self.pasta_usuarios = os.path.join(self.base_dir, "usuarios")
        self.pasta_desligados = os.path.join(self.base_dir, "desligados")
        # Define o caminho para o arquivo de configuração (para a URA)
        self.config_file = os.path.join(self.base_dir, "config.json")

        # Garante que as pastas existam
        os.makedirs(self.pasta_equipes, exist_ok=True)
        os.makedirs(self.pasta_usuarios, exist_ok=True)
        os.makedirs(self.pasta_desligados, exist_ok=True)

        # Remove o ícone da janela
        master.iconbitmap("")

        # Define tamanho máximo da tela
        largura_tela = master.winfo_screenwidth()
        altura_tela = master.winfo_screenheight() - 40
        master.geometry(f"{largura_tela}x{altura_tela}+0+0")
        master.resizable(False, False)
        master.configure(fg_color="white") # Cor de fundo da janela principal para branco

        self.caminho_imagem_intro = resource_path("imagem/Essencicred.jpg")

        # Variáveis para frames de controle
        self.frame_indice = None
        self.frame_divisor = None
        self.frame_conteudo = None
        self.frame_intro = None
        self.saving_frame = None # Novo frame para a tela de salvamento
        
        # Variáveis para sugestões de desligamento
        self.all_employee_names = [] # Armazena nomes e filenames de todos os funcionários ativos
        self.suggestion_frame = None
        self.current_employee_data = None # Para armazenar os dados do funcionário selecionado

        # Variável para a quantidade de URA e os labels que a exibem
        self.ura_quantity = self._load_ura_quantity() # Carrega a URA ao iniciar
        self.ura_label_value = None # Para armazenar a referência do label da URA
        self.argus_contracted_label = None # Para armazenar a referência do label de Argus Contratados

        # Configura o protocolo de fechar a janela
        master.protocol("WM_DELETE_WINDOW", self.on_closing)

        self.show_intro_screen() # Inicia com a tela de introdução

    def _get_base_directory(self):
        """Determina o diretório base para arquivos do programa (onde o executável/script está)."""
        if getattr(sys, 'frozen', False):
            # Se for um executável PyInstaller, o diretório base é onde o executável está
            return os.path.dirname(sys.executable)
        else:
            # Se for um script Python, o diretório base é onde o script está
            return os.path.dirname(os.path.abspath(__file__))

    def clear_content_frames(self):
        """Destrói os frames de conteúdo e introdução."""
        if self.frame_conteudo:
            self.frame_conteudo.destroy()
            self.frame_conteudo = None
        if self.frame_intro:
            self.frame_intro.destroy()
            self.frame_intro = None
        if self.saving_frame: # Destrói o frame de salvamento se estiver visível
            self.saving_frame.destroy()
            self.saving_frame = None

    def clear_sidebar_frames(self):
        """Destrói os frames do índice (menu lateral) e do divisor."""
        if self.frame_indice:
            self.frame_indice.destroy()
            self.frame_indice = None
        if self.frame_divisor:
            self.frame_divisor.destroy()
            self.frame_divisor = None

    def create_custom_title_bar(self, parent_toplevel, title_text):
        """Cria uma barra de título personalizada e arrastável para um CTkToplevel."""
        title_bar = ctk.CTkFrame(parent_toplevel, fg_color="black", height=30, corner_radius=0)
        title_bar.pack(fill="x", side="top")
        title_bar.grid_columnconfigure(0, weight=1) # Para o label do título
        title_bar.grid_columnconfigure(1, weight=0) # Para o botão de fechar

        title_label = ctk.CTkLabel(title_bar, text=title_text, font=ctk.CTkFont("Arial", 14, weight="bold") or ctk.CTkFont("Segoe UI", 14, weight="bold"), text_color="white")
        title_label.grid(row=0, column=0, sticky="w", padx=10)

        close_button = ctk.CTkButton(title_bar, text="X", command=parent_toplevel.destroy,
                                     width=30, height=30, corner_radius=5, # Alterado: corner_radius para 5
                                     fg_color="#CC3333", hover_color="#A02020", # Tons de vermelho menos vibrantes
                                     font=ctk.CTkFont("Arial", 14, weight="bold") or ctk.CTkFont("Segoe UI", 14, weight="bold"))
        close_button.grid(row=0, column=1, sticky="e")

        # Funcionalidade de arrastar
        x_start = 0
        y_start = 0

        def start_drag(event):
            nonlocal x_start, y_start
            x_start = event.x
            y_start = event.y

        def do_drag(event):
            x = parent_toplevel.winfo_x() + event.x - x_start
            y = parent_toplevel.winfo_y() + event.y - y_start
            parent_toplevel.geometry(f"+{x}+{y}")

        title_bar.bind("<Button-1>", start_drag)
        title_bar.bind("<B1-Motion>", do_drag)
        title_label.bind("<Button-1>", start_drag)
        title_label.bind("<B1-Motion>", do_drag)

        return title_bar

    def show_message_confirmation(self, message, on_confirm):
        """Exibe um popup de confirmação com opções Sim/Não."""
        popup = ctk.CTkToplevel(self.master)
        popup.geometry("400x180")
        popup.transient(self.master)
        popup.grab_set()
        popup.overrideredirect(True) # Remove a barra de título nativa

        self.create_custom_title_bar(popup, "Confirmação")

        # Alterado: Fundo do conteúdo para um preto mais escuro
        content_frame = ctk.CTkFrame(popup, fg_color="#1a1a1a")
        content_frame.pack(fill="both", expand=True)

        # Centraliza o popup
        popup.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() / 2) - (popup.winfo_width() / 2)
        y = self.master.winfo_y() + (self.master.winfo_height() / 2) - (popup.winfo_height() / 2)
        popup.geometry(f"+{int(x)}+{int(y)}")

        label = ctk.CTkLabel(content_frame, text=message, font=ctk.CTkFont("Arial", 16) or ctk.CTkFont("Segoe UI", 16), text_color="white")
        label.pack(pady=20)

        frame_buttons = ctk.CTkFrame(content_frame, fg_color="transparent")
        frame_buttons.pack(pady=10)

        def confirm():
            on_confirm()
            popup.grab_release()
            popup.destroy()

        def cancel():
            popup.grab_release()
            popup.destroy()

        ctk.CTkButton(frame_buttons, text="Sim", width=100, command=confirm, font=ctk.CTkFont("Arial", 14) or ctk.CTkFont("Segoe UI", 14)).grid(row=0, column=0, padx=10)
        ctk.CTkButton(frame_buttons, text="Não", width=100, fg_color="gray", command=cancel, font=ctk.CTkFont("Arial", 14) or ctk.CTkFont("Segoe UI", 14)).grid(row=0, column=1, padx=10)

    def show_info_message_auto_close(self, message, duration_ms=2000):
        """Exibe um popup informativo que fecha automaticamente."""
        popup = ctk.CTkToplevel(self.master)
        popup.geometry("350x150")
        popup.transient(self.master)
        popup.grab_set()
        popup.overrideredirect(True) # Remove a barra de título nativa

        self.create_custom_title_bar(popup, "Informação")

        # Alterado: Fundo do conteúdo para um preto mais escuro
        content_frame = ctk.CTkFrame(popup, fg_color="#1a1a1a")
        content_frame.pack(fill="both", expand=True)

        # Centraliza o popup
        popup.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() / 2) - (popup.winfo_width() / 2)
        y = self.master.winfo_y() + (self.master.winfo_height() / 2) - (popup.winfo_height() / 2)
        popup.geometry(f"+{int(x)}+{int(y)}")

        label = ctk.CTkLabel(content_frame, text=message, font=ctk.CTkFont("Arial", 16) or ctk.CTkFont("Segoe UI", 16), text_color="white")
        label.pack(pady=30)

        def close_popup():
            popup.grab_release()
            popup.destroy()

        popup.after(duration_ms, close_popup)

    # --- Funções para Salvar/Carregar Configurações (URA) ---
    def _save_ura_quantity(self):
        """Salva a quantidade de URA em um arquivo de configuração."""
        try:
            with open(self.config_file, 'w') as f:
                json.dump({"ura_quantity": self.ura_quantity}, f)
        except Exception as e:
            print(f"Erro ao salvar quantidade de URA: {e}")

    def _load_ura_quantity(self):
        """Carrega a quantidade de URA de um arquivo de configuração."""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    return config.get("ura_quantity", 0)
            return 0
        except Exception as e:
            print(f"Erro ao carregar quantidade de URA: {e}")
            return 0

    # --- Funções para o contador de URA ---
    def _increment_ura(self):
        """Aumenta a quantidade de URA em 1 e atualiza os labels."""
        self.ura_quantity += 1
        self.ura_label_value.configure(text=str(self.ura_quantity))
        self._update_argus_contracted_display()

    def _decrement_ura(self):
        """Diminui a quantidade de URA em 1, sem ir abaixo de zero, e atualiza os labels."""
        if self.ura_quantity > 0:
            self.ura_quantity -= 1
            self.ura_label_value.configure(text=str(self.ura_quantity))
            self._update_argus_contracted_display()

    def _update_argus_contracted_display(self):
        """Atualiza o label que mostra a soma de URA e usuários Argus."""
        # Recalcula os totais para pegar o valor atualizado de 'Argus Total'
        counts = self._analyze_home_screen_data()
        argus_total_users = counts.get("Argus Total", 0)
        total_contracted = self.ura_quantity + argus_total_users
        if self.argus_contracted_label:
            self.argus_contracted_label.configure(text=str(total_contracted))

    # --- Telas da Aplicação ---
    def show_intro_screen(self):
        """Exibe a tela de introdução com a imagem e o botão "Entrar"."""
        self.clear_content_frames()
        self.clear_sidebar_frames()

        self.master.iconbitmap("") # Garante que o ícone seja removido na tela de introdução

        # O frame de introdução agora será branco, conforme solicitado
        self.frame_intro = ctk.CTkFrame(self.master, fg_color="white")
        self.frame_intro.grid(row=0, column=0, sticky="nsew", padx=50, pady=50, columnspan=3)

        self.master.grid_columnconfigure(0, weight=1)
        self.master.grid_columnconfigure(1, weight=1)
        self.master.grid_columnconfigure(2, weight=1)
        self.master.grid_rowconfigure(0, weight=1)

        if os.path.exists(self.caminho_imagem_intro):
            try:
                imagem = Image.open(self.caminho_imagem_intro)
                largura_maxima = 400
                proporcao = largura_maxima / imagem.width
                nova_altura = int(imagem.height * proporcao)
                imagem = imagem.resize((largura_maxima, nova_altura), Image.LANCZOS)
                imagem_ctk = ctk.CTkImage(light_image=imagem, size=(largura_maxima, nova_altura))
                label_imagem = ctk.CTkLabel(self.frame_intro, image=imagem_ctk, text="")
                label_imagem.pack(pady=(0, 20)) # Ajuste o padding superior aqui
            except Exception as e:
                label_erro = ctk.CTkLabel(self.frame_intro, text=f"Erro ao carregar imagem: {e}", text_color="black") # Alterado para black
                label_erro.pack(pady=20)
        else:
            label_erro = ctk.CTkLabel(self.frame_intro, text="Imagem 'Essencicred.jpg' não encontrada.", text_color="black") # Alterado para black
            label_erro.pack(pady=20)

        # Alterado: Títulos com 'Georgia' e corpo com 'Segoe UI' ou 'Arial'
        ctk.CTkLabel(self.frame_intro, text="Bem-vindo!",
                     font=ctk.CTkFont("Georgia", 28, weight="bold"),
                     text_color="black").pack(pady=(10, 5)) # Alterado para black

        ctk.CTkLabel(self.frame_intro, text="Sistema de Controle de Usuários",
                     font=ctk.CTkFont("Georgia", 23),
                     text_color="black").pack(pady=(5, 10)) # Alterado para black

        ctk.CTkLabel(self.frame_intro, text="Gerencie seus colaboradores e equipes de forma eficiente.",
                     font=ctk.CTkFont("Segoe UI", 14) or ctk.CTkFont("Arial", 14),
                     text_color="gray").pack(pady=(0, 30))

        ctk.CTkButton(self.frame_intro, text="Entrar", command=self.start_main_system,
                      width=200, height=50, corner_radius=25,
                      fg_color="#046F91", font=ctk.CTkFont("Segoe UI", 20, weight="bold") or ctk.CTkFont("Arial", 20, weight="bold")).pack(pady=20)


    def start_main_system(self):
        """Inicia o sistema principal após a tela de introdução."""
        ctk.set_appearance_mode("dark")
        self.master.configure(fg_color="#222222")
        self.clear_content_frames()

        self.master.iconbitmap("") # Garante que o ícone seja removido ao iniciar o sistema principal

        self.frame_indice = ctk.CTkFrame(self.master, fg_color="#333333", width=220)
        self.frame_indice.grid(row=0, column=0, sticky="ns")
        self.frame_indice.grid_propagate(False)

        button_properties = {
            "width": 180,
            "height": 40,
            "corner_radius": 10,
            "fg_color": "#046F91",
            "font": ctk.CTkFont("Georgia", 16, weight="bold") or ctk.CTkFont("Arial", 16, weight="bold") # Alterado: Font para Georgia
        }

        self.home_button = ctk.CTkButton(self.frame_indice, text="INICIO", command=lambda: self.show_home_screen(force_reload=True), **button_properties)
        self.home_button.pack(pady=(30, 10))
        self.teams_button = ctk.CTkButton(self.frame_indice, text="EQUIPES", command=lambda: self.show_teams_screen(force_reload=True), **button_properties)
        self.teams_button.pack(pady=10)
        self.registration_button = ctk.CTkButton(self.frame_indice, text="CADASTRO", command=self.show_registration_screen, **button_properties)
        self.registration_button.pack(pady=10)
        self.termination_button = ctk.CTkButton(self.frame_indice, text="DESLIGAMENTO", command=self.show_termination_screen, **button_properties)
        self.termination_button.pack(pady=10)

        # Novo botão SAIR
        self.exit_button = ctk.CTkButton(self.frame_indice, text="SAIR", command=self.on_closing, **button_properties)
        self.exit_button.pack(pady=(50, 10)) # Adicione um padding maior para separar do resto

        self.frame_divisor = ctk.CTkFrame(self.master, fg_color="#555555", width=2)
        self.frame_divisor.grid(row=0, column=1, sticky="ns")

        self.master.grid_columnconfigure(0, weight=0)
        self.master.grid_columnconfigure(1, weight=0)
        self.master.grid_columnconfigure(2, weight=1)
        self.master.grid_rowconfigure(0, weight=1)

        self.show_home_screen(force_reload=True) # Exibe a tela inicial ao carregar o sistema

    def _analyze_home_screen_data(self):
        """
        Analisa os dados de todos os funcionários para gerar os totais para a tela inicial.
        Retorna um dicionário com as contagens otimizadas.
        """
        counts = {
            "Vendedor Interno EssenciCred": 0,
            "Formalizacao Interno EssenciCred": 0,
            "Total Interno EssenciCred": 0,
            "Vendedor Externo Parceiros": 0,
            "Formalizacao Externo Parceiros": 0,
            "Total Externo Parceiros": 0,
            "Corban Total": 0,
            "Argus Total": 0
        }

        if os.path.exists(self.pasta_usuarios):
            for user_file in os.listdir(self.pasta_usuarios):
                if user_file.endswith(".txt"):
                    user_data = self._read_employee_data(os.path.join(self.pasta_usuarios, user_file))
                    if user_data:
                        role = user_data.get('Função')
                        company = user_data.get('Empresa')
                        corban_access = user_data.get('Usuário do meu Corban')
                        argus_access = user_data.get('Usuário da Discadora Argus')

                        if company == "Interno EssenciCred":
                            if role == "Vendedor":
                                counts["Vendedor Interno EssenciCred"] += 1
                            elif role == "Formalização":
                                counts["Formalizacao Interno EssenciCred"] += 1
                            counts["Total Interno EssenciCred"] += 1
                        elif company == "Externo Parceiros":
                            if role == "Vendedor":
                                counts["Vendedor Externo Parceiros"] += 1
                            elif role == "Formalização":
                                counts["Formalizacao Externo Parceiros"] += 1
                            counts["Total Externo Parceiros"] += 1

                        if corban_access == "Sim":
                            counts["Corban Total"] += 1
                        if argus_access == "Sim":
                            counts["Argus Total"] += 1
        return counts


    def show_home_screen(self, force_reload=False):
        """Exibe a tela inicial 'Funcionários Vendedores' com os totais de relatórios e o contador de URA."""
        if self.frame_conteudo and not force_reload:
            # Se o frame de conteúdo já existe e não é para forçar recarregar, apenas retorna
            return

        self.clear_content_frames()

        self.frame_conteudo = ctk.CTkFrame(self.master, fg_color=self.master.cget("fg_color"))
        self.frame_conteudo.grid(row=0, column=2, sticky="nsew", padx=50, pady=50)
        self.frame_conteudo.grid_columnconfigure(0, weight=1)
        # Aumentamos o número de linhas configuradas para acomodar o novo bloco
        self.frame_conteudo.grid_rowconfigure((0,1,2,3,4,5), weight=1)

        # Última Atualização
        current_time = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        ctk.CTkLabel(
            self.frame_conteudo,
            text=f"Última Atualização: {current_time}",
            font=ctk.CTkFont("Arial", 12, weight="bold"), # Alterado para Arial
            text_color="white" # Alterado para branco
        ).grid(row=5, column=0, pady=(20, 0), sticky="s") # Usando grid para fixar na parte inferior e centralizar

        self.frame_conteudo.grid_rowconfigure(5, weight=0) # Garante que essa linha não se expanda
        
        ctk.CTkLabel(
            self.frame_conteudo,
            text="Visão Geral dos Funcionários",
            font=ctk.CTkFont("Georgia", 36, weight="bold"), # Título, mantém Georgia
            text_color="white",
            anchor="center"
        ).grid(row=0, column=0, pady=(0, 40), sticky="ew")

        counts = self._analyze_home_screen_data()

        def create_total_row(parent_frame, label_text, count_value, row_num, column_num=0, is_total=False):
            font_size = 16 # Reduzido de 18
            text_color = "white"
            font_weight = "normal"
            if is_total:
                font_size = 18 # Reduzido de 20
                font_weight = "bold"
                text_color = "#046F91" # Mantém a cor para destaque do total
            
            # Garante que a fonte seja Arial para todos os labels de dados
            ctk.CTkLabel(parent_frame, text=f"{label_text}:", font=ctk.CTkFont("Arial", font_size, weight=font_weight), text_color=text_color, anchor="w").grid(row=row_num, column=column_num, sticky="w", padx=20, pady=3) # Reduzido pady de 5 para 3
            ctk.CTkLabel(parent_frame, text=str(count_value), font=ctk.CTkFont("Arial", font_size, weight=font_weight), text_color=text_color if not is_total else "#046F91", anchor="e").grid(row=row_num, column=column_num + 1, sticky="e", padx=20, pady=3) # Reduzido pady de 5 para 3

        frame_company_role_totals = ctk.CTkFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10)
        # Reduzindo o pady para diminuir o espaço entre os blocos
        frame_company_role_totals.grid(row=1, column=0, pady=(10, 10), padx=10, sticky="nsew") 
        frame_company_role_totals.grid_columnconfigure(0, weight=1)
        frame_company_role_totals.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            frame_company_role_totals,
            text="Totais por Empresa e Função",
            font=ctk.CTkFont("Georgia", 20, weight="bold"), # Título, mantém Georgia
            text_color="white"
        ).grid(row=0, column=0, columnspan=2, pady=(15, 10))

        grid_row = 1

        # Labels de subtítulo com fonte Arial e cor branca
        ctk.CTkLabel(frame_company_role_totals, text="ESSENCICRED (INTERNO)", font=ctk.CTkFont("Arial", 16, weight="bold"), text_color="white", anchor="w").grid(row=grid_row, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 0)) # Reduzido font size de 18 para 16, pady de 10 para 8
        grid_row += 1
        create_total_row(frame_company_role_totals, "Vendedores", counts["Vendedor Interno EssenciCred"], grid_row)
        grid_row += 1
        create_total_row(frame_company_role_totals, "Formalização", counts["Formalizacao Interno EssenciCred"], grid_row)
        grid_row += 1
        create_total_row(frame_company_role_totals, "Total EssenciCred", counts["Total Interno EssenciCred"], grid_row, is_total=True)
        grid_row += 1

        ctk.CTkFrame(frame_company_role_totals, height=2, fg_color="#444444").grid(row=grid_row, column=0, columnspan=2, sticky="ew", pady=10, padx=10)
        grid_row += 1

        # Labels de subtítulo com fonte Arial e cor branca
        ctk.CTkLabel(frame_company_role_totals, text="PARCEIROS (EXTERNO)", font=ctk.CTkFont("Arial", 16, weight="bold"), text_color="white", anchor="w").grid(row=grid_row, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 0)) # Reduzido font size de 18 para 16, pady de 10 para 8
        grid_row += 1
        create_total_row(frame_company_role_totals, "Vendedores", counts["Vendedor Externo Parceiros"], grid_row)
        grid_row += 1
        create_total_row(frame_company_role_totals, "Formalização", counts["Formalizacao Externo Parceiros"], grid_row)
        grid_row += 1
        create_total_row(frame_company_role_totals, "Total Parceiros", counts["Total Externo Parceiros"], grid_row, is_total=True)
        grid_row += 1

        frame_access_totals = ctk.CTkFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10)
        # Reduzindo o pady para diminuir o espaço entre os blocos
        frame_access_totals.grid(row=2, column=0, pady=(10, 10), padx=10, sticky="nsew")
        frame_access_totals.grid_columnconfigure(0, weight=1)
        frame_access_totals.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            frame_access_totals,
            text="Total de Usuários com Acesso",
            font=ctk.CTkFont("Georgia", 20, weight="bold"), # Título, mantém Georgia
            text_color="white"
        ).grid(row=0, column=0, columnspan=2, pady=(15, 10))

        grid_row_access = 1
        create_total_row(frame_access_totals, "Meu Corban", counts["Corban Total"], grid_row_access)
        grid_row_access += 1
        create_total_row(frame_access_totals, "Discadora Argus", counts["Argus Total"], grid_row_access)
        grid_row_access += 1

        # --- Bloco de Quantidade de URA ---
        frame_ura_counter = ctk.CTkFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10)
        # Reduzindo o pady para diminuir o espaço entre os blocos
        frame_ura_counter.grid(row=3, column=0, pady=(10, 10), padx=10, sticky="nsew") # Posicionado na linha 3
        frame_ura_counter.grid_columnconfigure((0, 1, 2), weight=1)

        ctk.CTkLabel(
            frame_ura_counter,
            text="Quantidade de URA",
            font=ctk.CTkFont("Georgia", 20, weight="bold"), # Título, mantém Georgia
            text_color="white"
        ).grid(row=0, column=0, columnspan=3, pady=(15, 10))

        ctk.CTkButton(
            frame_ura_counter,
            text="▼",
            font=ctk.CTkFont("Arial", 18, weight="bold"), # Reduzido de 20 para 18
            width=50,
            height=40,
            fg_color="#046F91",
            command=self._decrement_ura
        ).grid(row=1, column=0, padx=(20, 5), pady=10, sticky="e")

        self.ura_label_value = ctk.CTkLabel(
            frame_ura_counter,
            text=str(self.ura_quantity),
            font=ctk.CTkFont("Arial", 28, weight="bold"), # Reduzido de 30 para 28
            text_color="white" # Alterado para branco
        )
        self.ura_label_value.grid(row=1, column=1, padx=5, pady=10)

        ctk.CTkButton(
            frame_ura_counter,
            text="▲",
            font=ctk.CTkFont("Arial", 18, weight="bold"), # Reduzido de 20 para 18
            width=50,
            height=40,
            fg_color="#046F91",
            command=self._increment_ura
        ).grid(row=1, column=2, padx=(5, 20), pady=10, sticky="w")
        # --- Fim do Bloco de Quantidade de URA ---

        # --- Novo Bloco: Quantidade de Usuários Contratados na Argus ---
        frame_argus_contracted = ctk.CTkFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10)
        # Posicionado na linha 4, logo abaixo do bloco de URA
        # Reduzindo o pady para diminuir o espaço entre os blocos
        frame_argus_contracted.grid(row=4, column=0, pady=(10, 10), padx=10, sticky="nsew")
        frame_argus_contracted.grid_columnconfigure((0, 1), weight=1) # Duas colunas para label e valor

        ctk.CTkLabel(
            frame_argus_contracted,
            text="Quantidade de Usuários Contratados na Argus:",
            font=ctk.CTkFont("Georgia", 20, weight="bold"), # Título, mantém Georgia
            text_color="white",
            anchor="w" # Alinha o texto à esquerda
        ).grid(row=0, column=0, padx=20, pady=(15, 10), sticky="ew")

        # Label de valor com fonte Arial e cor branca
        argus_total_users = counts.get("Argus Total", 0)
        total_contracted_argus = self.ura_quantity + argus_total_users

        self.argus_contracted_label = ctk.CTkLabel(
            frame_argus_contracted,
            text=str(total_contracted_argus),
            font=ctk.CTkFont("Arial", 28, weight="bold"), # Reduzido de 30 para 28
            text_color="white", # Alterado para branco
            anchor="e" # Alinha o número à direita
        )
        self.argus_contracted_label.grid(row=0, column=1, padx=20, pady=(15, 10), sticky="ew")
        # --- Fim do Novo Bloco ---

        # Chama a função para garantir que o display seja atualizado corretamente na inicialização
        self._update_argus_contracted_display()


    def _analyze_team_data(self):
        """
        Analisa os dados dos funcionários para contar por equipe e por tipo de acesso.
        Retorna dois dicionários:
        - team_member_counts: {'Nome da Equipe': total_membros}
        - team_access_counts: {'Nome da Equipe': {'Corban': 0, 'Argus': 0}}
        """
        team_member_counts = {}
        team_access_counts = {}

        if os.path.exists(self.pasta_equipes):
            for team_file in os.listdir(self.pasta_equipes):
                if team_file.endswith(".txt"):
                    team_name_raw = os.path.splitext(team_file)[0]
                    team_name_display = team_name_raw.replace("_", " ")
                    team_member_counts[team_name_display] = 0
                    team_access_counts[team_name_display] = {'Corban': 0, 'Argus': 0}

        if os.path.exists(self.pasta_usuarios):
            for user_file in os.listdir(self.pasta_usuarios):
                if user_file.endswith(".txt"):
                    user_data = self._read_employee_data(os.path.join(self.pasta_usuarios, user_file))
                    if user_data:
                        team_name = user_data.get('Equipe', '').strip() # Garante que o nome da equipe seja limpo
                        corban_access = user_data.get('Usuário do meu Corban')
                        argus_access = user_data.get('Usuário da Discadora Argus')

                        if team_name and team_name in team_member_counts: # Verifica se a equipe existe
                            team_member_counts[team_name] += 1
                            if corban_access == "Sim":
                                team_access_counts[team_name]['Corban'] += 1
                            if argus_access == "Sim":
                                team_access_counts[team_name]['Argus'] += 1
        return team_member_counts, team_access_counts

    def show_teams_screen(self, force_reload=False):
        """
        Exibe a tela 'Gerenciar Equipes', agora com informações detalhadas por equipe.
        """
        if self.frame_conteudo and not force_reload:
            return

        self.clear_content_frames()

        self.frame_conteudo = ctk.CTkFrame(self.master, fg_color=self.master.cget("fg_color"))
        self.frame_conteudo.grid(row=0, column=2, sticky="nsew", padx=50, pady=50)
        self.frame_conteudo.grid_columnconfigure(0, weight=1)
        # Removido o grid_rowconfigure(4, weight=1) pois a lista de funcionários agora será um pop-up
        self.frame_conteudo.grid_rowconfigure(5, weight=0) # Garante que o botão de relatório não se expanda

        ctk.CTkLabel(
            self.frame_conteudo,
            text="Gerenciar Equipes e Relatórios",
            font=ctk.CTkFont("Georgia", 30), # Alterado
            text_color="white"
        ).grid(row=0, column=0, pady=(0, 20), sticky="ew")

        frame_action_buttons = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        frame_action_buttons.grid(row=1, column=0, pady=10)
        frame_action_buttons.grid_columnconfigure((0, 1), weight=1)

        ctk.CTkButton(
            frame_action_buttons,
            text="Criar Equipe",
            width=150,
            height=40,
            fg_color="#046F91",
            font=ctk.CTkFont("Arial", 16), # Alterado para Arial
            command=lambda: self.create_team_popup() # Chamada direta para o popup
        ).grid(row=0, column=0, padx=(0, 10), sticky="ew")

        ctk.CTkButton(
            frame_action_buttons,
            text="Excluir Equipe",
            width=150,
            height=40,
            fg_color="#990000",
            hover_color="#770000",
            font=ctk.CTkFont("Arial", 16), # Alterado para Arial
            command=self.open_delete_team_selection_window
        ).grid(row=0, column=1, sticky="ew")

        frame_total_users = ctk.CTkFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10)
        frame_total_users.grid(row=2, column=0, pady=(30, 15), padx=10, sticky="nsew")
        frame_total_users.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            frame_total_users,
            text="Total de Usuários por Equipe",
            font=ctk.CTkFont("Georgia", 20, weight="bold"), # Alterado
            text_color="white"
        ).pack(pady=(15, 10))

        team_member_counts, team_access_counts = self._analyze_team_data()

        if not team_member_counts or all(count == 0 for count in team_member_counts.values()):
            ctk.CTkLabel(frame_total_users, text="Nenhuma equipe cadastrada ou sem usuários ativos.", text_color="gray", font=ctk.CTkFont("Arial", 14)).pack(pady=10) # Alterado para Arial
        else:
            header_frame_total = ctk.CTkFrame(frame_total_users, fg_color="transparent")
            header_frame_total.pack(fill="x", padx=10, pady=(5, 0))
            ctk.CTkLabel(header_frame_total, text="Equipe", font=ctk.CTkFont("Arial", 14, weight="bold"), text_color="white", width=200, anchor="w").pack(side="left", padx=5) # Alterado para Arial
            ctk.CTkLabel(header_frame_total, text="Total de Usuários", font=ctk.CTkFont("Arial", 14, weight="bold"), text_color="white", width=150, anchor="w").pack(side="left", padx=5) # Alterado para Arial

            scrollable_frame_total = ctk.CTkScrollableFrame(frame_total_users, fg_color="transparent", height=150)
            scrollable_frame_total.pack(fill="both", expand=True, padx=10, pady=(0, 10))

            for team, count in team_member_counts.items():
                if count > 0: # Apenas mostra equipes com usuários
                    row_frame = ctk.CTkFrame(scrollable_frame_total, fg_color="#3e3e3e", corner_radius=5)
                    row_frame.pack(fill="x", pady=2, padx=2)
                    team_label = ctk.CTkLabel(row_frame, text=team, text_color="white", font=ctk.CTkFont("Arial", 12), width=200, anchor="w", cursor="hand2") # Alterado para Arial
                    team_label.pack(side="left", padx=5)
                    # O clique agora chama o novo método para o pop-up
                    team_label.bind("<Button-1>", lambda e, team_name=team: self.show_employees_popup(team_name))
                    ctk.CTkLabel(row_frame, text=str(count), text_color="white", font=ctk.CTkFont("Arial", 12), width=150, anchor="w").pack(side="left", padx=5) # Alterado para Arial

        frame_access_users = ctk.CTkFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10)
        frame_access_users.grid(row=3, column=0, pady=(15, 10), padx=10, sticky="nsew")
        frame_access_users.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            frame_access_users,
            text="Usuários por Tipo de Acesso (Corban / Argus)",
            font=ctk.CTkFont("Georgia", 20, weight="bold"), # Alterado
            text_color="white"
        ).pack(pady=(15, 10))

        if not team_access_counts or all(sum(ad.values()) == 0 for ad in team_access_counts.values()):
            ctk.CTkLabel(frame_access_users, text="Nenhuma equipe cadastrada ou dados de usuários.", text_color="gray", font=ctk.CTkFont("Arial", 14)).pack(pady=10) # Alterado para Arial
        else:
            header_frame_access = ctk.CTkFrame(frame_access_users, fg_color="transparent")
            header_frame_access.pack(fill="x", padx=10, pady=(5, 0))
            ctk.CTkLabel(header_frame_access, text="Equipe", font=ctk.CTkFont("Arial", 14, weight="bold"), text_color="white", width=200, anchor="w").pack(side="left", padx=5) # Alterado para Arial
            ctk.CTkLabel(header_frame_access, text="Meu Corban", font=ctk.CTkFont("Arial", 14, weight="bold"), text_color="white", width=100, anchor="w").pack(side="left", padx=5) # Alterado para Arial
            ctk.CTkLabel(header_frame_access, text="Discadora Argus", font=ctk.CTkFont("Arial", 14, weight="bold"), text_color="white", width=150, anchor="w").pack(side="left", padx=5) # Alterado para Arial

            scrollable_frame_access = ctk.CTkScrollableFrame(frame_access_users, fg_color="transparent", height=150)
            scrollable_frame_access.pack(fill="both", expand=True, padx=10, pady=(0, 10))

            for team, access_data in team_access_counts.items():
                # Apenas mostra equipes que têm usuários com Corban ou Argus
                if access_data['Corban'] > 0 or access_data['Argus'] > 0:
                    row_frame = ctk.CTkFrame(scrollable_frame_access, fg_color="#3e3e3e", corner_radius=5)
                    row_frame.pack(fill="x", pady=2, padx=2)
                    ctk.CTkLabel(row_frame, text=team, text_color="white", font=ctk.CTkFont("Arial", 12), width=200, anchor="w").pack(side="left", padx=5) # Alterado para Arial
                    ctk.CTkLabel(row_frame, text=str(access_data['Corban']), text_color="white", font=ctk.CTkFont("Arial", 12), width=100, anchor="w").pack(side="left", padx=5) # Alterado para Arial
                    ctk.CTkLabel(row_frame, text=str(access_data['Argus']), text_color="white", font=ctk.CTkFont("Arial", 12), width=150, anchor="w").pack(side="left", padx=5) # Alterado para Arial


        # Botão para gerar relatório de usuários (CSV)
        ctk.CTkButton(
            self.frame_conteudo,
            text="Gerar Relatório de Usuários (CSV)",
            width=250,
            height=40,
            fg_color="#046F91",
            font=ctk.CTkFont("Arial", 16, weight="bold"), # Alterado para Arial
            command=self.generate_user_report_csv
        ).grid(row=4, column=0, pady=(20, 0), sticky="s") # Posicionado na parte inferior da tela de equipes


    def show_employees_popup(self, team_name):
        """
        Exibe um pop-up com a lista de funcionários que pertencem à equipe selecionada,
        incluindo informações sobre acesso Argus e Meu Corban.
        """
        print(f"DEBUG: Attempting to show employees for team: '{team_name}' in popup.")
        employees_popup = ctk.CTkToplevel(self.master)
        employees_popup.geometry("600x500") # Aumentado a largura para acomodar as novas colunas
        employees_popup.transient(self.master)
        employees_popup.grab_set()
        employees_popup.overrideredirect(True) # Remove a barra de título nativa

        self.create_custom_title_bar(employees_popup, f"Funcionários da Equipe: {team_name}")

        # Alterado: Fundo do conteúdo para um preto mais escuro
        content_frame = ctk.CTkFrame(employees_popup, fg_color="#1a1a1a")
        content_frame.pack(fill="both", expand=True)

        # Centraliza o popup
        employees_popup.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() / 2) - (employees_popup.winfo_width() / 2)
        y = self.master.winfo_y() + (self.master.winfo_height() / 2) - (employees_popup.winfo_height() / 2)
        employees_popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(
            content_frame,
            text=f"Lista de Funcionários da Equipe: {team_name}",
            font=ctk.CTkFont("Arial", 18, weight="bold"),
            text_color="white"
        ).pack(pady=(15, 10))

        scrollable_employees_frame = ctk.CTkScrollableFrame(content_frame, fg_color="transparent")
        scrollable_employees_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Configura as colunas para o grid dentro do scrollable_employees_frame
        scrollable_employees_frame.grid_columnconfigure(0, weight=3) # Nome do funcionário
        scrollable_employees_frame.grid_columnconfigure(1, weight=1) # Meu Corban
        scrollable_employees_frame.grid_columnconfigure(2, weight=1) # Discadora Argus

        # Cabeçalhos da tabela no pop-up - Alterado para branco e fonte Arial
        ctk.CTkLabel(scrollable_employees_frame, text="Nome do Funcionário", font=ctk.CTkFont("Arial", 14, weight="bold"), text_color="white", anchor="w").grid(row=0, column=0, sticky="ew", padx=5, pady=2)
        ctk.CTkLabel(scrollable_employees_frame, text="Meu Corban", font=ctk.CTkFont("Arial", 14, weight="bold"), text_color="white", anchor="center").grid(row=0, column=1, sticky="ew", padx=5, pady=2)
        ctk.CTkLabel(scrollable_employees_frame, text="Discadora Argus", font=ctk.CTkFont("Arial", 14, weight="bold"), text_color="white", anchor="center").grid(row=0, column=2, sticky="ew", padx=5, pady=2)

        employees_data_in_selected_team = []
        print(f"DEBUG: Checking directory: {self.pasta_usuarios}")
        
        if os.path.exists(self.pasta_usuarios):
            for user_file in os.listdir(self.pasta_usuarios):
                if user_file.endswith(".txt"):
                    file_path = os.path.join(self.pasta_usuarios, user_file)
                    user_data = self._read_employee_data(file_path)
                    
                    employee_team = user_data.get('Equipe', '').strip()
                    
                    if employee_team == team_name.strip():
                        employees_data_in_selected_team.append({
                            'Nome Completo': user_data.get('Nome Completo', ''),
                            'Meu Corban': user_data.get('Usuário do meu Corban', 'Não'),
                            'Discadora Argus': user_data.get('Usuário da Discadora Argus', 'Não')
                        })
        
        # Ordena os funcionários pelo nome
        employees_data_in_selected_team.sort(key=lambda x: x['Nome Completo'])

        print(f"DEBUG: Found employees data for popup: {employees_data_in_selected_team}")
        
        if employees_data_in_selected_team:
            for i, employee in enumerate(employees_data_in_selected_team):
                row_num = i + 1 # Começa na linha 1 após os cabeçalhos
                ctk.CTkLabel(scrollable_employees_frame, text=employee['Nome Completo'], text_color="white", font=ctk.CTkFont("Arial", 12), anchor="w").grid(row=row_num, column=0, sticky="ew", padx=5, pady=1)
                ctk.CTkLabel(scrollable_employees_frame, text=employee['Meu Corban'], text_color="white", font=ctk.CTkFont("Arial", 12), anchor="center").grid(row=row_num, column=1, sticky="ew", padx=5, pady=1)
                ctk.CTkLabel(scrollable_employees_frame, text=employee['Discadora Argus'], text_color="white", font=ctk.CTkFont("Arial", 12), anchor="center").grid(row=row_num, column=2, sticky="ew", padx=5, pady=1)
        else:
            # Mensagem de nenhum funcionário - Alterado para branco e fonte Arial
            ctk.CTkLabel(scrollable_employees_frame, text="Nenhum funcionário encontrado nesta equipe.", text_color="white", font=ctk.CTkFont("Arial", 14)).grid(row=1, column=0, columnspan=3, pady=10)

        ctk.CTkButton(
            content_frame, # Botão "Fechar" agora no content_frame
            text="Fechar",
            command=lambda: employees_popup.destroy(),
            width=100,
            height=30,
            fg_color="#046F91",
            font=ctk.CTkFont("Arial", 14, weight="bold")
        ).pack(pady=10)


    def generate_user_report_csv(self):
        """Gera um arquivo CSV com todos os dados dos usuários ativos."""
        report_data = []
        headers = [
            "Nome", "CPF", "Empresa", "Função", "Equipe", "Data de Início",
            "Data de Término", "Usuário do meu Corban", "Usuário da Discadora Argus",
            "Obs.", "Email", "Telefone", "Endereço"
        ]
        report_data.append(headers)

        if os.path.exists(self.pasta_usuarios):
            for user_file in os.listdir(self.pasta_usuarios):
                if user_file.endswith(".txt"):
                    user_data = self._read_employee_data(os.path.join(self.pasta_usuarios, user_file))
                    if user_data:
                        row = [
                            user_data.get('Nome Completo', ''), # Alterado para 'Nome Completo'
                            user_data.get('CPF', ''),
                            user_data.get('Empresa', ''),
                            user_data.get('Função', ''),
                            user_data.get('Equipe', ''),
                            user_data.get('Data de Início', ''),
                            user_data.get('Data de Término', ''),
                            user_data.get('Usuário do meu Corban', ''),
                            user_data.get('Usuário da Discadora Argus', ''),
                            user_data.get('Obs.', ''),
                            user_data.get('Email', ''),
                            user_data.get('Telefone', ''),
                            user_data.get('Endereço', '')
                        ]
                        report_data.append(row)
        
        if not report_data or len(report_data) == 1: # Apenas cabeçalho
            self.show_info_message_auto_close("Nenhum dado para o relatório.", duration_ms=2000) # Mensagem concisa
            return

        # Caminho onde o arquivo será salvo
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        report_filename = f"relatorio_usuarios_{timestamp}.csv"
        report_path = os.path.join(self.base_dir, report_filename)

        try:
            import csv # Importar csv apenas quando necessário
            with open(report_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerows(report_data)
            self.show_info_message_auto_close(f"Relatório salvo com sucesso!", duration_ms=2000) # Mensagem concisa
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório CSV: {e}")

    def create_team_popup(self):
        """Exibe um popup para criar uma nova equipe."""
        popup = ctk.CTkToplevel(self.master)
        popup.geometry("400x200")
        popup.transient(self.master)
        popup.grab_set()
        popup.overrideredirect(True) # Remove a barra de título nativa

        self.create_custom_title_bar(popup, "Criar Nova Equipe")

        # Alterado: Fundo do conteúdo para um preto mais escuro
        content_frame = ctk.CTkFrame(popup, fg_color="#1a1a1a")
        content_frame.pack(fill="both", expand=True)

        # Centraliza o popup
        popup.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() / 2) - (popup.winfo_width() / 2)
        y = self.master.winfo_y() + (self.master.winfo_height() / 2) - (popup.winfo_height() / 2)
        popup.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(content_frame, text="Nome da Nova Equipe:", font=ctk.CTkFont("Arial", 16), text_color="white").pack(pady=10) # Alterado para Arial
        entry_team_name = ctk.CTkEntry(content_frame, width=250, font=ctk.CTkFont("Arial", 14)) # Alterado para Arial
        entry_team_name.pack(pady=5)

        def save_new_team():
            team_name = entry_team_name.get().strip()
            if not team_name:
                self.show_info_message_auto_close("Nome da equipe vazio.", duration_ms=2000) # Mensagem concisa
                return

            # Substitui espaços por underscores para o nome do arquivo, e remove caracteres especiais
            formatted_team_name = "".join(c for c in team_name if c.isalnum() or c.isspace()).replace(" ", "_")
            team_file_path = os.path.join(self.pasta_equipes, f"{formatted_team_name}.txt")

            if os.path.exists(team_file_path):
                self.show_info_message_auto_close("Equipe já existe.", duration_ms=2000) # Mensagem concisa
            else:
                try:
                    with open(team_file_path, 'w') as f:
                        f.write(f"Equipe: {team_name}\n")
                        f.write(f"Data de Criação: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                    self.show_info_message_auto_close(f"Equipe criada com sucesso!", duration_ms=2000) # Mensagem concisa
                    popup.grab_release()
                    popup.destroy()
                    self.show_teams_screen(force_reload=True) # Recarrega a tela de equipes
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao criar equipe: {e}")

        ctk.CTkButton(content_frame, text="Salvar Equipe", command=save_new_team, font=ctk.CTkFont("Arial", 14)).pack(pady=10) # Alterado para Arial

    def open_delete_team_selection_window(self):
        """Abre uma nova janela para selecionar e excluir equipes."""
        teams = [os.path.splitext(f)[0].replace("_", " ") for f in os.listdir(self.pasta_equipes) if f.endswith(".txt")]
        if not teams:
            self.show_info_message_auto_close("Nenhuma equipe para excluir.", duration_ms=2000) # Mensagem concisa
            return

        selection_window = ctk.CTkToplevel(self.master)
        selection_window.geometry("400x300")
        selection_window.transient(self.master)
        selection_window.grab_set()
        selection_window.overrideredirect(True) # Remove a barra de título nativa

        self.create_custom_title_bar(selection_window, "Excluir Equipe")

        # Alterado: Fundo do conteúdo para um preto mais escuro
        content_frame = ctk.CTkFrame(selection_window, fg_color="#1a1a1a")
        content_frame.pack(fill="both", expand=True)

        # Centraliza o popup
        selection_window.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() / 2) - (selection_window.winfo_width() / 2)
        y = self.master.winfo_y() + (self.master.winfo_height() / 2) - (selection_window.winfo_height() / 2)
        selection_window.geometry(f"+{int(x)}+{int(y)}")

        ctk.CTkLabel(content_frame, text="Selecione a equipe para excluir:", font=ctk.CTkFont("Arial", 16), text_color="white").pack(pady=10) # Alterado para Arial

        # Usar CTkComboBox para a seleção
        selected_team_var = ctk.StringVar(content_frame)
        selected_team_var.set(teams[0]) # Define o primeiro item como padrão
        team_dropdown = ctk.CTkComboBox(content_frame, values=teams, variable=selected_team_var,
                                        font=ctk.CTkFont("Arial", 14), width=300) # Alterado para Arial
        team_dropdown.pack(pady=10)

        def confirm_delete():
            team_to_delete_display = selected_team_var.get()
            team_to_delete_filename = team_to_delete_display.replace(" ", "_") + ".txt"
            team_file_path = os.path.join(self.pasta_equipes, team_to_delete_filename)

            # Verificar se existem funcionários associados a esta equipe
            employees_in_team = []
            if os.path.exists(self.pasta_usuarios):
                for user_file in os.listdir(self.pasta_usuarios):
                    if user_file.endswith(".txt"):
                        user_data = self._read_employee_data(os.path.join(self.pasta_usuarios, user_file))
                        if user_data and user_data.get('Equipe', '').strip() == team_to_delete_display.strip():
                            employees_in_team.append(user_data.get('Nome', ''))

            if employees_in_team:
                message = f"Não é possível excluir a equipe '{team_to_delete_display}'.\n" \
                          f"Funcionários associados: {', '.join(employees_in_team)}." # Mensagem concisa
                self.show_info_message_auto_close(message, duration_ms=5000)
            else:
                self.show_message_confirmation(f"Excluir '{team_to_delete_display}'?", # Mensagem concisa
                                                lambda: delete_team_action(team_file_path, selection_window))

        def delete_team_action(path, window_to_close):
            try:
                if os.path.exists(path):
                    os.remove(path)
                    self.show_info_message_auto_close(f"Equipe excluída com sucesso!", duration_ms=2000) # Mensagem concisa
                    window_to_close.grab_release()
                    window_to_close.destroy()
                    self.show_teams_screen(force_reload=True) # Recarrega a tela de equipes
                else:
                    self.show_info_message_auto_close("Erro: Arquivo não encontrado.", duration_ms=2000) # Mensagem concisa
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao excluir equipe: {e}")

        ctk.CTkButton(content_frame, text="Excluir", command=confirm_delete,
                      fg_color="#990000", hover_color="#770000", font=ctk.CTkFont("Arial", 14)).pack(pady=20) # Alterado para Arial
        
    # --- Telas de Cadastro ---
    def show_registration_screen(self):
        """Exibe a tela de cadastro de novos funcionários."""
        self.clear_content_frames()
        self.frame_conteudo = ctk.CTkFrame(self.master, fg_color=self.master.cget("fg_color"))
        self.frame_conteudo.grid(row=0, column=2, sticky="nsew", padx=50, pady=50)
        self.frame_conteudo.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.frame_conteudo,
            text="Cadastro de Novo Funcionário",
            font=ctk.CTkFont("Georgia", 30), # Alterado
            text_color="white"
        ).pack(pady=(0, 20))

        # Frame para os campos do formulário
        form_frame = ctk.CTkScrollableFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10, height=500)
        form_frame.pack(fill="both", expand=True, padx=10, pady=10)
        form_frame.grid_columnconfigure(0, weight=1)
        form_frame.grid_columnconfigure(1, weight=2)

        self.entry_fields = {} # Dicionário para armazenar as referências dos campos de entrada

        # Campos restaurados para a versão anterior
        labels_and_types = {
            "Nome Completo": "entry",
            "Empresa": ["Interno EssenciCred", "Externo Parceiros"],
            "Função": ["Vendedor", "Formalização", "Outros"],
            "Equipe": "dropdown", # Este será populado dinamicamente
            "Data de Início": "calendar",
            "Usuário do meu Corban": ["Sim", "Não"],
            "Usuário da Discadora Argus": ["Sim", "Não"],
        }

        row_num = 0
        for label_text, field_type in labels_and_types.items():
            ctk.CTkLabel(form_frame, text=f"{label_text}:", font=ctk.CTkFont("Arial", 14), text_color="white", anchor="w").grid(row=row_num, column=0, sticky="w", padx=10, pady=5) # Alterado para Arial

            if isinstance(field_type, list): # Dropdown com opções fixas
                combobox_var = ctk.StringVar(form_frame)
                combobox = ctk.CTkComboBox(form_frame, values=field_type, variable=combobox_var,
                                          font=ctk.CTkFont("Arial", 14), width=300) # Alterado para Arial
                combobox.set("") # Define como vazio por padrão
                combobox.grid(row=row_num, column=1, sticky="ew", padx=10, pady=5)
                self.entry_fields[label_text] = combobox_var
            elif field_type == "calendar":
                date_entry = DateEntry(form_frame, selectmode='day', date_pattern='dd/mm/yyyy',
                                       font=ctk.CTkFont("Arial", 14), # Alterado para Arial
                                       background='#046F91', foreground='white', bordercolor='#046F91',
                                       headersbackground='#046F91', headersforeground='white',
                                       normalbackground='white', normalforeground='black',
                                       selectbackground='#046F91', selectforeground='white')
                date_entry.grid(row=row_num, column=1, sticky="ew", padx=10, pady=5)
                self.entry_fields[label_text] = date_entry
                # Define a data atual como sugestão inicial
                date_entry.set_date(datetime.now())
            elif field_type == "dropdown": # Para Equipe, populado dinamicamente
                self.team_names = self._load_team_names()
                team_combobox_var = ctk.StringVar(form_frame)
                # Não define um valor padrão aqui, deixando-o vazio
                team_combobox_var.set("")
                self.team_dropdown_widget = ctk.CTkComboBox(form_frame, values=self.team_names if self.team_names else ["Nenhuma Equipe"],
                                                           variable=team_combobox_var, state="readonly",
                                                           font=ctk.CTkFont("Arial", 14), width=300) # Alterado para Arial
                self.team_dropdown_widget.grid(row=row_num, column=1, sticky="ew", padx=10, pady=5)
                self.entry_fields[label_text] = team_combobox_var
            else: # Campos de entrada de texto comuns
                entry = ctk.CTkEntry(form_frame, width=300, font=ctk.CTkFont("Arial", 14)) # Alterado para Arial
                entry.grid(row=row_num, column=1, sticky="ew", padx=10, pady=5)
                self.entry_fields[label_text] = entry
            row_num += 1
        
        # Botão Salvar
        ctk.CTkButton(self.frame_conteudo, text="Salvar Funcionário", command=self.save_employee_data,
                      width=200, height=40, fg_color="#046F91",
                      font=ctk.CTkFont("Arial", 16, weight="bold")).pack(pady=20) # Alterado para Arial


    def _load_team_names(self):
        """Carrega os nomes das equipes existentes nos arquivos de equipe."""
        teams = []
        if os.path.exists(self.pasta_equipes):
            for filename in os.listdir(self.pasta_equipes):
                if filename.endswith(".txt"):
                    team_name = os.path.splitext(filename)[0].replace("_", " ")
                    teams.append(team_name)
        return sorted(teams) # Retorna os nomes das equipes em ordem alfabética

    def save_employee_data(self):
        """Salva os dados do novo funcionário em um arquivo de texto."""
        data = {}
        for label, field_widget in self.entry_fields.items():
            if isinstance(field_widget, DateEntry):
                data[label] = field_widget.get()
            elif isinstance(field_widget, ctk.CTkTextbox): # Este caso não deve ocorrer mais com os campos removidos
                data[label] = field_widget.get("1.0", "end-1c").strip()
            else:
                data[label] = field_widget.get().strip() # Garante que todos os campos sejam limpos
        
        # Validação básica
        required_fields = ["Nome Completo", "Empresa", "Função", "Data de Início"] # CPF removido dos obrigatórios
        for field in required_fields:
            if not data.get(field):
                self.show_info_message_auto_close(f"Campo '{field}' obrigatório.", duration_ms=2000) # Mensagem concisa
                return

        # CPF e outros campos não serão mais validados aqui, pois foram removidos da interface
        cpf_numeric = "" # CPF será vazio ou não presente nos dados

        # Formatar Nome para o nome do arquivo (sem CPF no nome do arquivo, como era antes)
        employee_file_name = f"{data['Nome Completo'].replace(' ', '_').replace('/', '')}.txt"
        file_path = os.path.join(self.pasta_usuarios, employee_file_name)

        if os.path.exists(file_path):
            self.show_info_message_auto_close("Funcionário já cadastrado.", duration_ms=2000) # Mensagem concisa
            return

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                for key, value in data.items():
                    # Trata o campo de equipe para o arquivo, se "Nenhuma Equipe" estiver selecionado, salva como vazio
                    if key == "Equipe" and value == "Nenhuma Equipe":
                        f.write(f"{key}: \n")
                    else:
                        f.write(f"{key}: {value}\n")
            
            self.show_info_message_auto_close(f"Funcionário salvo com sucesso!", duration_ms=2000) # Mensagem concisa
            # Limpar campos após salvar
            for label, field_widget in self.entry_fields.items():
                if isinstance(field_widget, DateEntry):
                    field_widget.set_date(datetime.now()) # Resetar para a data atual
                elif isinstance(field_widget, ctk.CTkTextbox):
                    field_widget.delete("1.0", "end")
                elif isinstance(field_widget, (ctk.CTkEntry, ctk.CTkComboBox)):
                    if label in ["Empresa", "Função", "Usuário do meu Corban", "Usuário da Discadora Argus", "Equipe"]:
                        field_widget.set("") # Define como vazio após salvar
                        if label == "Equipe": # Recarrega a lista de equipes para o caso de novas equipes
                            self.team_names = self._load_team_names()
                            self.team_dropdown_widget.configure(values=self.team_names if self.team_names else ["Nenhuma Equipe"])
                    else:
                        field_widget.delete(0, tk.END)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar dados do funcionário: {e}")

    def _read_employee_data(self, file_path):
        """Lê os dados de um funcionário a partir de um arquivo."""
        data = {}
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    if ':' in line:
                        key, value = line.split(':', 1)
                        data[key.strip()] = value.strip()
        except Exception as e:
            print(f"Erro ao ler arquivo do funcionário {file_path}: {e}")
        return data

    # --- Telas de Desligamento ---
    def show_termination_screen(self):
        """Exibe a tela de desligamento de funcionários."""
        self.clear_content_frames()
        self.frame_conteudo = ctk.CTkFrame(self.master, fg_color=self.master.cget("fg_color"))
        self.frame_conteudo.grid(row=0, column=2, sticky="nsew", padx=50, pady=50)
        self.frame_conteudo.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            self.frame_conteudo,
            text="Desligamento de Funcionário",
            font=ctk.CTkFont("Georgia", 30), # Alterado
            text_color="white"
        ).pack(pady=(0, 20))

        # Campo de busca/sugestão
        search_frame = ctk.CTkFrame(self.frame_conteudo, fg_color="transparent")
        search_frame.pack(pady=10)

        ctk.CTkLabel(search_frame, text="Buscar Funcionário:", font=ctk.CTkFont("Arial", 16), text_color="white").pack(side="left", padx=10) # Alterado para Arial
        self.search_entry = ctk.CTkEntry(search_frame, width=300, font=ctk.CTkFont("Arial", 14)) # Alterado para Arial
        self.search_entry.pack(side="left", padx=10)
        self.search_entry.bind("<KeyRelease>", self._update_suggestions)

        self.employee_details_frame = ctk.CTkFrame(self.frame_conteudo, fg_color="#2b2b2b", corner_radius=10)
        self.employee_details_frame.pack(pady=20, fill="x", padx=10)
        
        self.employee_details_label = ctk.CTkLabel(self.employee_details_frame, text="Detalhes do Funcionário", font=ctk.CTkFont("Arial", 16, weight="bold"), text_color="white", anchor="w") # Alterado para Arial
        self.employee_details_label.pack(pady=10, padx=10)

        # Atualiza a lista de todos os funcionários ativos para sugestões
        self._load_all_employee_names()

        # Botão para efetivar o desligamento
        ctk.CTkButton(self.frame_conteudo, text="Efetivar Desligamento", command=self.confirm_termination,
                      width=250, height=40, fg_color="#990000", hover_color="#770000",
                      font=ctk.CTkFont("Arial", 16, weight="bold")).pack(pady=20) # Alterado para Arial
        
    def _load_all_employee_names(self):
        """Carrega todos os nomes de funcionários ativos e seus nomes de arquivo."""
        self.all_employee_names = []
        if os.path.exists(self.pasta_usuarios):
            for filename in os.listdir(self.pasta_usuarios):
                if filename.endswith(".txt"):
                    try:
                        file_path = os.path.join(self.pasta_usuarios, filename)
                        with open(file_path, 'r', encoding='utf-8') as f:
                            first_line = f.readline()
                            if first_line.startswith("Nome Completo:"):
                                name = first_line.split(":", 1)[1].strip()
                                self.all_employee_names.append({"name": name, "filename": filename})
                    except Exception as e:
                        print(f"Erro ao ler nome do arquivo {filename}: {e}")

    def _update_suggestions(self, event=None):
        """Atualiza a lista de sugestões de funcionários com base na entrada de busca."""
        search_term = self.search_entry.get().lower()
        if self.suggestion_frame:
            self.suggestion_frame.destroy()

        if search_term:
            self.suggestion_frame = ctk.CTkFrame(self.employee_details_frame, fg_color="#3e3e3e")
            self.suggestion_frame.pack(fill="x", padx=10, pady=5)

            found_suggestions = [
                item for item in self.all_employee_names if search_term in item["name"].lower()
            ]

            if found_suggestions:
                for item in found_suggestions:
                    suggestion_label = ctk.CTkLabel(self.suggestion_frame, text=item["name"], font=ctk.CTkFont("Arial", 14), text_color="white", anchor="w") # Alterado para Arial
                    suggestion_label.pack(fill="x", padx=5, pady=2)
                    suggestion_label.bind("<Button-1>", lambda e, data=item: self._select_employee(data))
            else:
                ctk.CTkLabel(self.suggestion_frame, text="Nenhum funcionário encontrado.", font=ctk.CTkFont("Arial", 14), text_color="gray").pack(padx=5, pady=5) # Alterado para Arial
        else:
            self.employee_details_label.configure(text="Detalhes do Funcionário")
            self.current_employee_data = None


    def _select_employee(self, employee_data):
        """Exibe os detalhes do funcionário selecionado e preenche a entrada de busca."""
        self.search_entry.delete(0, tk.END)
        self.search_entry.insert(0, employee_data["name"])
        if self.suggestion_frame:
            self.suggestion_frame.destroy()

        file_path = os.path.join(self.pasta_usuarios, employee_data["filename"])
        self.current_employee_data = self._read_employee_data(file_path)
        self._display_employee_details()

    def _display_employee_details(self):
        """Atualiza o label com os detalhes do funcionário selecionado."""
        if not self.current_employee_data:
            self.employee_details_label.configure(text="Detalhes do Funcionário")
            return

        details_text = "Detalhes do Funcionário:\n"
        for key, value in self.current_employee_data.items():
            details_text += f"{key}: {value}\n"
        
        self.employee_details_label.configure(text=details_text)

    def confirm_termination(self):
        """Confirma o desligamento do funcionário selecionado."""
        if not self.current_employee_data:
            self.show_info_message_auto_close("Selecione um funcionário para desligar.", duration_ms=2000) # Mensagem concisa
            return

        employee_name = self.current_employee_data.get("Nome Completo", "Funcionário Desconhecido")
        self.show_message_confirmation(f"Desligar '{employee_name}'?", # Mensagem concisa
                                        self._perform_termination)

    def _perform_termination(self):
        """Move o arquivo do funcionário para a pasta de desligados."""
        if not self.current_employee_data:
            return

        employee_name = self.current_employee_data.get("Nome Completo", "funcionario_desligado")
        # original_filename = f"{employee_name.replace(' ', '_').replace('/', '')}_{cpf_numeric}.txt"
        original_filename = f"{employee_name.replace(' ', '_').replace('/', '')}.txt"
        
        original_path = os.path.join(self.pasta_usuarios, original_filename)
        # Adiciona a data de desligamento ao nome do arquivo movido
        termination_date = datetime.now().strftime("%Y%m%d")
        terminated_filename = f"{employee_name.replace(' ', '_').replace('/', '')}_DESLIGADO_{termination_date}.txt"
        destination_path = os.path.join(self.pasta_desligados, terminated_filename)

        try:
            if os.path.exists(original_path):
                # Atualizar os dados do funcionário com a data de término antes de mover
                self.current_employee_data['Data de Término (Opcional)'] = datetime.now().strftime("%d/%m/%Y")
                
                # Salvar os dados atualizados no novo local
                with open(destination_path, 'w', encoding='utf-8') as f:
                    for key, value in self.current_employee_data.items():
                        f.write(f"{key}: {value}\n")

                os.remove(original_path) # Remove o arquivo original
                self.show_info_message_auto_close(f"Funcionário desligado com sucesso!", duration_ms=2000) # Mensagem concisa
                
                # Limpar a interface de desligamento ANTES de recarregar a tela principal
                self.search_entry.delete(0, tk.END)
                self.employee_details_label.configure(text="Detalhes do Funcionário")
                self.current_employee_data = None
                self._load_all_employee_names() # Recarrega a lista de ativos
                self.show_home_screen(force_reload=True) # Atualiza a tela inicial
            else:
                self.show_info_message_auto_close("Erro: Arquivo não encontrado.", duration_ms=2000) # Mensagem concisa
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao efetivar desligamento: {e}")

    # --- Funções de Fechamento ---
    def on_closing(self):
        """Ação ao fechar a janela principal."""
        self.show_message_confirmation("Deseja sair?", self.perform_exit) # Mensagem concisa

    def perform_exit(self):
        """Executa o salvamento e fecha a aplicação."""
        # Não é necessário self.master.focus_set() aqui, pois o grab_set/release do toplevel já gerencia o foco.
        self.show_saving_screen() # Mostra a tela de salvamento antes de sair
        # A destruição da janela principal agora é tratada dentro de _simulate_saving
        # para garantir que ocorra após a animação da barra de progresso.

    def show_saving_screen(self):
        """Exibe uma tela de salvamento com barra de progresso."""
        self.clear_content_frames()
        self.clear_sidebar_frames() # Garante que nada mais esteja visível

        self.saving_frame = ctk.CTkFrame(self.master, fg_color="#222222")
        self.saving_frame.grid(row=0, column=0, sticky="nsew", padx=0, pady=0, columnspan=3) # Ocupa toda a tela
        self.master.grid_columnconfigure((0, 1, 2), weight=1)
        self.master.grid_rowconfigure(0, weight=1)

        ctk.CTkLabel(
            self.saving_frame,
            text="Salvando e Encerrando...",
            font=ctk.CTkFont("Arial", 20, weight="bold"), # Tamanho da fonte reduzido, prioridade Arial
            text_color="white"
        ).pack(pady=(100, 50))

        # Barra de progresso - Largura e altura reduzidas
        self.progress_bar = ctk.CTkProgressBar(self.saving_frame, width=250, height=10, corner_radius=10, fg_color="#555555", progress_color="#046F91")
        self.progress_bar.pack(pady=20)
        self.progress_bar.set(0) # Inicia em 0%

        # Label para a porcentagem - Tamanho da fonte reduzido, prioridade Arial
        self.progress_percentage_label = ctk.CTkLabel(self.saving_frame, text="0%",
                                                       font=ctk.CTkFont("Arial", 14, weight="bold"), # Alterado
                                                       text_color="white")
        self.progress_percentage_label.pack(pady=10)

        # Inicia a simulação de salvamento de forma não bloqueante
        self._simulate_saving(100)

    def _simulate_saving(self, max_percentage, current_percentage=0):
        """Simula o preenchimento da barra de progresso e salva os dados de forma não bloqueante."""
        if current_percentage <= max_percentage:
            self.progress_bar.set(current_percentage / 100)
            self.progress_percentage_label.configure(text=f"{current_percentage}%")
            # Salva a quantidade de URA apenas uma vez, no início da simulação
            if current_percentage == 0:
                self._save_ura_quantity()
            # Agenda a próxima atualização da barra de progresso
            self.master.after(10, self._simulate_saving, max_percentage, current_percentage + 1)
        else:
            # Ao completar, destrói a janela principal
            self.master.destroy()

# --- Início da Aplicação ---
if __name__ == "__main__":
    import csv # Importar csv apenas quando necessário
    root = ctk.CTk()
    app = UserManagerApp(root)
    root.mainloop()
