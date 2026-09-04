"""
Organizador Inteligente de Planilhas - Interface Gráfica Desktop
Construída com CustomTkinter para Windows com suporte a Tema Escuro e Claro.
"""

import os
import sys
import json
import threading
import subprocess
from datetime import datetime
from pathlib import Path
from typing import List, Optional

import customtkinter as ctk
from tkinter import filedialog, messagebox

from engine import SpreadsheetOrganizer, THEMES

CONFIG_FILE = Path(__file__).resolve().parent / "config.json"


# Configuração inicial do tema CustomTkinter
ctk.set_appearance_mode("System")  # Segue o sistema (Dark ou Light)
ctk.set_default_color_theme("blue")


class ModernFileCard(ctk.CTkFrame):
    """Card individual de exibição de arquivo na fila."""

    def __init__(self, master, filepath: str, on_remove_callback, **kwargs):
        super().__init__(master, corner_radius=8, fg_color=("gray90", "#2B2D31"), **kwargs)
        self.filepath = filepath
        self.path_obj = Path(filepath)
        self.on_remove_callback = on_remove_callback

        self.grid_columnconfigure(1, weight=1)

        # Ícone / Badge de extensão
        ext = self.path_obj.suffix.lower()
        badge_color = "#10B981" if ext in [".xlsx", ".xls"] else "#3B82F6"
        self.badge = ctk.CTkLabel(
            self,
            text=ext.upper().replace(".", ""),
            font=ctk.CTkFont(size=11, weight="bold"),
            fg_color=badge_color,
            text_color="white",
            corner_radius=6,
            width=48,
            height=24
        )
        self.badge.grid(row=0, column=0, padx=(10, 8), pady=8)

        # Nome do arquivo e caminho
        info_frame = ctk.CTkFrame(self, fg_color="transparent")
        info_frame.grid(row=0, column=1, sticky="w", padx=4, pady=4)

        self.lbl_name = ctk.CTkLabel(
            info_frame,
            text=self.path_obj.name,
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w"
        )
        self.lbl_name.pack(anchor="w")

        size_kb = os.path.getsize(filepath) / 1024 if os.path.exists(filepath) else 0
        self.lbl_path = ctk.CTkLabel(
            info_frame,
            text=f"{size_kb:.1f} KB • {self.path_obj.parent}",
            font=ctk.CTkFont(size=10),
            text_color=("gray50", "gray60"),
            anchor="w"
        )
        self.lbl_path.pack(anchor="w")

        # Botão remover
        self.btn_remove = ctk.CTkButton(
            self,
            text="✕",
            width=28,
            height=28,
            fg_color="transparent",
            hover_color=("#FECDD3", "#881337"),
            text_color=("gray40", "gray70"),
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: self.on_remove_callback(self.filepath)
        )
        self.btn_remove.grid(row=0, column=2, padx=(4, 10), pady=8)

        # Abertura rápida do arquivo original via duplo clique
        for widget in (self, self.badge, info_frame, self.lbl_name, self.lbl_path):
            widget.bind("<Double-Button-1>", self._open_original_file)

    def _open_original_file(self, event=None):
        """Abre o arquivo original no aplicativo padrão do Windows."""
        if os.path.exists(self.filepath):
            os.startfile(self.filepath)


class SpreadsheetOrganizerApp(ctk.CTk):
    """Janela principal do Organizador Inteligente de Planilhas."""

    def __init__(self):
        super().__init__()

        self.title("Organizador Inteligente de Planilhas")
        self.geometry("1100x780")
        self.minsize(980, 680)

        # Estado da aplicação
        self.selected_files: List[str] = []
        self.custom_output_dir: Optional[str] = None
        self.last_organized_file: Optional[str] = None
        self.last_output_dir: Optional[str] = None
        self.is_processing = False
        self.cancel_requested = False
        self.is_scanning = False

        self._init_ui()
        self._load_config()
        self.protocol("WM_DELETE_WINDOW", self._on_closing)

    def _init_ui(self):
        """Monta toda a estrutura de componentes visuais."""
        self.grid_rowconfigure(1, weight=1)
        self.grid_columnconfigure(0, weight=1)

        # 1. Header / Barra Superior
        self._create_header()

        # 2. Área Central (Dividida em Esquerda: Arquivos | Direita: Configurações)
        main_content = ctk.CTkFrame(self, fg_color="transparent")
        main_content.grid(row=1, column=0, sticky="nsew", padx=20, pady=(10, 10))
        main_content.grid_columnconfigure(0, weight=3)
        main_content.grid_columnconfigure(1, weight=2)
        main_content.grid_rowconfigure(0, weight=1)

        self._create_files_panel(main_content)
        self._create_options_panel(main_content)

        # 3. Rodapé (Progresso, Logs e Ações Rápidas)
        self._create_footer()

    def _create_header(self):
        """Cabeçalho superior com título, badge e botão de alternância de tema."""
        header_frame = ctk.CTkFrame(self, fg_color=("gray95", "#1E1F22"), corner_radius=0, height=65)
        header_frame.grid(row=0, column=0, sticky="ew")
        header_frame.grid_columnconfigure(1, weight=1)

        # Logo / Título
        title_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        title_box.grid(row=0, column=0, padx=20, pady=12, sticky="w")

        lbl_logo = ctk.CTkLabel(
            title_box,
            text="📊 Organizador de Planilhas",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        lbl_logo.pack(side="left")

        badge_version = ctk.CTkLabel(
            title_box,
            text="v1.0 Pro",
            font=ctk.CTkFont(size=10, weight="bold"),
            fg_color=("#2563EB", "#1D4ED8"),
            text_color="white",
            corner_radius=10,
            padx=8,
            pady=2
        )
        badge_version.pack(side="left", padx=10)

        # Subtítulo
        lbl_subtitle = ctk.CTkLabel(
            header_frame,
            text="Organização visual, ajuste inteligente de colunas e estética executiva sem alterar dados.",
            font=ctk.CTkFont(size=12),
            text_color=("gray40", "gray60")
        )
        lbl_subtitle.grid(row=0, column=1, sticky="w", padx=10)

        # Alternador de Tema Dark/Light
        theme_box = ctk.CTkFrame(header_frame, fg_color="transparent")
        theme_box.grid(row=0, column=2, padx=20, pady=12, sticky="e")

        self.theme_switch = ctk.CTkSwitch(
            theme_box,
            text="Tema Escuro",
            command=self._toggle_theme,
            font=ctk.CTkFont(size=12)
        )
        self.theme_switch.select() if ctk.get_appearance_mode() == "Dark" else self.theme_switch.deselect()
        self.theme_switch.pack(side="right")

    def _create_files_panel(self, parent):
        """Painel esquerdo com a lista de arquivos e seleção."""
        files_frame = ctk.CTkFrame(parent, corner_radius=12, fg_color=("white", "#1E1F22"))
        files_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10), pady=0)
        files_frame.grid_rowconfigure(2, weight=1)
        files_frame.grid_columnconfigure(0, weight=1)

        # Cabeçalho do Painel
        head_box = ctk.CTkFrame(files_frame, fg_color="transparent")
        head_box.grid(row=0, column=0, sticky="ew", padx=15, pady=(15, 5))
        
        lbl_title = ctk.CTkLabel(
            head_box,
            text="📁 Planilhas para Organizar",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        lbl_title.pack(side="left")

        self.lbl_file_count = ctk.CTkLabel(
            head_box,
            text="0 arquivos selecionados",
            font=ctk.CTkFont(size=12),
            text_color=("gray50", "gray50")
        )
        self.lbl_file_count.pack(side="right")

        # Botões de Ação para adicionar arquivos
        btn_box = ctk.CTkFrame(files_frame, fg_color="transparent")
        btn_box.grid(row=1, column=0, sticky="ew", padx=15, pady=5)
        btn_box.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_add_files = ctk.CTkButton(
            btn_box,
            text="📄 Adicionar Arquivos",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=self._choose_files,
            height=34
        )
        self.btn_add_files.grid(row=0, column=0, padx=(0, 5), sticky="ew")

        self.btn_add_folder = ctk.CTkButton(
            btn_box,
            text="📂 Adicionar Pasta",
            font=ctk.CTkFont(size=12),
            fg_color=("gray80", "#2B2D31"),
            text_color=("gray10", "white"),
            hover_color=("gray70", "#3F4148"),
            command=self._choose_folder,
            height=34
        )
        self.btn_add_folder.grid(row=0, column=1, padx=5, sticky="ew")

        self.btn_clear = ctk.CTkButton(
            btn_box,
            text="🗑️ Limpar",
            font=ctk.CTkFont(size=12),
            fg_color=("gray85", "#2B2D31"),
            text_color=("gray20", "white"),
            hover_color=("#FEE2E2", "#7F1D1D"),
            command=self._clear_files,
            height=34
        )
        self.btn_clear.grid(row=0, column=2, padx=(5, 0), sticky="ew")

        # Scrollable Frame com os Cards dos arquivos
        self.scroll_files = ctk.CTkScrollableFrame(
            files_frame,
            fg_color=("gray95", "#18191C"),
            corner_radius=8
        )
        self.scroll_files.grid(row=2, column=0, sticky="nsew", padx=15, pady=10)

        # Placeholder quando vazio
        self.empty_label = ctk.CTkLabel(
            self.scroll_files,
            text="Nenhum arquivo adicionado ainda.\n\nClique em 'Adicionar Arquivos' (.xlsx, .xls, .csv)\nou 'Adicionar Pasta' para carregar em lote.",
            font=ctk.CTkFont(size=13),
            text_color=("gray50", "gray50"),
            justify="center"
        )
        self.empty_label.pack(pady=60)

        # Destino de Salvamento
        dest_frame = ctk.CTkFrame(files_frame, fg_color="transparent")
        dest_frame.grid(row=3, column=0, sticky="ew", padx=15, pady=(5, 15))
        dest_frame.grid_columnconfigure(0, weight=1)

        self.lbl_output = ctk.CTkLabel(
            dest_frame,
            text="📍 Destino: Mesma pasta do arquivo original (Salvo como [nome]_organizado.xlsx)",
            font=ctk.CTkFont(size=11),
            text_color=("gray40", "gray60"),
            anchor="w"
        )
        self.lbl_output.grid(row=0, column=0, sticky="w", pady=(0, 4))

        dest_btn_box = ctk.CTkFrame(dest_frame, fg_color="transparent")
        dest_btn_box.grid(row=1, column=0, sticky="ew")

        self.btn_dest = ctk.CTkButton(
            dest_btn_box,
            text="Alterar Pasta de Destino...",
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            border_width=1,
            border_color=("gray70", "gray40"),
            text_color=("gray20", "gray80"),
            height=28,
            command=self._choose_output_dir
        )
        self.btn_dest.pack(side="left")

        self.btn_reset_dest = ctk.CTkButton(
            dest_btn_box,
            text="Restaurar Padrão",
            font=ctk.CTkFont(size=11),
            fg_color="transparent",
            text_color=("gray40", "gray60"),
            height=28,
            command=self._reset_output_dir
        )
        self.btn_reset_dest.pack(side="left", padx=10)

    def _create_options_panel(self, parent):
        """Painel direito com as configurações de estilização e botão principal."""
        opt_frame = ctk.CTkFrame(parent, corner_radius=12, fg_color=("white", "#1E1F22"))
        opt_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0), pady=0)
        opt_frame.grid_rowconfigure(1, weight=1)
        opt_frame.grid_columnconfigure(0, weight=1)

        # Título
        lbl_title = ctk.CTkLabel(
            opt_frame,
            text="⚙️ Opções de Organização",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w"
        )
        lbl_title.grid(row=0, column=0, sticky="w", padx=20, pady=(15, 10))

        # Scrollable Frame de Opções
        scroll_opt = ctk.CTkScrollableFrame(opt_frame, fg_color="transparent")
        scroll_opt.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        # Paleta de Cores do Tema
        lbl_theme = ctk.CTkLabel(
            scroll_opt,
            text="Paleta de Cores da Planilha:",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w"
        )
        lbl_theme.pack(anchor="w", padx=10, pady=(5, 4))

        self.theme_var = ctk.StringVar(value="Azul Corporativo")
        self.theme_dropdown = ctk.CTkOptionMenu(
            scroll_opt,
            values=list(THEMES.keys()),
            variable=self.theme_var,
            command=lambda _: self._save_config(),
            font=ctk.CTkFont(size=12),
            height=34
        )
        self.theme_dropdown.pack(fill="x", padx=10, pady=(0, 15))

        # Separador / Switches de Personalização
        self.opt_vars = {}
        self.opt_switches = {}

        switches_data = [
            ("auto_column_width", "Ajuste Automático de Colunas", "Elimina textos cortados e erros numéricos (###)", True),
            ("header_style", "Cabeçalho Profissional", "Fundo executivo, texto em negrito e bordas suaves", True),
            ("freeze_panes", "Congelar Painel no Cabeçalho", "Mantém a 1ª linha fixa ao rolar a planilha", True),
            ("auto_filters", "Filtros Automáticos", "Aplica botões de filtro no cabeçalho das colunas", True),
            ("zebra_stripes", "Leitura Ergonômica (Zebra)", "Linhas alternadas com fundo suave para fácil leitura", True),
            ("smart_align", "Alinhamento Inteligente", "Textos à esq., números à dir., datas/códigos centrados", True),
            ("smart_formats", "Formatação de Dados", "Formata datas (DD/MM/AAAA) e valores (R$ / decimais)", True),
            ("show_gridlines", "Exibir Linhas de Grade", "Garante visualização nítida de grades no Excel", True),
        ]

        for key, title, desc, default_val in switches_data:
            var = ctk.BooleanVar(value=default_val)
            self.opt_vars[key] = var

            item_box = ctk.CTkFrame(scroll_opt, fg_color=("gray95", "#2B2D31"), corner_radius=8)
            item_box.pack(fill="x", padx=5, pady=4)

            switch = ctk.CTkSwitch(
                item_box,
                text=title,
                variable=var,
                command=self._save_config,
                font=ctk.CTkFont(size=12, weight="bold"),
            )
            switch.pack(anchor="w", padx=10, pady=(8, 2))
            self.opt_switches[key] = switch

            lbl_sub = ctk.CTkLabel(
                item_box,
                text=desc,
                font=ctk.CTkFont(size=10),
                text_color=("gray50", "gray60")
            )
            lbl_sub.pack(anchor="w", padx=10, pady=(0, 8))

        # Botão Principal de Processamento e Cancelamento
        btn_action_box = ctk.CTkFrame(opt_frame, fg_color="transparent")
        btn_action_box.grid(row=2, column=0, sticky="ew", padx=15, pady=15)

        self.btn_process = ctk.CTkButton(
            btn_action_box,
            text="✨ ORGANIZAR PLANILHAS AGORA",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=("#2563EB", "#1D4ED8"),
            hover_color=("#1D4ED8", "#1E40AF"),
            height=48,
            corner_radius=10,
            command=self._start_processing
        )
        self.btn_process.pack(fill="x")

        self.btn_cancel = ctk.CTkButton(
            btn_action_box,
            text="🛑 Cancelar Processamento",
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=("#DC2626", "#991B1B"),
            hover_color=("#B91C1C", "#7F1D1D"),
            height=36,
            corner_radius=8,
            command=self._cancel_processing
        )

    def _create_footer(self):
        """Rodapé inferior com console de logs, barra de progresso e botões de atalho."""
        footer_frame = ctk.CTkFrame(self, fg_color=("gray95", "#1E1F22"), corner_radius=0)
        footer_frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        footer_frame.grid_columnconfigure(0, weight=1)

        # Barra de Progresso e Status
        status_box = ctk.CTkFrame(footer_frame, fg_color="transparent")
        status_box.grid(row=0, column=0, sticky="ew", padx=20, pady=(10, 4))
        status_box.grid_columnconfigure(0, weight=1)

        self.lbl_status = ctk.CTkLabel(
            status_box,
            text="Pronto para organizar planilhas.",
            font=ctk.CTkFont(size=12, weight="bold"),
            anchor="w"
        )
        self.lbl_status.grid(row=0, column=0, sticky="w")

        self.lbl_progress_pct = ctk.CTkLabel(
            status_box,
            text="0%",
            font=ctk.CTkFont(size=11),
            text_color=("gray50", "gray50")
        )
        self.lbl_progress_pct.grid(row=0, column=1, sticky="e")

        self.progress_bar = ctk.CTkProgressBar(footer_frame, height=8, corner_radius=4)
        self.progress_bar.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 8))
        self.progress_bar.set(0)

        # Console de Logs e Ações Rápidas
        log_action_frame = ctk.CTkFrame(footer_frame, fg_color="transparent")
        log_action_frame.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 12))
        log_action_frame.grid_columnconfigure(0, weight=1)

        # Caixa de Logs
        self.txt_logs = ctk.CTkTextbox(
            log_action_frame,
            height=90,
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=("white", "#18191C"),
            corner_radius=8
        )
        self.txt_logs.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.txt_logs.insert("end", f"[{self._now()}] Organizador Inteligente pronto. Selecione planilhas para começar.\n")
        self.txt_logs.configure(state="disabled")

        # Botões de Ação Rápida
        actions_box = ctk.CTkFrame(log_action_frame, fg_color="transparent")
        actions_box.grid(row=0, column=1, sticky="ns")

        self.btn_open_folder = ctk.CTkButton(
            actions_box,
            text="📁 Abrir Pasta",
            font=ctk.CTkFont(size=12),
            width=140,
            height=40,
            fg_color=("gray85", "#2B2D31"),
            text_color=("gray20", "white"),
            state="disabled",
            command=self._open_output_folder
        )
        self.btn_open_folder.pack(pady=(0, 6))

        self.btn_open_file = ctk.CTkButton(
            actions_box,
            text="📊 Abrir Arquivo",
            font=ctk.CTkFont(size=12),
            width=140,
            height=40,
            fg_color=("gray85", "#2B2D31"),
            text_color=("gray20", "white"),
            state="disabled",
            command=self._open_last_file
        )
        self.btn_open_file.pack()

    # --- Métodos de Interação com Arquivos ---

    def _choose_files(self):
        """Abre diálogo para seleção de múltiplos arquivos."""
        files = filedialog.askopenfilenames(
            title="Selecione as Planilhas",
            filetypes=[
                ("Planilhas Suportadas", "*.xlsx *.xls *.csv"),
                ("Excel (.xlsx)", "*.xlsx"),
                ("Excel Antigo (.xls)", "*.xls"),
                ("Arquivos CSV (.csv)", "*.csv"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if files:
            for f in files:
                f_norm = os.path.abspath(f)
                if f_norm not in self.selected_files:
                    self.selected_files.append(f_norm)
            self._refresh_file_list()

    def _choose_folder(self):
        """Varre uma pasta selecionada em busca de arquivos suportados de forma assíncrona."""
        if self.is_processing or self.is_scanning:
            return

        folder = filedialog.askdirectory(title="Selecione uma Pasta com Planilhas")
        if not folder:
            return

        self.is_scanning = True
        self.btn_add_files.configure(state="disabled")
        self.btn_add_folder.configure(state="disabled")
        self.btn_clear.configure(state="disabled")
        self.btn_process.configure(state="disabled")

        self.lbl_status.configure(text="Buscando planilhas...")
        self._log(f"Buscando planilhas em '{folder}'...")

        thread = threading.Thread(
            target=self._scan_folder_worker,
            args=(folder,),
            daemon=True
        )
        thread.start()

    def _scan_folder_worker(self, folder: str):
        """Worker em thread secundária para busca de arquivos sem travar a interface."""
        found_files = []
        try:
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.startswith("~$") or f.endswith("_organizado.xlsx"):
                        continue
                    ext = Path(f).suffix.lower()
                    if ext in [".xlsx", ".xls", ".csv"]:
                        full_path = os.path.abspath(os.path.join(root, f))
                        found_files.append(full_path)
        except Exception as e:
            self._log(f"Erro durante a varredura da pasta: {str(e)}")

        self.after(0, self._scan_folder_finished, found_files)

    def _scan_folder_finished(self, found_files: List[str]):
        """Atualiza a UI na thread principal após a conclusão da varredura da pasta."""
        self.is_scanning = False
        added = 0
        for full_path in found_files:
            if full_path not in self.selected_files:
                self.selected_files.append(full_path)
                added += 1

        self.btn_add_files.configure(state="normal")
        self.btn_add_folder.configure(state="normal")
        self.btn_clear.configure(state="normal")
        self.btn_process.configure(state="normal")

        self.lbl_status.configure(text=f"Varredura concluída: {added} novas planilhas encontradas.")
        self._log(f"Pasta adicionada: {added} planilhas encontradas.")
        self._refresh_file_list()

    def _choose_output_dir(self):
        """Altera a pasta onde os arquivos organizados serão salvos."""
        folder = filedialog.askdirectory(title="Selecione a Pasta de Saída")
        if folder:
            self.custom_output_dir = os.path.abspath(folder)
            self.lbl_output.configure(
                text=f"📍 Destino Personalizado: {self.custom_output_dir}"
            )
            self._save_config()
            self._log(f"Pasta de saída definida: {self.custom_output_dir}")

    def _reset_output_dir(self):
        """Restaura o salvamento para a pasta original de cada arquivo."""
        self.custom_output_dir = None
        self.lbl_output.configure(
            text="📍 Destino: Mesma pasta do arquivo original (Salvo como [nome]_organizado.xlsx)"
        )
        self._save_config()
        self._log("Pasta de saída redefinida para o diretório de origem de cada arquivo.")

    def _remove_file(self, filepath: str):
        """Remove um arquivo específico da fila."""
        if filepath in self.selected_files:
            self.selected_files.remove(filepath)
            self._refresh_file_list()

    def _clear_files(self):
        """Limpa toda a fila de arquivos."""
        self.selected_files.clear()
        self._refresh_file_list()
        self._log("Lista de arquivos limpa.")

    def _refresh_file_list(self):
        """Atualiza visualmente os cards da lista de arquivos com limite de exibição."""
        for widget in self.scroll_files.winfo_children():
            widget.destroy()

        count = len(self.selected_files)
        self.lbl_file_count.configure(
            text=f"{count} arquivo{'s' if count != 1 else ''} selecionado{'s' if count != 1 else ''}"
        )

        if not self.selected_files:
            self.empty_label = ctk.CTkLabel(
                self.scroll_files,
                text="Nenhum arquivo adicionado ainda.\n\nClique em 'Adicionar Arquivos' (.xlsx, .xls, .csv)\nou 'Adicionar Pasta' para carregar em lote.",
                font=ctk.CTkFont(size=13),
                text_color=("gray50", "gray50"),
                justify="center"
            )
            self.empty_label.pack(pady=60)
        else:
            max_cards = 100
            for f in self.selected_files[:max_cards]:
                card = ModernFileCard(self.scroll_files, f, on_remove_callback=self._remove_file)
                card.pack(fill="x", padx=4, pady=3)

            if count > max_cards:
                hidden_count = count - max_cards
                lbl_hidden = ctk.CTkLabel(
                    self.scroll_files,
                    text=f"... e mais {hidden_count} arquivos ocultos na visualização, mas prontos para processar",
                    font=ctk.CTkFont(size=12, slant="italic"),
                    text_color=("gray50", "gray60")
                )
                lbl_hidden.pack(fill="x", padx=4, pady=10)

    # --- Execução do Processamento em Background ---

    def _start_processing(self):
        """Inicia o processamento seguro das planilhas em uma thread separada."""
        if not self.selected_files:
            messagebox.showwarning(
                "Nenhum arquivo selecionado",
                "Por favor, adicione ao menos uma planilha (.xlsx, .xls, .csv) antes de iniciar."
            )
            return

        if self.is_processing:
            return

        self.cancel_requested = False
        self.is_processing = True
        self.btn_process.configure(state="disabled", text="⏳ Processando Planilhas...")
        self.btn_cancel.configure(state="normal", text="🛑 Cancelar Processamento")
        self.btn_cancel.pack(fill="x", pady=(8, 0))

        self.btn_add_files.configure(state="disabled")
        self.btn_add_folder.configure(state="disabled")
        self.btn_clear.configure(state="disabled")
        self.progress_bar.set(0)
        self.lbl_progress_pct.configure(text="0%")

        self._save_config()

        # Opções selecionadas
        options = {key: var.get() for key, var in self.opt_vars.items()}
        options["theme"] = self.theme_var.get()

        thread = threading.Thread(
            target=self._process_worker,
            args=(list(self.selected_files), self.custom_output_dir, options),
            daemon=True
        )
        thread.start()

    def _cancel_processing(self):
        """Solicita a interrupção do processamento em andamento."""
        if self.is_processing:
            self.cancel_requested = True
            self.btn_cancel.configure(state="disabled", text="Interrompendo...")
            self.lbl_status.configure(text="Cancelando processamento...")
            self._log("Cancelamento solicitado. Abortando após a planilha atual...")

    def _process_worker(self, files: List[str], output_dir: Optional[str], options: dict):
        """Worker que roda na thread secundária."""
        total = len(files)
        success_count = 0
        organizer = SpreadsheetOrganizer(options)

        self._log(f"--- Iniciando organização de {total} arquivo(s) [Tema: {options.get('theme')}] ---")

        for idx, filepath in enumerate(files, start=1):
            if self.cancel_requested:
                self._log("Processamento cancelado pelo usuário.")
                self.after(0, self._processing_finished, success_count, total)
                return

            file_name = Path(filepath).name
            base_progress = (idx - 1) / total

            def file_progress_cb(message: str, pct_within_file: float):
                overall_pct = base_progress + (pct_within_file / total)
                self.after(0, self._update_ui_progress, message, overall_pct)

            try:
                out_file = organizer.process_file(
                    filepath,
                    output_dir=output_dir,
                    progress_callback=file_progress_cb
                )
                success_count += 1
                self.last_organized_file = out_file
                self.last_output_dir = str(Path(out_file).parent)
                self._log(f"✔ Concluído: {file_name} -> {Path(out_file).name}")
            except Exception as e:
                self._log(f"✖ Erro em '{file_name}': {str(e)}")

        self.after(0, self._processing_finished, success_count, total)

    def _update_ui_progress(self, message: str, pct: float):
        """Atualiza a barra de progresso e mensagem na UI principal."""
        self.lbl_status.configure(text=message)
        self.progress_bar.set(min(pct, 1.0))
        self.lbl_progress_pct.configure(text=f"{int(pct * 100)}%")

    def _processing_finished(self, success_count: int, total: int):
        """Restaura o estado da UI ao finalizar ou cancelar."""
        was_cancelled = self.cancel_requested
        self.is_processing = False
        self.cancel_requested = False
        self.btn_cancel.pack_forget()
        self.btn_process.configure(state="normal", text="✨ ORGANIZAR PLANILHAS AGORA")
        self.btn_add_files.configure(state="normal")
        self.btn_add_folder.configure(state="normal")
        self.btn_clear.configure(state="normal")

        if self.last_output_dir:
            self.btn_open_folder.configure(state="normal", fg_color=("#2563EB", "#1D4ED8"))
        if self.last_organized_file:
            self.btn_open_file.configure(state="normal", fg_color=("#10B981", "#059669"))

        if was_cancelled:
            self.lbl_status.configure(text=f"Processamento cancelado ({success_count}/{total} concluídas).")
            self._log(f"--- Processamento cancelado pelo usuário: {success_count}/{total} arquivos organizados ---")
            messagebox.showinfo(
                "Processamento Cancelado",
                f"O processamento foi interrompido com sucesso.\n\n{success_count} de {total} planilhas foram organizadas."
            )
            return

        self.progress_bar.set(1.0)
        self.lbl_progress_pct.configure(text="100%")
        self.lbl_status.configure(text=f"Processamento concluído! ({success_count}/{total} com sucesso)")

        self._log(f"--- Fim do Processo: {success_count}/{total} arquivos organizados com sucesso! ---")

        if success_count == total:
            messagebox.showinfo(
                "Sucesso!",
                f"Todas as {total} planilhas foram organizadas com sucesso!\nOs arquivos originais permaneceram 100% intactos."
            )
        else:
            messagebox.showwarning(
                "Processamento Concluído com Avisos",
                f"{success_count} de {total} planilhas foram processadas com sucesso. Verifique o log para detalhes."
            )

    # --- Ações Rápidas do Rodapé ---

    def _open_output_folder(self):
        """Abre a pasta onde os arquivos organizados foram salvos no Windows Explorer."""
        target_dir = self.last_output_dir or self.custom_output_dir
        if target_dir and os.path.exists(target_dir):
            os.startfile(target_dir) if hasattr(os, "startfile") else subprocess.run(["explorer", target_dir])
        else:
            messagebox.showinfo("Aviso", "Pasta de saída ainda não disponível.")

    def _open_last_file(self):
        """Abre a última planilha organizada no aplicativo padrão (Excel)."""
        if self.last_organized_file and os.path.exists(self.last_organized_file):
            os.startfile(self.last_organized_file) if hasattr(os, "startfile") else subprocess.run(["cmd", "/c", "start", "", self.last_organized_file])
        else:
            messagebox.showinfo("Aviso", "Nenhum arquivo organizado disponível para abrir.")

    # --- Utilitários ---

    def _toggle_theme(self):
        """Alterna entre Tema Escuro e Claro."""
        mode = "Dark" if self.theme_switch.get() == 1 else "Light"
        ctk.set_appearance_mode(mode)
        self._save_config()
        self._log(f"Tema alterado para: {mode}")

    def _save_config(self):
        """Salva as preferências do usuário no arquivo config.json."""
        try:
            config = {
                "theme": self.theme_var.get(),
                "appearance_mode": ctk.get_appearance_mode(),
                "custom_output_dir": self.custom_output_dir,
                "options": {key: var.get() for key, var in self.opt_vars.items()}
            }
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
        except Exception:
            pass

    def _load_config(self):
        """Carrega as preferências salvas no config.json se existente."""
        if not CONFIG_FILE.exists():
            return

        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                config = json.load(f)

            saved_theme = config.get("theme")
            if saved_theme and saved_theme in THEMES:
                self.theme_var.set(saved_theme)
                self.theme_dropdown.set(saved_theme)

            saved_mode = config.get("appearance_mode")
            if saved_mode in ["Dark", "Light"]:
                ctk.set_appearance_mode(saved_mode)
                if saved_mode == "Dark":
                    self.theme_switch.select()
                else:
                    self.theme_switch.deselect()

            saved_output_dir = config.get("custom_output_dir")
            if saved_output_dir and os.path.exists(saved_output_dir):
                self.custom_output_dir = os.path.abspath(saved_output_dir)
                self.lbl_output.configure(
                    text=f"📍 Destino Personalizado: {self.custom_output_dir}"
                )

            saved_options = config.get("options") or config.get("switches") or {}
            for key, val in saved_options.items():
                if key in self.opt_vars and isinstance(val, bool):
                    self.opt_vars[key].set(val)
                    if hasattr(self, "opt_switches") and key in self.opt_switches:
                        if val:
                            self.opt_switches[key].select()
                        else:
                            self.opt_switches[key].deselect()

        except Exception:
            pass

    def _on_closing(self):
        """Salva preferências ao fechar a janela."""
        self._save_config()
        self.destroy()

    def _log(self, text: str):
        """Escreve uma linha no console de logs da UI."""
        def append():
            self.txt_logs.configure(state="normal")
            self.txt_logs.insert("end", f"[{self._now()}] {text}\n")
            self.txt_logs.see("end")
            self.txt_logs.configure(state="disabled")
        self.after(0, append)

    @staticmethod
    def _now() -> str:
        return datetime.now().strftime("%H:%M:%S")


def main():
    app = SpreadsheetOrganizerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
