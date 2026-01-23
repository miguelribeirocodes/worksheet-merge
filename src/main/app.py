"""Aplicativo principal unificado com interface tipo ZKBio CVSecurity."""
import sys
import os
import tkinter as tk
from tkinter import messagebox, filedialog, ttk

# Adicionar o caminho do módulo utils ao path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from utils import (
    load_columns_from_excel,
    categorize_columns,
    MergeEngine,
    ConfigManager,
    CategoryFrame,
    ScrollableFrame,
    set_path,
    validar_entrada,
    validar_colunas_selecionadas
)
import pandas as pd


class MergeApp(tk.Tk):
    """Aplicativo unificado para merge de planilhas com interface de checkboxes."""

    def __init__(self):
        """Inicializa o aplicativo."""
        super().__init__()

        self.title("Mesclador de Planilhas ZKBio CVSecurity")
        self.geometry("900x900")
        self.resizable(True, True)

        # Inicializar manager de configurações
        self.config_manager = ConfigManager()
        self.merge_engine = MergeEngine()

        # Variáveis de controle
        self.path_pessoas = tk.StringVar()
        self.path_secundario = tk.StringVar()
        self.df_pessoas = None
        self.df_secundario = None
        self.colunas_categorias_pessoas = {}
        self.colunas_categorias_secundario = {}
        self.category_frames_pessoas = {}
        self.category_frames_secundario = {}

        # Criar widgets
        self._create_widgets()

    def _create_widgets(self):
        """Cria todos os widgets da interface."""

        # ===== FRAME SUPERIOR: SELEÇÃO DE ARQUIVOS =====
        frame_arquivos = tk.LabelFrame(
            self,
            text="SELEÇÃO DE ARQUIVOS",
            padx=10,
            pady=10,
            font=("Arial", 10, "bold")
        )
        frame_arquivos.pack(fill="x", padx=10, pady=10)

        # Arquivo de Pessoas
        tk.Label(frame_arquivos, text="Arquivo de Pessoas:").pack(anchor="w")
        frame_pessoas_path = tk.Frame(frame_arquivos)
        frame_pessoas_path.pack(fill="x", pady=5)
        tk.Entry(frame_pessoas_path, textvariable=self.path_pessoas, width=60).pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        tk.Button(
            frame_pessoas_path,
            text="Selecionar",
            command=self._select_pessoas_file,
            width=12
        ).pack(side="left")

        # Arquivo Secundário
        tk.Label(frame_arquivos, text="Arquivo Secundário:").pack(anchor="w", pady=(10, 0))
        frame_secundario_path = tk.Frame(frame_arquivos)
        frame_secundario_path.pack(fill="x", pady=5)
        tk.Entry(frame_secundario_path, textvariable=self.path_secundario, width=60).pack(
            side="left", fill="x", expand=True, padx=(0, 5)
        )
        tk.Button(
            frame_secundario_path,
            text="Selecionar",
            command=self._select_secundario_file,
            width=12
        ).pack(side="left")

        # ===== FRAME CENTRAL: SELEÇÃO DE COLUNAS =====
        frame_colunas = tk.LabelFrame(
            self,
            text="SELEÇÃO DE COLUNAS",
            padx=10,
            pady=10,
            font=("Arial", 10, "bold")
        )
        frame_colunas.pack(fill="both", expand=True, padx=10, pady=10)

        # Criar PanedWindow com 2 abas
        self.paned = ttk.PanedWindow(frame_colunas, orient="horizontal")
        self.paned.pack(fill="both", expand=True)

        # Aba 1: Pessoas
        frame_aba_pessoas = tk.LabelFrame(
            self.paned,
            text="Pessoas",
            padx=5,
            pady=5,
            font=("Arial", 9, "bold")
        )
        self.paned.add(frame_aba_pessoas)

        self.scrollable_pessoas = ScrollableFrame(frame_aba_pessoas)
        self.scrollable_pessoas.pack(fill="both", expand=True)

        # Label no aba pessoas
        self.label_pessoas = tk.Label(
            self.scrollable_pessoas.get_frame(),
            text="Selecione um arquivo de Pessoas para ver as colunas disponíveis.",
            fg="gray"
        )
        self.label_pessoas.pack()

        # Aba 2: Secundário
        frame_aba_secundario = tk.LabelFrame(
            self.paned,
            text="Registros / Níveis de Acesso",
            padx=5,
            pady=5,
            font=("Arial", 9, "bold")
        )
        self.paned.add(frame_aba_secundario)

        self.scrollable_secundario = ScrollableFrame(frame_aba_secundario)
        self.scrollable_secundario.pack(fill="both", expand=True)

        # Label no aba secundário
        self.label_secundario = tk.Label(
            self.scrollable_secundario.get_frame(),
            text="Selecione um arquivo Secundário para ver as colunas disponíveis.",
            fg="gray"
        )
        self.label_secundario.pack()

        # ===== FRAME INFERIOR: OPÇÕES =====
        frame_opcoes = tk.LabelFrame(
            self,
            text="OPÇÕES",
            padx=10,
            pady=10,
            font=("Arial", 10, "bold")
        )
        frame_opcoes.pack(fill="x", padx=10, pady=10)

        # Ordenação
        frame_ordenacao = tk.Frame(frame_opcoes)
        frame_ordenacao.pack(fill="x", pady=5)
        tk.Label(frame_ordenacao, text="Ordenar por coluna:").pack(side="left", padx=(0, 5))
        self.combo_sort = ttk.Combobox(
            frame_ordenacao,
            values=["Horário", "Nome do Nível"],
            state="readonly",
            width=20
        )
        self.combo_sort.set("Horário")
        self.combo_sort.pack(side="left", padx=(0, 10))

        self.var_sort_order = tk.StringVar(value="DESC")
        tk.Radiobutton(
            frame_ordenacao,
            text="Decrescente",
            variable=self.var_sort_order,
            value="DESC"
        ).pack(side="left", padx=5)
        tk.Radiobutton(
            frame_ordenacao,
            text="Crescente",
            variable=self.var_sort_order,
            value="ASC"
        ).pack(side="left", padx=5)

        # Salvar configuração
        frame_salvar_config = tk.Frame(frame_opcoes)
        frame_salvar_config.pack(fill="x", pady=5)
        tk.Label(frame_salvar_config, text="Salvar configuração com nome:").pack(side="left", padx=(0, 5))
        self.entry_config_name = tk.Entry(frame_salvar_config, width=30)
        self.entry_config_name.pack(side="left", padx=(0, 5))
        tk.Button(
            frame_salvar_config,
            text="Salvar",
            command=self._save_config,
            width=10
        ).pack(side="left")

        # Carregar configuração
        frame_carregar_config = tk.Frame(frame_opcoes)
        frame_carregar_config.pack(fill="x", pady=5)
        tk.Label(frame_carregar_config, text="Carregar configuração:").pack(side="left", padx=(0, 5))
        self.combo_configs = ttk.Combobox(
            frame_carregar_config,
            state="readonly",
            width=30
        )
        self.combo_configs.pack(side="left", padx=(0, 5))
        tk.Button(
            frame_carregar_config,
            text="Carregar",
            command=self._load_config,
            width=10
        ).pack(side="left", padx=2)
        tk.Button(
            frame_carregar_config,
            text="Excluir",
            command=self._delete_config,
            width=10
        ).pack(side="left")

        # ===== FRAME BOTÕES =====
        frame_botoes = tk.Frame(self)
        frame_botoes.pack(fill="x", padx=10, pady=10)

        tk.Button(
            frame_botoes,
            text="MESCLAR",
            command=self._merge,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=30,
            pady=10
        ).pack(side="left", padx=5)

        tk.Button(
            frame_botoes,
            text="SAIR",
            command=self.quit,
            bg="#f44336",
            fg="white",
            font=("Arial", 11, "bold"),
            padx=30,
            pady=10
        ).pack(side="left", padx=5)

    def _select_pessoas_file(self):
        """Seleciona arquivo de pessoas e carrega colunas."""
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
        )

        if path:
            self.path_pessoas.set(path)
            self._load_pessoas_columns()
            self._update_config_dropdown()

    def _select_secundario_file(self):
        """Seleciona arquivo secundário e carrega colunas."""
        path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xls *.xlsx"), ("All Files", "*.*")]
        )

        if path:
            self.path_secundario.set(path)
            self._load_secundario_columns()
            self._update_config_dropdown()

    def _load_pessoas_columns(self):
        """Carrega e categoriza colunas de pessoas."""
        try:
            path = self.path_pessoas.get()
            if not path:
                return

            # Carregar colunas
            colunas = load_columns_from_excel(path, header_row=1)
            self.df_pessoas = pd.read_excel(path, header=1)

            # Categorizar
            self.colunas_categorias_pessoas = categorize_columns(colunas, "pessoa")

            # Limpar frames antigos
            for frame in self.category_frames_pessoas.values():
                frame.destroy()
            self.category_frames_pessoas.clear()

            # Renderizar categorias
            frame_container = self.scrollable_pessoas.get_frame()
            for widget in frame_container.winfo_children():
                widget.destroy()

            for categoria, colunas_cat in self.colunas_categorias_pessoas.items():
                category_frame = CategoryFrame(
                    frame_container,
                    title=categoria,
                    columns=colunas_cat,
                    mandatory_column="ID Pessoal" if categoria == "Informações Básicas" else None
                )
                category_frame.pack(fill="x", padx=5, pady=5)
                self.category_frames_pessoas[categoria] = category_frame

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar colunas de Pessoas:\n{str(e)}")

    def _load_secundario_columns(self):
        """Carrega e categoriza colunas do arquivo secundário."""
        try:
            path = self.path_secundario.get()
            if not path:
                return

            # Carregar colunas
            colunas = load_columns_from_excel(path, header_row=1)
            self.df_secundario = pd.read_excel(path, header=1)

            # Detectar tipo baseado nas colunas
            if "Horário" in colunas:
                tipo = "registros"
            else:
                tipo = "registros"

            # Categorizar
            self.colunas_categorias_secundario = categorize_columns(colunas, tipo)

            # Limpar frames antigos
            for frame in self.category_frames_secundario.values():
                frame.destroy()
            self.category_frames_secundario.clear()

            # Renderizar categorias
            frame_container = self.scrollable_secundario.get_frame()
            for widget in frame_container.winfo_children():
                widget.destroy()

            for categoria, colunas_cat in self.colunas_categorias_secundario.items():
                category_frame = CategoryFrame(
                    frame_container,
                    title=categoria,
                    columns=colunas_cat,
                    mandatory_column="ID Pessoal" if "ID Pessoal" in colunas_cat else None
                )
                category_frame.pack(fill="x", padx=5, pady=5)
                self.category_frames_secundario[categoria] = category_frame

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar colunas do arquivo secundário:\n{str(e)}")

    def _get_selected_columns_pessoas(self):
        """Retorna colunas selecionadas de pessoas."""
        colunas = []
        for frame in self.category_frames_pessoas.values():
            colunas.extend(frame.get_selected_columns())
        return colunas

    def _get_selected_columns_secundario(self):
        """Retorna colunas selecionadas do arquivo secundário."""
        colunas = []
        for frame in self.category_frames_secundario.values():
            colunas.extend(frame.get_selected_columns())
        return colunas

    def _merge(self):
        """Executa o merge das planilhas."""
        try:
            # Validações básicas
            if not validar_entrada(
                self.path_pessoas.get(),
                self.path_secundario.get(),
                "arquivo secundário"
            ):
                return

            # Obter colunas selecionadas
            colunas_pessoas = self._get_selected_columns_pessoas()
            colunas_secundario = self._get_selected_columns_secundario()

            # Validar seleções
            if not validar_colunas_selecionadas(
                self.df_pessoas,
                self.df_secundario,
                colunas_pessoas,
                colunas_secundario,
                "arquivo secundário"
            ):
                return

            # Executar merge
            df_result = self.merge_engine.merge(
                self.path_pessoas.get(),
                self.path_secundario.get(),
                colunas_pessoas,
                colunas_secundario,
                sort_column=self.combo_sort.get(),
                sort_order=self.var_sort_order.get()
            )

            # Solicitar caminho para salvar
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
                initialfile="planilha_mesclada.xlsx"
            )

            if save_path:
                df_result.to_excel(save_path, index=False)
                messagebox.showinfo(
                    "Sucesso",
                    f"Planilhas mescladas com sucesso!\n\nArquivo salvo em:\n{save_path}"
                )

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar merge:\n{str(e)}")

    def _save_config(self):
        """Salva a configuração atual."""
        config_name = self.entry_config_name.get().strip()

        if not config_name:
            messagebox.showwarning("Aviso", "Digite um nome para a configuração")
            return

        colunas_pessoas = self._get_selected_columns_pessoas()
        colunas_secundario = self._get_selected_columns_secundario()

        if not colunas_pessoas or not colunas_secundario:
            messagebox.showwarning("Aviso", "Selecione colunas em ambos os painéis antes de salvar")
            return

        if self.config_manager.save_config(
            config_name,
            colunas_pessoas,
            colunas_secundario,
            self.combo_sort.get(),
            self.var_sort_order.get()
        ):
            messagebox.showinfo("Sucesso", f"Configuração '{config_name}' salva com sucesso!")
            self._update_config_dropdown()
            self.entry_config_name.delete(0, tk.END)
        else:
            messagebox.showerror("Erro", "Erro ao salvar configuração")

    def _load_config(self):
        """Carrega a configuração selecionada."""
        config_name = self.combo_configs.get()

        if not config_name:
            messagebox.showwarning("Aviso", "Selecione uma configuração")
            return

        config = self.config_manager.load_config(config_name)

        if config:
            # Carregar seleções de colunas
            for frame in self.category_frames_pessoas.values():
                frame.set_selected_columns(config.get("pessoas", []))

            for frame in self.category_frames_secundario.values():
                frame.set_selected_columns(config.get("secundario", []))

            # Carregar configurações de ordenação
            if config.get("sort_column"):
                self.combo_sort.set(config["sort_column"])
            if config.get("sort_order"):
                self.var_sort_order.set(config["sort_order"])

            messagebox.showinfo("Sucesso", f"Configuração '{config_name}' carregada!")

    def _delete_config(self):
        """Deleta a configuração selecionada."""
        config_name = self.combo_configs.get()

        if not config_name:
            messagebox.showwarning("Aviso", "Selecione uma configuração")
            return

        if messagebox.askyesno("Confirmar", f"Deletar configuração '{config_name}'?"):
            if self.config_manager.delete_config(config_name):
                messagebox.showinfo("Sucesso", f"Configuração '{config_name}' deletada!")
                self._update_config_dropdown()

    def _update_config_dropdown(self):
        """Atualiza dropdown de configurações."""
        configs = self.config_manager.list_configs()
        self.combo_configs["values"] = configs


def main():
    """Função principal que inicia o aplicativo."""
    app = MergeApp()
    app.mainloop()


if __name__ == "__main__":
    main()
