"""Funções auxiliares de interface para os scripts de mesclagem de planilhas."""
import tkinter as tk
from tkinter import filedialog
from typing import List, Optional


def set_path(entry_field):
    """
    Abre diálogo de seleção de arquivo e insere o caminho no campo de entrada.

    Args:
        entry_field: Campo tkinter Entry onde inserir o caminho
    """
    path = filedialog.askopenfilename()
    entry_field.delete(0, tk.END)
    entry_field.insert(0, path)


def gerar_texto_dicas_dinamico(colunas_pessoas, colunas_secundario, tipo_secundario="níveis de acesso"):
    """
    Gera o texto de dicas automaticamente baseado nas colunas configuradas.

    Args:
        colunas_pessoas: Lista de colunas obrigatórias em pessoas
        colunas_secundario: Lista de colunas obrigatórias no arquivo secundário
        tipo_secundario: Nome do arquivo secundário (ex: "níveis de acesso", "registros")

    Returns:
        str: Texto formatado com as dicas
    """
    colunas_pessoas_str = ", ".join(colunas_pessoas)
    colunas_secundario_str = ", ".join(colunas_secundario)

    return f"""ATENÇÃO!

Para evitar erros, siga as instruções abaixo:

• Selecione corretamente a planilha de pessoas e a planilha de {tipo_secundario};

• A planilha de pessoas deverá ter obrigatoriamente as seguintes colunas:
  - {colunas_pessoas_str}

• A planilha de {tipo_secundario} deverá ter obrigatoriamente as seguintes colunas:
  - {colunas_secundario_str}

• Não faça nenhuma alteração na planilha antes de utilizar no aplicativo. Deverá ser inserida a planilha exatamente da forma como ela é exportada do sistema ZKBio CVSecurity."""


class ScrollableFrame(tk.Frame):
    """Frame com scrollbar para acomodar muitos checkboxes."""

    def __init__(self, parent, **kwargs):
        """
        Cria um frame com scroll automático.

        Args:
            parent: Widget pai
        """
        super().__init__(parent, **kwargs)

        # Criar canvas e scrollbar
        self.canvas = tk.Canvas(self, bg="white", highlightthickness=0)
        scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")

        # Configurar eventos de resize
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        # Criar janela no canvas
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        # Layout
        self.canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Permitir scroll com mouse wheel
        self._bind_mousewheel()

    def _bind_mousewheel(self):
        """Permite scroll com mouse wheel."""
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        self.canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def get_frame(self) -> tk.Frame:
        """Retorna o frame interno para adicionar widgets."""
        return self.scrollable_frame


class CategoryFrame(tk.LabelFrame):
    """Frame com título de categoria e checkboxes organizados."""

    def __init__(
        self,
        parent,
        title: str,
        columns: List[str],
        mandatory_column: Optional[str] = None,
        **kwargs
    ):
        """
        Cria um LabelFrame com checkboxes para uma categoria.

        Args:
            parent: Widget pai
            title: Título da categoria (ex: "Informações Básicas")
            columns: Lista de nomes de colunas
            mandatory_column: Coluna que deve estar sempre selecionada e desabilitada
        """
        super().__init__(parent, text=title, padx=10, pady=10, **kwargs)

        self.columns = columns
        self.mandatory_column = mandatory_column
        self.var_dict = {}

        # Criar checkboxes para cada coluna
        for col in columns:
            var = tk.BooleanVar(value=col == mandatory_column)
            self.var_dict[col] = var

            checkbox = tk.Checkbutton(
                self,
                text=col,
                variable=var,
                state="disabled" if col == mandatory_column else "normal"
            )
            checkbox.pack(anchor="w", padx=5, pady=2)

    def get_selected_columns(self) -> List[str]:
        """
        Retorna lista de colunas com checkbox marcado.

        Returns:
            Lista de colunas selecionadas
        """
        return [
            col for col, var in self.var_dict.items()
            if var.get()
        ]

    def set_selected_columns(self, columns: List[str]) -> None:
        """
        Marca checkboxes de acordo com lista fornecida.

        Args:
            columns: Lista de colunas para marcar
        """
        for col, var in self.var_dict.items():
            # Manter coluna obrigatória sempre marcada
            if col == self.mandatory_column:
                var.set(True)
            else:
                var.set(col in columns)
