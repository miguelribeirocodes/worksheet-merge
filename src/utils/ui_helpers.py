"""Funções auxiliares de interface para os scripts de mesclagem de planilhas."""
import tkinter
from tkinter import filedialog


def set_path(entry_field):
    """
    Abre diálogo de seleção de arquivo e insere o caminho no campo de entrada.

    Args:
        entry_field: Campo tkinter Entry onde inserir o caminho
    """
    path = filedialog.askopenfilename()
    entry_field.delete(0, tkinter.END)
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
