"""Funções de validação para os scripts de mesclagem de planilhas."""
import os
from tkinter import messagebox


def validar_entrada(path_pessoas, path_arquivo_secundario, tipo_arquivo="acesso"):
    """
    Valida se os arquivos existem e são arquivos Excel válidos.

    Args:
        path_pessoas: Caminho do arquivo de pessoas
        path_arquivo_secundario: Caminho do arquivo secundário (acessos ou registros)
        tipo_arquivo: Tipo do arquivo secundário para mensagens customizadas

    Returns:
        bool: True se válido, False caso contrário
    """
    if not path_pessoas or not path_arquivo_secundario:
        messagebox.showerror("Erro", "Selecione ambas as planilhas.")
        return False

    if not os.path.exists(path_pessoas):
        messagebox.showerror("Erro", f"Arquivo não encontrado: {path_pessoas}")
        return False

    if not os.path.exists(path_arquivo_secundario):
        messagebox.showerror("Erro", f"Arquivo não encontrado: {path_arquivo_secundario}")
        return False

    if not (path_pessoas.endswith('.xlsx') or path_pessoas.endswith('.xls')):
        messagebox.showerror("Erro", "A planilha de pessoas deve ser um arquivo Excel (.xlsx ou .xls)")
        return False

    if not (path_arquivo_secundario.endswith('.xlsx') or path_arquivo_secundario.endswith('.xls')):
        messagebox.showerror("Erro", f"A planilha de {tipo_arquivo} deve ser um arquivo Excel (.xlsx ou .xls)")
        return False

    return True


def validar_colunas(df_pessoas, df_secundario, colunas_obrigatorias_pessoas, colunas_obrigatorias_secundario, tipo_secundario="acesso"):
    """
    Valida se as colunas obrigatórias existem nos DataFrames.

    Args:
        df_pessoas: DataFrame da planilha de pessoas
        df_secundario: DataFrame da planilha secundária
        colunas_obrigatorias_pessoas: Lista de colunas obrigatórias em pessoas
        colunas_obrigatorias_secundario: Lista de colunas obrigatórias no arquivo secundário
        tipo_secundario: Nome do arquivo secundário para mensagens customizadas

    Returns:
        bool: True se válido, False caso contrário
    """
    colunas_faltando_pessoas = [col for col in colunas_obrigatorias_pessoas if col not in df_pessoas.columns]
    if colunas_faltando_pessoas:
        messagebox.showerror("Erro", f"Colunas faltando na planilha de pessoas:\n{', '.join(colunas_faltando_pessoas)}")
        return False

    colunas_faltando_secundario = [col for col in colunas_obrigatorias_secundario if col not in df_secundario.columns]
    if colunas_faltando_secundario:
        messagebox.showerror("Erro", f"Colunas faltando na planilha de {tipo_secundario}:\n{', '.join(colunas_faltando_secundario)}")
        return False

    return True


def validar_colunas_selecionadas(
    df_pessoas,
    df_secundario,
    colunas_pessoas,
    colunas_secundario,
    tipo="registros"
):
    """
    Valida se as colunas selecionadas existem nos DataFrames.

    Args:
        df_pessoas: DataFrame da planilha de pessoas
        df_secundario: DataFrame da planilha secundária
        colunas_pessoas: Lista de colunas selecionadas em pessoas
        colunas_secundario: Lista de colunas selecionadas no arquivo secundário
        tipo: Nome do arquivo secundário para mensagens customizadas

    Returns:
        bool: True se válido, False caso contrário
    """
    # Validar que pelo menos uma coluna foi selecionada
    if not colunas_pessoas:
        messagebox.showerror("Erro", "Selecione pelo menos uma coluna de Pessoas")
        return False

    if not colunas_secundario:
        messagebox.showerror("Erro", f"Selecione pelo menos uma coluna de {tipo}")
        return False

    # Validar que "ID Pessoal" está em ambas seleções
    if "ID Pessoal" not in colunas_pessoas:
        messagebox.showerror("Erro", "'ID Pessoal' deve estar selecionado em Pessoas")
        return False

    if "ID Pessoal" not in colunas_secundario:
        messagebox.showerror("Erro", f"'ID Pessoal' deve estar selecionado em {tipo}")
        return False

    # Validar que colunas existem nos DataFrames
    colunas_faltando_pessoas = [col for col in colunas_pessoas if col not in df_pessoas.columns]
    if colunas_faltando_pessoas:
        messagebox.showerror("Erro", f"Colunas não encontradas em Pessoas:\n{', '.join(colunas_faltando_pessoas)}")
        return False

    colunas_faltando_secundario = [col for col in colunas_secundario if col not in df_secundario.columns]
    if colunas_faltando_secundario:
        messagebox.showerror("Erro", f"Colunas não encontradas em {tipo}:\n{', '.join(colunas_faltando_secundario)}")
        return False

    return True
