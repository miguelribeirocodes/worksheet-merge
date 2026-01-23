"""Funções para descoberta e categorização dinâmica de colunas."""
import pandas as pd
from typing import List, Dict, Optional


def load_columns_from_excel(file_path: str, header_row: int = 1) -> List[str]:
    """
    Carrega os nomes das colunas reais de um arquivo Excel.

    Args:
        file_path: Caminho do arquivo Excel
        header_row: Número da linha que contém o header (0-indexed, padrão: 1 para segunda linha)

    Returns:
        Lista com os nomes das colunas presentes no arquivo

    Raises:
        FileNotFoundError: Se o arquivo não existe
        ValueError: Se o arquivo não é um Excel válido
    """
    try:
        df = pd.read_excel(file_path, header=header_row)
        return list(df.columns)
    except FileNotFoundError:
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")
    except Exception as e:
        raise ValueError(f"Erro ao ler arquivo Excel: {str(e)}")


def categorize_columns(columns: List[str], data_type: str) -> Dict[str, List[str]]:
    """
    Categoriza colunas automaticamente por tipo.

    Args:
        columns: Lista de nomes de colunas
        data_type: Tipo de dados ("pessoa" ou "registros")

    Returns:
        Dicionário com categorias e suas respectivas colunas
        {
            "Categoria 1": ["coluna1", "coluna2", ...],
            "Categoria 2": [...],
            ...
        }
    """
    if data_type.lower() == "pessoa":
        return _categorize_pessoa_columns(columns)
    elif data_type.lower() in ["registros", "registro", "registros_acesso"]:
        return _categorize_registros_columns(columns)
    else:
        # Tipo desconhecido, retorna como personalizado
        return {"Colunas": columns}


def _categorize_pessoa_columns(columns: List[str]) -> Dict[str, List[str]]:
    """Categoriza colunas de planilha de Pessoas."""

    categorized = {
        "Informações Básicas": [],
        "Documentação": [],
        "Trabalho": [],
        "Datas": [],
        "Contato": [],
        "Endereço": [],
        "Observações": [],
        "Personalizadas": []
    }

    # Palavras-chave para cada categoria
    categoria_map = {
        "Informações Básicas": [
            "Nome", "Sobrenome", "ID Pessoal", "Email", "Gênero",
            "Numero de Departamento", "Nome do Departamento",
            "Número de Departamento"
        ],
        "Documentação": [
            "Tipo de Documento", "Número do Documento", "Número do Cartão",
            "CPF", "RG", "CNH"
        ],
        "Trabalho": [
            "Nome do Cargo", "Número do Cargo", "Data de Contratação",
            "Título do Trabalho", "Título do Tr"
        ],
        "Datas": [
            "Data de Nascimento", "Data de Contratação", "Data de",
            "Nascimento"
        ],
        "Contato": [
            "Celular", "Telefone", "Email", "Telefone Comercial",
            "Telefone Residencial", "Telefone Re"
        ],
        "Endereço": [
            "Rua", "Endereço", "Endereço Comercial", "Endereço Residencial",
            "País", "Naturalidade"
        ],
        "Observações": [
            "Observação", "ASO", "Observação 1"
        ]
    }

    # Categorizar cada coluna
    for col in columns:
        categorized_flag = False

        for categoria, keywords in categoria_map.items():
            for keyword in keywords:
                if keyword.lower() in col.lower():
                    if col not in categorized[categoria]:
                        categorized[categoria].append(col)
                    categorized_flag = True
                    break

            if categorized_flag:
                break

        # Se não foi categorizada, adicionar em Personalizadas
        if not categorized_flag:
            categorized["Personalizadas"].append(col)

    # Remover categorias vazias
    return {k: v for k, v in categorized.items() if v}


def _categorize_registros_columns(columns: List[str]) -> Dict[str, List[str]]:
    """Categoriza colunas de planilha de Registros/Acessos."""

    categorized = {
        "Informações de Acesso": [],
        "Dados de Evento": [],
        "Dados de Pessoa": [],
        "Segurança": [],
        "Personalizadas": []
    }

    # Palavras-chave para cada categoria
    categoria_map = {
        "Informações de Acesso": [
            "Horário", "Nome da Área", "Nome do Dispositivo", "Descrição do Evento",
            "Ponto do Evento", "Ponto do Ev"
        ],
        "Dados de Evento": [
            "Nível do Evento", "Nível do Eve", "Arquivo de Mídia", "Arquivo de M",
            "ID do Evento"
        ],
        "Dados de Pessoa": [
            "ID Pessoal", "Nome", "Sobrenome", "Nome do Departamento",
            "Nome do Leitor", "Nome do Leit", "Número do Cartão"
        ],
        "Segurança": [
            "Modo de Verificação", "Criptografia", "Modo de Ver"
        ]
    }

    # Categorizar cada coluna
    for col in columns:
        categorized_flag = False

        for categoria, keywords in categoria_map.items():
            for keyword in keywords:
                if keyword.lower() in col.lower():
                    if col not in categorized[categoria]:
                        categorized[categoria].append(col)
                    categorized_flag = True
                    break

            if categorized_flag:
                break

        # Se não foi categorizada, adicionar em Personalizadas
        if not categorized_flag:
            categorized["Personalizadas"].append(col)

    # Remover categorias vazias
    return {k: v for k, v in categorized.items() if v}
