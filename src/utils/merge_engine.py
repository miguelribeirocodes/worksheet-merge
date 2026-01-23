"""Engine para merge parametrizado de planilhas com SQLite."""
import pandas as pd
import sqlite3
import tempfile
import os
from typing import List, Optional, Tuple


class MergeEngine:
    """Engine para realizar merge dinâmico de duas planilhas via SQLite."""

    def merge(
        self,
        path_pessoas: str,
        path_secundario: str,
        selected_columns_pessoas: List[str],
        selected_columns_secundario: List[str],
        sort_column: Optional[str] = None,
        sort_order: str = "DESC"
    ) -> pd.DataFrame:
        """
        Realiza merge dinâmico de duas planilhas usando LEFT JOIN.

        Args:
            path_pessoas: Caminho do arquivo de pessoas
            path_secundario: Caminho do arquivo secundário (níveis ou registros)
            selected_columns_pessoas: Lista de colunas selecionadas da planilha de pessoas
            selected_columns_secundario: Lista de colunas selecionadas do arquivo secundário
            sort_column: Coluna para ordenação (opcional)
            sort_order: Ordem de ordenação ("ASC" ou "DESC", padrão: "DESC")

        Returns:
            DataFrame com os dados mesclados

        Raises:
            ValueError: Se houver erro na validação ou processamento
            FileNotFoundError: Se os arquivos não existem
        """
        db_path = None

        try:
            # 1. Carregar as planilhas
            df_pessoas = self._load_excel(path_pessoas, header_row=1)
            df_secundario = self._load_excel(path_secundario, header_row=1)

            # 2. Validar colunas selecionadas
            self._validate_selected_columns(
                df_pessoas, df_secundario,
                selected_columns_pessoas, selected_columns_secundario
            )

            # 3. Validar que "ID Pessoal" está em ambas seleções
            if "ID Pessoal" not in selected_columns_pessoas:
                raise ValueError("'ID Pessoal' deve estar selecionado em Pessoas")
            if "ID Pessoal" not in selected_columns_secundario:
                raise ValueError("'ID Pessoal' deve estar selecionado em Registros/Níveis")

            # 4. Criar banco de dados temporário
            temp_db = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
            db_path = temp_db.name
            temp_db.close()

            # 5. Inserir dados em tabelas SQLite
            conn = sqlite3.connect(db_path)
            df_pessoas.to_sql('Pessoas', conn, if_exists='replace', index=False)
            df_secundario.to_sql('Secundario', conn, if_exists='replace', index=False)

            # 6. Construir SELECT dinamicamente
            colunas_sql = self._build_select_list(
                selected_columns_secundario,
                selected_columns_pessoas
            )

            # 7. Construir e executar query
            query = self._build_query(colunas_sql, sort_column, sort_order)
            df_result = pd.read_sql_query(query, conn)
            conn.close()

            return df_result

        except FileNotFoundError as e:
            raise FileNotFoundError(f"Arquivo não encontrado: {str(e)}")
        except ValueError as e:
            raise ValueError(f"Erro de validação: {str(e)}")
        except Exception as e:
            raise ValueError(f"Erro ao processar merge: {str(e)}")
        finally:
            # 8. Limpar arquivo temporário
            if db_path and os.path.exists(db_path):
                try:
                    os.remove(db_path)
                except:
                    pass  # Ignorar erro ao deletar arquivo temporário

    @staticmethod
    def _load_excel(file_path: str, header_row: int = 1) -> pd.DataFrame:
        """
        Carrega um arquivo Excel com tratamento de erros.

        Args:
            file_path: Caminho do arquivo
            header_row: Linha que contém o header (0-indexed)

        Returns:
            DataFrame com os dados

        Raises:
            FileNotFoundError: Se o arquivo não existe
            ValueError: Se há erro ao ler o arquivo
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

        try:
            df = pd.read_excel(file_path, header=header_row)
            return df
        except Exception as e:
            raise ValueError(f"Erro ao ler arquivo Excel: {str(e)}")

    @staticmethod
    def _validate_selected_columns(
        df_pessoas: pd.DataFrame,
        df_secundario: pd.DataFrame,
        selected_columns_pessoas: List[str],
        selected_columns_secundario: List[str]
    ) -> None:
        """
        Valida se as colunas selecionadas existem nos DataFrames.

        Args:
            df_pessoas: DataFrame de pessoas
            df_secundario: DataFrame secundário
            selected_columns_pessoas: Colunas selecionadas de pessoas
            selected_columns_secundario: Colunas selecionadas do arquivo secundário

        Raises:
            ValueError: Se alguma coluna não existe
        """
        # Validar colunas de Pessoas
        missing_pessoas = [
            col for col in selected_columns_pessoas
            if col not in df_pessoas.columns
        ]
        if missing_pessoas:
            raise ValueError(
                f"Colunas não encontradas em Pessoas: {', '.join(missing_pessoas)}"
            )

        # Validar colunas de Secundário
        missing_secundario = [
            col for col in selected_columns_secundario
            if col not in df_secundario.columns
        ]
        if missing_secundario:
            raise ValueError(
                f"Colunas não encontradas em Registros/Níveis: {', '.join(missing_secundario)}"
            )

    @staticmethod
    def _build_select_list(
        selected_secundario: List[str],
        selected_pessoas: List[str]
    ) -> str:
        """
        Constrói a lista de colunas para SELECT SQL.

        Ordem: Todas colunas de Secundário + Colunas de Pessoas não duplicadas

        Args:
            selected_secundario: Colunas selecionadas do arquivo secundário
            selected_pessoas: Colunas selecionadas de pessoas

        Returns:
            String formatada para SQL
        """
        colunas_sql = []

        # 1. Adicionar todas colunas de Secundário
        for col in selected_secundario:
            colunas_sql.append(f'Secundario."{col}"')

        # 2. Adicionar colunas de Pessoas que não estão em Secundário
        for col in selected_pessoas:
            # Verificar se a coluna já está em Secundário (evitar duplicação)
            if col not in selected_secundario:
                colunas_sql.append(f'Pessoas."{col}"')

        return ", ".join(colunas_sql)

    @staticmethod
    def _build_query(
        colunas_sql: str,
        sort_column: Optional[str] = None,
        sort_order: str = "DESC"
    ) -> str:
        """
        Constrói a query SQL completa.

        Args:
            colunas_sql: String com colunas para SELECT
            sort_column: Coluna para ordenação
            sort_order: ASC ou DESC

        Returns:
            Query SQL completa
        """
        query = f"""
        SELECT {colunas_sql}
        FROM Secundario
        LEFT JOIN Pessoas ON Secundario."ID Pessoal" = Pessoas."ID Pessoal"
        """

        # Adicionar ORDER BY se fornecido
        if sort_column:
            # Determinar qual tabela contém a coluna
            # Tentativa 1: Assume que é de Secundário
            query += f'\nORDER BY Secundario."{sort_column}" {sort_order}'

        return query
