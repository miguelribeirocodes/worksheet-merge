"""Gerenciador de configurações para salvar/carregar seleções de checkboxes."""
import json
import os
from pathlib import Path
from typing import List, Dict, Optional


class ConfigManager:
    """Gerencia persistência de configurações de checkboxes em JSON."""

    def __init__(self, config_dir: Optional[str] = None):
        """
        Inicializa o gerenciador de configurações.

        Args:
            config_dir: Diretório para armazenar arquivos de config
                       (padrão: ~/.worksheet-merge/)
        """
        if config_dir is None:
            config_dir = os.path.join(
                Path.home(),
                ".worksheet-merge"
            )

        self.config_dir = config_dir
        self.config_file = os.path.join(config_dir, "configs.json")

        # Criar diretório se não existir
        os.makedirs(config_dir, exist_ok=True)

        # Criar arquivo de config se não existir
        if not os.path.exists(self.config_file):
            self._save_configs({})

    def save_config(
        self,
        config_name: str,
        selected_columns_pessoas: List[str],
        selected_columns_secundario: List[str],
        sort_column: Optional[str] = None,
        sort_order: str = "DESC"
    ) -> bool:
        """
        Salva uma configuração de checkboxes em arquivo JSON.

        Args:
            config_name: Nome da configuração
            selected_columns_pessoas: Colunas selecionadas de pessoas
            selected_columns_secundario: Colunas selecionadas do arquivo secundário
            sort_column: Coluna para ordenação (opcional)
            sort_order: ASC ou DESC

        Returns:
            True se salvo com sucesso, False caso contrário
        """
        try:
            configs = self._load_configs()

            configs[config_name] = {
                "pessoas": selected_columns_pessoas,
                "secundario": selected_columns_secundario,
                "sort_column": sort_column,
                "sort_order": sort_order
            }

            self._save_configs(configs)
            return True

        except Exception as e:
            print(f"Erro ao salvar configuração: {str(e)}")
            return False

    def load_config(self, config_name: str) -> Optional[Dict]:
        """
        Carrega uma configuração salva.

        Args:
            config_name: Nome da configuração

        Returns:
            Dicionário com configuração ou None se não encontrada
        """
        try:
            configs = self._load_configs()

            if config_name in configs:
                return configs[config_name]
            return None

        except Exception as e:
            print(f"Erro ao carregar configuração: {str(e)}")
            return None

    def list_configs(self) -> List[str]:
        """
        Lista todos os nomes de configurações salvas.

        Returns:
            Lista com nomes de configurações
        """
        try:
            configs = self._load_configs()
            return list(configs.keys())
        except Exception as e:
            print(f"Erro ao listar configurações: {str(e)}")
            return []

    def delete_config(self, config_name: str) -> bool:
        """
        Deleta uma configuração salva.

        Args:
            config_name: Nome da configuração

        Returns:
            True se deletado com sucesso, False caso contrário
        """
        try:
            configs = self._load_configs()

            if config_name in configs:
                del configs[config_name]
                self._save_configs(configs)
                return True

            return False

        except Exception as e:
            print(f"Erro ao deletar configuração: {str(e)}")
            return False

    def _load_configs(self) -> Dict:
        """
        Carrega todas as configurações do arquivo JSON.

        Returns:
            Dicionário com todas as configurações
        """
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"Erro ao ler arquivo de config: {str(e)}")
            return {}

    def _save_configs(self, configs: Dict) -> None:
        """
        Salva todas as configurações no arquivo JSON.

        Args:
            configs: Dicionário com todas as configurações
        """
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(configs, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Erro ao salvar arquivo de config: {str(e)}")
            raise
