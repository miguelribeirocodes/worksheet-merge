# 📊 Worksheet Merge

Uma aplicação desktop para mesclar e processar dados de acesso e pessoal do sistema ZKBio CVSecurity.

## 🎯 Funcionalidades

O projeto oferece dois utilitários complementares:

### 1. **Mesclar Níveis de Acesso**
- Combina dados de pessoal com níveis de acesso
- Realiza LEFT JOIN baseado em ID Pessoal
- Exporta resultado em arquivo Excel
- Ordena por Nome do Nível em ordem crescente

**Colunas esperadas:**
- **Pessoas**: Nome, Sobrenome, Nome do Departamento, Número do Documento, Nome do Cargo, ID Pessoal
- **Níveis de Acesso**: Nome do Nível, ID Pessoal

### 2. **Mesclar Registros de Acesso**
- Combina dados de pessoal com logs de acesso
- Enriquece registros com informações pessoais
- Exporta resultado em arquivo Excel ordenado por timestamp (mais recente primeiro)

**Colunas esperadas:**
- **Pessoas**: Nome do Cargo, Tipo de Documento, Número do Documento, ID Pessoal, Observações
- **Registros**: Horário, Nome da Área, Nome do Dispositivo, Descrição do Evento, ID Pessoal, Nome, Sobrenome, Nome do Departamento

## 🚀 Como Usar

### Método 1: Executável (Recomendado para Usuários Finais)

1. Baixe o arquivo executável da [página de releases](https://github.com/miguelribeirocodes/worksheet-merge/releases)
2. Clique duas vezes para executar o programa
3. Selecione as duas planilhas (pessoas e acesso/registros)
4. Clique em "Mesclar planilhas"
5. Escolha onde salvar o arquivo mesclado

### Método 2: Código Fonte (Para Desenvolvedores)

#### Pré-requisitos
- Python 3.7+
- pip (gerenciador de pacotes Python)

#### Instalação

```bash
# Clone o repositório
git clone https://github.com/miguelribeirocodes/worksheet-merge.git
cd worksheet-merge

# Instale as dependências
pip install -r requirements.txt
```

#### Execução

```bash
# Mesclar Níveis de Acesso
python src/main/pers_access_level_merge.py

# Mesclar Registros de Acesso
python src/main/pers_access_log_merge.py
```

## 📦 Estrutura do Projeto

```
worksheet-merge/
├── src/
│   ├── main/                           # Scripts principais
│   │   ├── __init__.py
│   │   ├── pers_access_level_merge.py  # Mesclar níveis de acesso
│   │   └── pers_access_log_merge.py    # Mesclar registros de acesso
│   └── utils/                          # Módulo de utilitários
│       ├── __init__.py
│       ├── validators.py               # Funções de validação
│       └── ui_helpers.py               # Funções auxiliares de UI
├── setup.py                            # Configuração para build
├── requirements.txt                    # Dependências Python
├── .gitignore                          # Arquivos ignorados pelo Git
└── README.md                           # Este arquivo
```

## 🔧 Compilando um Executável

Se você deseja criar seu próprio executável:

```bash
# Instale cx_Freeze
pip install cx_Freeze

# Compile (gera pasta 'dist' com o executável)
python setup.py build

# O executável estará em: dist/
```

## 📋 Dependências

- **pandas**: Manipulação de dados e I/O de Excel
- **openpyxl**: Suporte avançado para arquivos Excel
- **tkinter**: Interface gráfica (já vem com Python)
- **sqlite3**: Banco de dados para merges (já vem com Python)

## ⚙️ Configuração de Colunas

Para adicionar ou remover colunas do processo de mesclagem, edite as constantes no topo de cada script:

```python
# Exemplo: src/main/pers_access_level_merge.py

# Colunas obrigatórias de cada planilha
COLUNAS_PESSOAS = ["Nome", "Sobrenome", ..., "ID Pessoal"]
COLUNAS_ACESSOS = ["Nome do Nível", "ID Pessoal"]

# Colunas para seleção na query SQL
COLUNAS_SELECT_ACESSOS = [
    "Nivel.\"Nome do Nível\"",
    "Nivel.\"ID Pessoal\"",
    "Pessoas.\"Nome\"",
    # ... adicione mais aqui
]
```

As dicas de UI serão automaticamente atualizadas com base nessas configurações.

## ✅ Validações Automáticas

O programa valida automaticamente:
- ✓ Existência dos arquivos selecionados
- ✓ Extensão dos arquivos (deve ser .xlsx ou .xls)
- ✓ Presença de todas as colunas obrigatórias
- ✓ Formato dos dados Excel

Caso algo esteja incorreto, mensagens de erro específicas ajudam na correção.

## 🔒 Sobre os Dados

- **Segurança**: O programa usa banco de dados SQLite temporário que é automaticamente deletado após o processamento
- **Privacidade**: Nenhum dado é armazenado ou enviado para servidor externo
- **Integridade**: Usa LEFT JOIN para preservar todos os registros de pessoas/acessos

## 🐛 Troubleshooting

### "Arquivo não encontrado"
- Verifique se o arquivo existe e se o caminho é correto
- Tente usar caminhos sem espaços ou caracteres especiais

### "Colunas faltando"
- Verifique os nomes exatos das colunas (sensível a maiúsculas/minúsculas)
- Veja a lista de colunas esperadas no painel de dicas do programa

### "Erro ao ler arquivo Excel"
- Certifique-se de que o arquivo não está corrompido
- Tente abrir no Excel ou LibreOffice para verificar

## 📝 Licença

Este projeto é de código aberto e está disponível sob a licença MIT.

## 👤 Autor

**Miguel Ribeiro Codes**
- GitHub: [@miguelribeirocodes](https://github.com/miguelribeirocodes)

## 🤝 Contribuições

Contribuições são bem-vindas! Por favor:
1. Faça um Fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/AmazingFeature`)
3. Commit suas mudanças (`git commit -m 'Add some AmazingFeature'`)
4. Push para a branch (`git push origin feature/AmazingFeature`)
5. Abra um Pull Request

## 📞 Suporte

Para reportar bugs ou sugerir melhorias, abra uma [issue](https://github.com/miguelribeirocodes/worksheet-merge/issues).

---

**Desenvolvido com ❤️ para facilitar a gestão de dados de acesso**
