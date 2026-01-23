# 📊 Worksheet Merge

Uma aplicação desktop para mesclar e processar dados de acesso e pessoal do sistema ZKBio CVSecurity.

## 🎯 Funcionalidades

Aplicativo desktop unificado para mesclar dados de acesso do ZKBio CVSecurity:

### **Aplicativo com Interface de Checkboxes**
- Interface intuitiva tipo ZKBio CVSecurity
- Seleção dinâmica de colunas via checkboxes
- Descoberta automática de colunas disponíveis
- Organização de colunas por categorias (Informações Básicas, Documentação, Datas, etc)
- Suporte a colunas customizadas não previstas
- Salvamento e carregamento de configurações reutilizáveis
- Merge parametrizado com validação automática
- Suporte completo a LEFT JOIN entre planilhas
- Ordenação flexível dos resultados

**Casos de Uso:**
- Mesclar Pessoas + Níveis de Acesso
- Mesclar Pessoas + Registros de Acesso
- Mesclar Pessoas + Qualquer outro arquivo do ZKBio
- Customização total de colunas no resultado final

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
# Iniciar o aplicativo
python src/main/app.py
```

## 📦 Estrutura do Projeto

```
worksheet-merge/
├── src/
│   ├── main/
│   │   ├── __init__.py
│   │   └── app.py                      # Aplicativo principal unificado
│   └── utils/
│       ├── __init__.py
│       ├── validators.py               # Validações de entrada e colunas
│       ├── ui_helpers.py               # Componentes Tkinter (CategoryFrame, ScrollableFrame)
│       ├── merge_engine.py             # Engine de merge parametrizado
│       ├── column_loader.py            # Descoberta e categorização dinâmica de colunas
│       └── config_manager.py           # Persistência de configurações
├── testes/                             # Dados de teste (exemplos do ZKBio)
├── README.md                           # Este arquivo
├── requirements.txt                    # Dependências Python
└── .gitignore                          # Arquivos ignorados pelo Git
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

## ⚙️ Usando o Novo Aplicativo com Checkboxes

### Passo a Passo:

1. **Inicie o Aplicativo:**
   ```bash
   python src/main/app.py
   ```

2. **Selecione os Arquivos:**
   - Clique em "Selecionar" para escolher o arquivo de **Pessoas** (.xls ou .xlsx)
   - Clique em "Selecionar" para escolher o arquivo **Secundário** (Registros ou Níveis de Acesso)
   - As colunas disponíveis serão carregadas automaticamente

3. **Escolha as Colunas:**
   - Organize suas seleções usando as duas abas (Pessoas e Registros/Níveis)
   - As colunas são organizadas por categorias (Informações Básicas, Documentação, etc)
   - Marque/desmarque os checkboxes conforme desejado
   - **"ID Pessoal" é obrigatório** em ambas as abas (sempre pré-selecionado)
   - Opcionalmente, adicione colunas customizadas não previstas

4. **Configure Opções:**
   - Escolha a coluna para ordenação (ex: Horário, Nome do Nível)
   - Selecione Crescente ou Decrescente

5. **Salve ou Carregue Configurações:**
   - Digite um nome e clique "Salvar" para guardar suas seleções
   - Use o dropdown para "Carregar" uma configuração salva anteriormente
   - Clique "Excluir" para remover uma configuração

6. **Realize o Merge:**
   - Clique no botão "MESCLAR"
   - Escolha o local para salvar o arquivo resultado
   - O sistema criará um novo arquivo Excel com as colunas selecionadas

### Notas:
- As configurações são salvas em `~/.worksheet-merge/configs.json`
- O sistema valida automaticamente se as colunas selecionadas existem nas planilhas
- O merge utiliza LEFT JOIN, preservando todos os registros da planilha secundária


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
