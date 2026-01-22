"""
Script para mesclar planilhas de Pessoas com Níveis de Acesso.
Exporta dados do sistema ZKBio CVSecurity.
"""
import sys
import pandas as pd
import tkinter
from tkinter import filedialog
from tkinter.filedialog import asksaveasfile
from tkinter import messagebox
import sqlite3
import tempfile
import os

# Adicionar o caminho do módulo utils ao path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))
from utils import validar_entrada, validar_colunas, set_path, gerar_texto_dicas_dinamico

# Configurações de colunas - Edite aqui para adicionar/remover colunas
COLUNAS_PESSOAS = ["Nome", "Sobrenome", "Nome do Departamento", "Número do Documento", "Nome do Cargo", "ID Pessoal"]
COLUNAS_ACESSOS = ["Nome do Nível", "ID Pessoal"]

# Colunas para seleção na query (tabela_coluna)
COLUNAS_SELECT_ACESSOS = [
    "Nivel.\"Nome do Nível\"",
    "Nivel.\"ID Pessoal\"",
    "Pessoas.\"Nome\"",
    "Pessoas.\"Sobrenome\"",
    "Pessoas.\"Nome do Departamento\"",
    "Pessoas.\"Número do Documento\"",
    "Pessoas.\"Nome do Cargo\""
]


def mesclar_planilhas(path_pessoas, path_acessos):
    """Mescla planilhas de pessoas com níveis de acesso."""
    # Validar entrada
    if not validar_entrada(path_pessoas, path_acessos, "níveis de acesso"):
        return

    try:
        # Carregar as planilhas do Excel
        planilha_pessoas = pd.read_excel(path_pessoas, header=1)
        planilha_nivel_acesso = pd.read_excel(path_acessos, header=1)

        # Validar colunas obrigatórias
        if not validar_colunas(planilha_pessoas, planilha_nivel_acesso, COLUNAS_PESSOAS, COLUNAS_ACESSOS, "níveis de acesso"):
            return

        # Criar banco de dados temporário
        temp_db = tempfile.NamedTemporaryFile(delete=False, suffix='.db')
        db_path = temp_db.name
        temp_db.close()

        try:
            # Conectar ao banco de dados SQLite
            conn = sqlite3.connect(db_path)

            # Salvar as planilhas como tabelas no banco de dados
            planilha_pessoas.to_sql('Pessoas', conn, if_exists='replace', index=False)
            planilha_nivel_acesso.to_sql('Nivel', conn, if_exists='replace', index=False)

            # Construir query dinamicamente a partir das colunas definidas
            colunas_sql = ", ".join(COLUNAS_SELECT_ACESSOS)
            query = f"""
            SELECT {colunas_sql}
            FROM Nivel
            left join Pessoas on Nivel."ID Pessoal" = Pessoas."ID Pessoal"
            ORDER BY Nivel."Nome do Nível" asc
            """

            df = pd.read_sql_query(query, conn)
            conn.close()

            # Solicitar local para salvar arquivo
            exportFile = filedialog.asksaveasfile(initialfile="planilha_mesclada", title="Salvar arquivo", defaultextension=".xlsx")

            if exportFile is None:
                messagebox.showwarning("Cancelado", "Operação cancelada pelo usuário.")
                return

            df.to_excel(exportFile.name, index=False)
            messagebox.showinfo("Sucesso", f"Planilhas mescladas e salvas em:\n{exportFile.name}")

        finally:
            # Limpar arquivo temporário
            if os.path.exists(db_path):
                os.remove(db_path)

    except pd.errors.ParserError:
        messagebox.showerror("Erro", "Erro ao ler o arquivo Excel. Verifique se o arquivo está corrompido.")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro inesperado:\n{str(e)}")


def main():
    """Função principal que cria a interface gráfica."""
    root = tkinter.Tk()
    root.title("Planilhas de Nível de Acesso")
    root.geometry("500x700")
    root.resizable(False, False)

    # Título
    lbl_title = tkinter.Label(root, text="Selecione as planilhas e clique em mesclar.", font=("Arial", 12, "bold"))
    lbl_title.pack(pady=10)

    # Frame para seleção de arquivos
    frame_files = tkinter.Frame(root)
    frame_files.pack(pady=10, padx=10, fill="both")

    lbl_pessoas = tkinter.Label(frame_files, text="Planilha de Pessoas:")
    lbl_pessoas.pack(anchor="w")
    txt_path_pessoas = tkinter.Entry(frame_files, width=50)
    txt_path_pessoas.pack(pady=5)
    btn_get_path_pessoas = tkinter.Button(frame_files, text="Selecione a planilha de pessoas", command=lambda: set_path(txt_path_pessoas))
    btn_get_path_pessoas.pack(pady=5)

    lbl_acessos = tkinter.Label(frame_files, text="Planilha de Níveis de Acesso:")
    lbl_acessos.pack(anchor="w", pady=(15, 0))
    txt_path_acessos = tkinter.Entry(frame_files, width=50)
    txt_path_acessos.pack(pady=5)
    btn_get_path_acessos = tkinter.Button(frame_files, text="Selecione a planilha de níveis de acessos", command=lambda: set_path(txt_path_acessos))
    btn_get_path_acessos.pack(pady=5)

    # Botão mesclar
    btn_mesclar = tkinter.Button(root, text="Mesclar planilhas", command=lambda: mesclar_planilhas(txt_path_pessoas.get(), txt_path_acessos.get()), bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), padx=20, pady=10)
    btn_mesclar.pack(pady=20)

    # Dicas (atualizado automaticamente conforme as colunas mudam)
    texto_dicas = gerar_texto_dicas_dinamico(COLUNAS_PESSOAS, COLUNAS_ACESSOS, "níveis de acesso")
    lbl_dica = tkinter.Label(
        root,
        text=texto_dicas,
        wraplength=450,
        justify="left",
        bg="#f0f0f0",
        relief="sunken",
        padx=10,
        pady=10
    )
    lbl_dica.pack(pady=10, padx=10, fill="both", expand=True)

    root.mainloop()


if __name__ == "__main__":
    main()
