"""Setup script para criar executáveis com cx_Freeze."""
import sys
from cx_Freeze import setup, Executable

# Detectar plataforma
platform = sys.platform

# Configurações base
base = None
if platform == "win32":
    base = "Win32GUI"  # Para não mostrar console no Windows

# Definir executáveis
executables = [
    Executable(
        script="src/main/pers_access_level_merge.py",
        base=base,
        target_name="Mesclar Níveis de Acesso.exe" if platform == "win32" else "mesclar-niveis-acesso",
        icon=None  # Adicione um ícone aqui se tiver: "icon.ico"
    ),
    Executable(
        script="src/main/pers_access_log_merge.py",
        base=base,
        target_name="Mesclar Registros de Acesso.exe" if platform == "win32" else "mesclar-registros-acesso",
        icon=None  # Adicione um ícone aqui se tiver: "icon.ico"
    ),
]

# Pacotes e módulos a incluir
includes = ["tkinter", "sqlite3", "pandas", "openpyxl"]
excludes = ["tkinter.test", "unittest", "test", "pytest"]

# Opções de build
build_exe_options = {
    "include_files": ["src/utils/"],
    "includes": includes,
    "excludes": excludes,
}

# Configuração geral
setup(
    name="Worksheet Merge",
    version="1.0",
    description="Aplicativo para mesclar planilhas de acesso do sistema ZKBio CVSecurity",
    author="Miguel Ribeiro",
    options={"build_exe": build_exe_options},
    executables=executables
)
