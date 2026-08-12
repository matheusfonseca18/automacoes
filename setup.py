import sys
import os
from cx_Freeze import setup, Executable

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_customtkinter_path():
    import customtkinter
    return os.path.dirname(customtkinter.__file__)

build_exe_options = {
    "packages": [
        "customtkinter",
        "pandas",
        "openpyxl",
        "xlwings",
        "win32com",
        "pythoncom",
        "dotenv"
    ],
    "includes": [
        "win32com.client",
        "pythoncom",
        "pywintypes",
        "customtkinter",

        "almap_africa",
        "almap_africa.almap_africa_refatorado_copy",

        "Indicador_class",
        "Indicador_class.indicador_class_copy_2",

        "shared",
        "shared.backup_utils",
        "shared.logger_utils"
    ],
    "include_files": [
        (
            get_customtkinter_path(),
            "customtkinter"
        ),
    ],
    "include_msvcr": True,

    "optimize": 1,
}

base = None
if sys.platform == "win32":
    base = "gui"

setup(
    name="Automações de Relatórios",
    version="1.0",
    description="Automações de Relatórios - 1.0",
    options={"build_exe": build_exe_options},
    executables=[
        Executable(
            script=os.path.join(BASE_DIR, "app.py"),
            base=base,
            target_name="Automações de Relatórios.exe",
            icon=None
        )
    ],
)
