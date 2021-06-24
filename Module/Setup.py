import sys
from cx_Freeze import setup, Executable

build_exe_options={
    "packages":["os","PySimpleGui","selenium","re","time","openpyxl","win32com","json","PyQt5","random","tempfile"],
    "include_files":[r"Module\__xbase__"]
    }

base=None
if sys.platform=="win32":
    base="Win32GUI"

setup(name="CadKmt",version="0.1",description="Aplicativo para cadastro de itens no site do Klassmatt de forma automatica.",
      options={"build_exe":build_exe_options},executables=[Executable("CadKlassmatt.py",base=base)])


#Linha de comando CMD:
# python Setup.py bdist_msi
