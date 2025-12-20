import PyInstaller.__main__
import customtkinter
import tkinterDnD  # python-tkdnd
import os
import sys

# Get module paths for data inclusion
ctk_path = os.path.dirname(customtkinter.__file__)
tkdnd_path = os.path.dirname(tkinterDnD.__file__)

print(f"CustomTkinter path: {ctk_path}")
print(f"TkinterDnD path: {tkdnd_path}")

# HWPCONV specific paths
project_root = os.path.dirname(os.path.abspath(__file__))
src_path = os.path.join(project_root, 'src')

# PyInstaller arguments
args = [
    'run_gui.py',                        # Main script
    '--name=HwpConverterPro',            # Executable name
    '--noconsole',                       # No console window
    '--onefile',                         # Single executable file
    f'--paths={src_path}',               # Add src to search path
    f'--add-data={ctk_path}{os.pathsep}customtkinter', # Include CustomTkinter themes
    f'--add-data={tkdnd_path}{os.pathsep}tkinterDnD',  # Include python-tkdnd
    '--clean',                           # Clean cache
    '--log-level=WARN',
]

print("Building EXE with arguments:", args)

PyInstaller.__main__.run(args)
