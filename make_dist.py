import os
import sys
import shutil
import subprocess
from pathlib import Path

def run(cmd):
    print(f"Running: {cmd}")
    subprocess.run(cmd, shell=True, check=True)

def build():
    root = Path(__file__).parent.resolve()
    os.chdir(root)

    print("=== 1. Compilation de convertisseur.exe ===")
    run('py -m PyInstaller --noconsole --onefile --hidden-import="win32com" --hidden-import="pythoncom" --hidden-import="pdf2docx" --name "convertisseur" convertisseur.py')

    print("=== 2. Compilation de native_host.exe ===")
    run('py -m PyInstaller --console --onefile --name "native_host" native_host.py')

    print("=== 3. Compilation de Uninstaller.exe ===")
    run('py -m PyInstaller --noconsole --onefile --name "Uninstaller" uninstaller.py')

    print("=== 4. Compilation de Setup.exe (Package d'installation complet) ===")
    run('py -m PyInstaller --noconsole --onefile --add-binary "dist/convertisseur.exe;." --add-binary "dist/Uninstaller.exe;." --add-binary "dist/native_host.exe;." --add-data "extension;extension" --name "Setup" --distpath "distrib" setup_installer.py')

    # Copie de l'uninstaller dans distrib
    shutil.copy2(root / "dist" / "Uninstaller.exe", root / "distrib" / "Uninstaller.exe")

    print("=== 5. Nettoyage des artefacts temporaires ===")
    for folder in ["build", "dist"]:
        shutil.rmtree(folder, ignore_errors=True)
    for spec in root.glob("*.spec"):
        try:
            spec.unlink()
        except Exception:
            pass

    print("\n[OK] Build termine avec succes dans le dossier 'distrib/' :")
    print(f" - {(root / 'distrib' / 'Setup.exe').resolve()}")
    print(f" - {(root / 'distrib' / 'Uninstaller.exe').resolve()}")

if __name__ == "__main__":
    build()
