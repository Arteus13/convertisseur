import sys
import os
import shutil
import json
import winreg
import ctypes
from pathlib import Path

def get_base_dir():
    if hasattr(sys, '_MEIPASS'):
        return Path(sys._MEIPASS)
    return Path(__file__).parent.resolve()

def install():
    try:
        base_dir = get_base_dir()
        local_app_data = os.environ.get("LOCALAPPDATA", "")
        if not local_app_data:
            local_app_data = str(Path.home() / "AppData" / "Local")

        target_dir = Path(local_app_data) / "Programs" / "ConvertisseurUniversel"
        target_dir.mkdir(parents=True, exist_ok=True)
        ext_target_dir = target_dir / "extension"
        ext_target_dir.mkdir(parents=True, exist_ok=True)

        # 1. Copier les fichiers exécutables
        for filename in ["convertisseur.exe", "Uninstaller.exe", "native_host.exe"]:
            src = base_dir / filename
            if src.exists():
                shutil.copy2(str(src), str(target_dir / filename))

        # 2. Copier l'extension Chrome
        src_ext = base_dir / "extension"
        if src_ext.exists():
            for item in src_ext.iterdir():
                if item.is_file():
                    shutil.copy2(str(item), str(ext_target_dir / item.name))

        # 3. Générer host_manifest.json avec les chemins absolus propres à cette machine
        native_host_path = target_dir / "native_host.exe"
        manifest_data = {
            "name": "com.convertisseur.universel",
            "description": "Host natif Convertisseur Universel",
            "path": str(native_host_path),
            "type": "stdio",
            "allowed_origins": [
                "chrome-extension://hmjckaeohligmgpndakbfcjicjmcnljg/",
                "chrome-extension://*"
            ]
        }
        manifest_path = target_dir / "host_manifest.json"
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest_data, f, indent=2)

        # 4. Enregistrer dans le Registre Windows
        convertisseur_exe = str(target_dir / "convertisseur.exe")
        uninstaller_exe = str(target_dir / "Uninstaller.exe")

        # Menu contextuel DOCX -> PDF
        key_docx = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\SystemFileAssociations\.docx\shell\ConvertToPdf")
        winreg.SetValueEx(key_docx, "", 0, winreg.REG_SZ, "Convertir en PDF")
        winreg.SetValueEx(key_docx, "Icon", 0, winreg.REG_SZ, f'"{convertisseur_exe}",0')
        key_docx_cmd = winreg.CreateKey(key_docx, "command")
        winreg.SetValueEx(key_docx_cmd, "", 0, winreg.REG_SZ, f'"{convertisseur_exe}" "%1"')
        winreg.CloseKey(key_docx_cmd)
        winreg.CloseKey(key_docx)

        # Menu contextuel PDF -> DOCX
        key_pdf = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\SystemFileAssociations\.pdf\shell\ConvertToDocx")
        winreg.SetValueEx(key_pdf, "", 0, winreg.REG_SZ, "Convertir en Word (DOCX)")
        winreg.SetValueEx(key_pdf, "Icon", 0, winreg.REG_SZ, f'"{convertisseur_exe}",0')
        key_pdf_cmd = winreg.CreateKey(key_pdf, "command")
        winreg.SetValueEx(key_pdf_cmd, "", 0, winreg.REG_SZ, f'"{convertisseur_exe}" "%1"')
        winreg.CloseKey(key_pdf_cmd)
        winreg.CloseKey(key_pdf)

        # Chrome Native Messaging Host
        key_chrome = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\NativeMessagingHosts\com.convertisseur.universel")
        winreg.SetValueEx(key_chrome, "", 0, winreg.REG_SZ, str(manifest_path))
        winreg.CloseKey(key_chrome)

        # Edge Native Messaging Host
        key_edge = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Edge\NativeMessagingHosts\com.convertisseur.universel")
        winreg.SetValueEx(key_edge, "", 0, winreg.REG_SZ, str(manifest_path))
        winreg.CloseKey(key_edge)

        # Entrée Ajout / Suppression de programmes Windows
        key_uninstall = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall\ConvertisseurUniversel")
        winreg.SetValueEx(key_uninstall, "DisplayName", 0, winreg.REG_SZ, "Convertisseur Universel (PDF / Word)")
        winreg.SetValueEx(key_uninstall, "DisplayIcon", 0, winreg.REG_SZ, f'"{convertisseur_exe}",0')
        winreg.SetValueEx(key_uninstall, "UninstallString", 0, winreg.REG_SZ, f'"{uninstaller_exe}"')
        winreg.SetValueEx(key_uninstall, "Publisher", 0, winreg.REG_SZ, "Ethan")
        winreg.SetValueEx(key_uninstall, "DisplayVersion", 0, winreg.REG_SZ, "1.0.0")
        winreg.SetValueEx(key_uninstall, "InstallLocation", 0, winreg.REG_SZ, str(target_dir))
        winreg.CloseKey(key_uninstall)

        try:
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
        except Exception:
            pass

        msg = (
            "Convertisseur Universel a été installé avec succès !\n\n"
            "✨ Fonctionnalités activées :\n"
            "• Clic droit sur un fichier .pdf -> 'Convertir en Word (DOCX)'\n"
            "• Clic droit sur un fichier .docx -> 'Convertir en PDF'\n"
            "• Extension de navigateur intégrée dans le dossier du programme.\n\n"
            f"Dossier de l'extension Chrome :\n{ext_target_dir}"
        )
        ctypes.windll.user32.MessageBoxW(0, msg, "Installation terminée", 0x40)

    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f"Erreur lors de l'installation : {e}", "Erreur d'installation", 0x10)

if __name__ == "__main__":
    install()
