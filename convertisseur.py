import sys
import os
import json
import struct
import shutil
import winreg
import ctypes
import subprocess
from pathlib import Path

# Fix streams en mode GUI
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

def afficher_notification(nom_fichier):
    ps_cmd = f"""
    [Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime] > $null
    $template = [Windows.UI.Notifications.ToastNotificationManager]::GetTemplateContent([Windows.UI.Notifications.ToastTemplateType]::ToastText02)
    $textNodes = $template.GetElementsByTagName("text")
    $textNodes.Item(0).AppendChild($template.CreateTextNode("Conversion terminee !")) > $null
    $textNodes.Item(1).AppendChild($template.CreateTextNode("{nom_fichier} est pret.")) > $null
    $toast = [Windows.UI.Notifications.ToastNotification]::new($template)
    [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier("Convertisseur Universel").Show($toast)
    """
    try:
        subprocess.Popen(["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", ps_cmd], creationflags=0x08000000)
    except Exception:
        pass

def trouver_soffice():
    chemins = [
        Path(os.environ.get("PROGRAMFILES", "")) / "LibreOffice" / "program" / "soffice.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "LibreOffice" / "program" / "soffice.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "LibreOffice" / "program" / "soffice.exe"
    ]
    for c in chemins:
        if c.exists():
            return str(c)
    return None

def convertir_docx_vers_pdf(docx_path, pdf_path):
    # 1. Tenter Microsoft Word (qualité native)
    word_error = None
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(docx_path), ReadOnly=True)
            doc.SaveAs(str(pdf_path), FileFormat=17) # 17 = wdFormatPDF
            return True
        except Exception as e:
            word_error = e
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()
    except Exception as e:
        word_error = e

    # 2. Fallback: LibreOffice
    soffice = trouver_soffice()
    if soffice:
        try:
            cmd = [soffice, "--headless", "--convert-to", "pdf", str(docx_path), "--outdir", str(docx_path.parent)]
            subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=0x08000000, check=True)
            if pdf_path.exists():
                return True
        except Exception:
            pass

    # 3. Message clair si aucun moteur n'est présent
    if word_error:
        raise RuntimeError(f"Microsoft Word est indisponible ({word_error}). Installez Word ou LibreOffice.")
    else:
        raise RuntimeError("Aucun moteur bureautique détecté. Installez Microsoft Word ou LibreOffice.")

def lancer_conversion(file_path_str):
    p = Path(file_path_str).resolve()
    if not p.exists():
        return

    ext = p.suffix.lower()

    if ext == ".pdf":
        out_path = p.with_suffix(".docx")
        try:
            from pdf2docx import Converter
            cv = Converter(str(p))
            cv.convert(str(out_path))
            cv.close()
            if out_path.exists():
                afficher_notification(out_path.name)
        except Exception as e:
            ctypes.windll.user32.MessageBoxW(0, f"Erreur lors de la conversion PDF -> Word :\n\n{e}", "Erreur Convertisseur", 0x10)

    elif ext == ".docx":
        out_path = p.with_suffix(".pdf")
        try:
            convertir_docx_vers_pdf(p, out_path)
            if out_path.exists():
                afficher_notification(out_path.name)
        except Exception as e:
            ctypes.windll.user32.MessageBoxW(0, f"Erreur lors de la conversion Word -> PDF :\n\n{e}", "Erreur Convertisseur", 0x10)

def mode_native_host():
    def read_msg():
        raw_len = sys.stdin.buffer.read(4)
        if len(raw_len) == 0:
            sys.exit(0)
        msg_len = struct.unpack('@I', raw_len)[0]
        data = sys.stdin.buffer.read(msg_len).decode('utf-8')
        return json.loads(data)

    def send_msg(obj):
        encoded = json.dumps(obj).encode('utf-8')
        sys.stdout.buffer.write(struct.pack('@I', len(encoded)))
        sys.stdout.buffer.write(encoded)
        sys.stdout.buffer.flush()

    try:
        msg = read_msg()
        file_path = msg.get("filePath")
        if file_path and os.path.exists(file_path):
            exe_path = sys.executable
            subprocess.Popen([exe_path, file_path], creationflags=0x08000000)
            send_msg({"status": "success", "file": file_path})
        else:
            send_msg({"status": "error", "message": "Fichier introuvable"})
    except Exception as e:
        try:
            send_msg({"status": "error", "message": str(e)})
        except Exception:
            pass

def delete_key_tree(root, subkey):
    try:
        key = winreg.OpenKey(root, subkey, 0, winreg.KEY_ALL_ACCESS)
        while True:
            try:
                child = winreg.EnumKey(key, 0)
                delete_key_tree(key, child)
            except OSError:
                break
        winreg.CloseKey(key)
        winreg.DeleteKey(root, subkey)
    except OSError:
        pass

def mode_desinstallation():
    MB_YESNO = 0x04
    MB_ICONQUESTION = 0x20
    IDYES = 6
    ret = ctypes.windll.user32.MessageBoxW(
        0, 
        "Voulez-vous vraiment désinstaller Convertisseur Universel ?", 
        "Désinstallation de Convertisseur Universel", 
        MB_YESNO | MB_ICONQUESTION
    )
    if ret != IDYES:
        return

    reg_keys = [
        r"Software\Classes\SystemFileAssociations\.docx\shell\ConvertToPdf",
        r"Software\Classes\SystemFileAssociations\.pdf\shell\ConvertToDocx",
        r"Software\Google\Chrome\NativeMessagingHosts\com.convertisseur.universel",
        r"Software\Microsoft\Edge\NativeMessagingHosts\com.convertisseur.universel",
        r"Software\Microsoft\Windows\CurrentVersion\Uninstall\ConvertisseurUniversel"
    ]
    for k in reg_keys:
        delete_key_tree(winreg.HKEY_CURRENT_USER, k)

    try:
        ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
    except Exception:
        pass

    local_app_data = os.environ.get("LOCALAPPDATA", "")
    target_dir = os.path.join(local_app_data, "Programs", "ConvertisseurUniversel")
    
    cmd = f'timeout /t 2 /nobreak > nul & rmdir /s /q "{target_dir}"'
    subprocess.Popen(["cmd.exe", "/c", cmd], creationflags=0x08000000)

    ctypes.windll.user32.MessageBoxW(0, "Convertisseur Universel a été désinstallé avec succès.", "Désinstallation terminée", 0x40)

def mode_installation():
    try:
        local_app_data = os.environ.get("LOCALAPPDATA", "")
        if not local_app_data:
            local_app_data = str(Path.home() / "AppData" / "Local")

        target_dir = Path(local_app_data) / "Programs" / "ConvertisseurUniversel"
        target_dir.mkdir(parents=True, exist_ok=True)
        ext_target_dir = target_dir / "extension"
        ext_target_dir.mkdir(parents=True, exist_ok=True)

        current_exe = Path(sys.executable).resolve()
        target_exe = target_dir / "convertisseur.exe"

        # Copier l'exécutable dans le dossier d'installation si différent
        if current_exe != target_exe:
            shutil.copy2(str(current_exe), str(target_exe))

        # Copier l'extension Chrome si présente dans les ressources ou à côté
        base_dir = Path(getattr(sys, '_MEIPASS', Path(__file__).parent)).resolve()
        src_ext = base_dir / "extension"
        if not src_ext.exists():
            src_ext = current_exe.parent / "extension"
        if src_ext.exists():
            for item in src_ext.iterdir():
                if item.is_file():
                    shutil.copy2(str(item), str(ext_target_dir / item.name))

        # Générer host_manifest.json
        manifest_data = {
            "name": "com.convertisseur.universel",
            "description": "Host natif Convertisseur Universel",
            "path": str(target_exe),
            "type": "stdio",
            "allowed_origins": [
                "chrome-extension://hmjckaeohligmgpndakbfcjicjmcnljg/",
                "chrome-extension://*"
            ]
        }
        manifest_path = target_dir / "host_manifest.json"
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest_data, f, indent=2)

        # Enregistrement registre Menus Contextuels Windows
        target_exe_str = str(target_exe)

        # DOCX -> PDF
        k_docx = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\SystemFileAssociations\.docx\shell\ConvertToPdf")
        winreg.SetValueEx(k_docx, "", 0, winreg.REG_SZ, "Convertir en PDF")
        winreg.SetValueEx(k_docx, "Icon", 0, winreg.REG_SZ, f'"{target_exe_str}",0')
        k_docx_cmd = winreg.CreateKey(k_docx, "command")
        winreg.SetValueEx(k_docx_cmd, "", 0, winreg.REG_SZ, f'"{target_exe_str}" "%1"')
        winreg.CloseKey(k_docx_cmd)
        winreg.CloseKey(k_docx)

        # PDF -> DOCX
        k_pdf = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Classes\SystemFileAssociations\.pdf\shell\ConvertToDocx")
        winreg.SetValueEx(k_pdf, "", 0, winreg.REG_SZ, "Convertir en Word (DOCX)")
        winreg.SetValueEx(k_pdf, "Icon", 0, winreg.REG_SZ, f'"{target_exe_str}",0')
        k_pdf_cmd = winreg.CreateKey(k_pdf, "command")
        winreg.SetValueEx(k_pdf_cmd, "", 0, winreg.REG_SZ, f'"{target_exe_str}" "%1"')
        winreg.CloseKey(k_pdf_cmd)
        winreg.CloseKey(k_pdf)

        # Chrome & Edge Native Messaging
        k_ch = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Google\Chrome\NativeMessagingHosts\com.convertisseur.universel")
        winreg.SetValueEx(k_ch, "", 0, winreg.REG_SZ, str(manifest_path))
        winreg.CloseKey(k_ch)

        k_ed = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Edge\NativeMessagingHosts\com.convertisseur.universel")
        winreg.SetValueEx(k_ed, "", 0, winreg.REG_SZ, str(manifest_path))
        winreg.CloseKey(k_ed)

        # Désinstallation Windows
        k_un = winreg.CreateKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall\ConvertisseurUniversel")
        winreg.SetValueEx(k_un, "DisplayName", 0, winreg.REG_SZ, "Convertisseur Universel (PDF / Word)")
        winreg.SetValueEx(k_un, "DisplayIcon", 0, winreg.REG_SZ, f'"{target_exe_str}",0')
        winreg.SetValueEx(k_un, "UninstallString", 0, winreg.REG_SZ, f'"{target_exe_str}" --uninstall')
        winreg.SetValueEx(k_un, "Publisher", 0, winreg.REG_SZ, "Ethan")
        winreg.SetValueEx(k_un, "DisplayVersion", 0, winreg.REG_SZ, "1.0.0")
        winreg.SetValueEx(k_un, "InstallLocation", 0, winreg.REG_SZ, str(target_dir))
        winreg.CloseKey(k_un)

        try:
            ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, None, None)
        except Exception:
            pass

        msg = (
            "Convertisseur Universel a été installé avec succès !\n\n"
            "✨ Fonctionnalités activées :\n"
            "• Clic droit sur un PDF -> 'Convertir en Word (DOCX)'\n"
            "• Clic droit sur un DOCX -> 'Convertir en PDF'\n"
            "• Menu Chrome / Edge disponible via l'extension incluse.\n\n"
            f"Dossier du programme :\n{target_dir}"
        )
        ctypes.windll.user32.MessageBoxW(0, msg, "Installation terminée", 0x40)

    except Exception as e:
        ctypes.windll.user32.MessageBoxW(0, f"Erreur lors de l'installation : {e}", "Erreur d'installation", 0x10)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()
        if arg == "--install" or arg == "/install":
            mode_installation()
        elif arg == "--uninstall" or arg == "/uninstall":
            mode_desinstallation()
        elif arg == "--host" or arg.startswith("chrome-extension://"):
            mode_native_host()
        else:
            lancer_conversion(sys.argv[1])
    else:
        # Si exécuté sans argument (double-clic direct sur Setup / convertisseur)
        mode_installation()
