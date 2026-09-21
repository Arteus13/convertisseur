import sys
import os

if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

from pathlib import Path
from pdf2docx import Converter
import subprocess
import ctypes

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
        Path(os.environ.get("PROGRAMFILES", "")) / "LibreOffice/program/soffice.exe",
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "LibreOffice/program/soffice.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs/LibreOffice/program/soffice.exe"
    ]
    for c in chemins:
        if c.exists():
            return str(c)
    return None

def lancer_conversion(file_path_str):
    p = Path(file_path_str)
    if not p.exists():
        return

    ext = p.suffix.lower()

    if ext == ".pdf":
        out_path = p.with_suffix(".docx")
        try:
            cv = Converter(str(p))
            cv.convert(str(out_path))
            cv.close()
            afficher_notification(out_path.name)
        except Exception as e:
            ctypes.windll.user32.MessageBoxW(0, f"Erreur PDF : {e}", "Erreur", 0x10)

    elif ext == ".docx":
        out_path = p.with_suffix(".pdf")
        soffice = trouver_soffice()
        
        if soffice:
            try:
                # Exécution 100% invisible en arrière-plan
                cmd = [soffice, "--headless", "--convert-to", "pdf", str(p), "--outdir", str(p.parent)]
                subprocess.run(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, creationflags=0x08000000, check=True)
                afficher_notification(out_path.name)
            except Exception as e:
                ctypes.windll.user32.MessageBoxW(0, f"Erreur DOCX : {e}", "Erreur", 0x10)
        else:
            # Si ni LibreOffice ni Word ne sont trouvés
            ctypes.windll.user32.MessageBoxW(
                0, 
                "Moteur bureautique introuvable pour convertir le DOCX en PDF.\nInstallez LibreOffice ou Microsoft Word.", 
                "Moteur manquant", 
                0x30
            )

if __name__ == "__main__":
    if len(sys.argv) > 1:
        lancer_conversion(sys.argv[1])
