import sys
import os

if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

from pathlib import Path
from pdf2docx import Converter
from spire.doc import Document, FileFormat
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
        try:
            # Conversion autonome sans Microsoft Word
            doc = Document()
            doc.LoadFromFile(str(p))
            doc.SaveToFile(str(out_path), FileFormat.PDF)
            doc.Close()
            afficher_notification(out_path.name)
        except Exception as e:
            ctypes.windll.user32.MessageBoxW(0, f"Erreur DOCX : {e}", "Erreur", 0x10)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        lancer_conversion(sys.argv[1])
