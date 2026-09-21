import sys
import os
import winreg
import ctypes
import subprocess

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

def uninstall():
    MB_YESNO = 0x04
    MB_ICONQUESTION = 0x20
    IDYES = 6
    ret = ctypes.windll.user32.MessageBoxW(
        0, 
        "Voulez-vous vraiment désinstaller Convertisseur Universel ?", 
        "Désinstallation", 
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

    ctypes.windll.user32.MessageBoxW(
        0, 
        "Convertisseur Universel a été désinstallé avec succès.", 
        "Désinstallation terminée", 
        0x40
    )

if __name__ == "__main__":
    uninstall()
