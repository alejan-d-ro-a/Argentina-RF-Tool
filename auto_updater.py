import os
import sys
import time
import requests
import tempfile
import subprocess
import tkinter as tk
from tkinter import messagebox

# =================================================================
# CONFIGURACIÓN DEL UPDATER
# =================================================================
CURRENT_VERSION = "v1.0.3" 
GITHUB_REPO = "alejan-d-ro-a/Argentina-RF-Tool"
EXE_NAME = "Argentina.RF.Tool.exe" 

def check_for_updates():
    """
    Verifica si hay actualizaciones en GitHub o si el programa acaba de actualizarse.
    """
    # 1. Verificar si acabamos de ser reiniciados por una actualización
    if "--just-updated" in sys.argv:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showinfo(
            "Actualización Completada", 
            f"La herramienta se ha actualizado silenciosamente a la versión {CURRENT_VERSION}.\n\nDisfruta de las mejoras."
        )
        root.destroy()
        return

    # 2. Buscar actualizaciones en segundo plano
    try:
        api_url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        response = requests.get(api_url, timeout=5)
        
        if response.status_code == 200:
            data = response.json()
            latest_version = data.get("tag_name", "")
            
            if latest_version and latest_version > CURRENT_VERSION:
                assets = data.get("assets", [])
                download_url = None
                for asset in assets:
                    if asset["name"] == EXE_NAME:
                        download_url = asset["browser_download_url"]
                        break
                
                if download_url:
                    _perform_silent_update(download_url)
    except Exception:
        pass

def _perform_silent_update(download_url):
    try:
        current_exe_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
        
        temp_dir = tempfile.gettempdir()
        new_exe_path = os.path.join(temp_dir, "new_" + EXE_NAME)
        
        response = requests.get(download_url, stream=True, timeout=30)
        with open(new_exe_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
                
        bat_path = os.path.join(temp_dir, "updater.bat")
        bat_script = f"""@echo off
timeout /t 2 /nobreak > NUL
move /Y "{new_exe_path}" "{current_exe_path}"
start "" "{current_exe_path}" --just-updated
del "%~f0"
"""
        with open(bat_path, 'w') as f:
            f.write(bat_script)
            
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        subprocess.Popen([bat_path], startupinfo=startupinfo, creationflags=subprocess.CREATE_NO_WINDOW)
        
        sys.exit(0)
        
    except Exception:
        pass
