import os
import subprocess
import sys
import shutil
import winreg
import time
from pathlib import Path

def run_command(command, shell=True, check=True):
    """Führt einen Shell-Befehl aus und gibt die Ausgabe zurück."""
    try:
        result = subprocess.run(command, shell=shell, check=check, text=True, capture_output=True)
        print(f"Erfolg: {command}")
        return result
    except subprocess.CalledProcessError as e:
        print(f"Fehler bei Befehl '{command}': {e.stderr}")
        return None
    except Exception as e:
        print(f"Fehler bei Befehl '{command}': {e}")
        return None

def check_python_version():
    """Prüft, ob Python 3.8–3.11 verwendet wird."""
    version = sys.version_info
    if not (3, 8) <= (version.major, version.minor) <= (3, 11):
        print(f"FEHLER: Python {version.major}.{version.minor} wird nicht unterstützt. Verwende Python 3.8–3.11.")
        sys.exit(1)
    print(f"Python-Version: {version.major}.{version.minor} (OK)")

def enable_long_paths():
    """Aktiviert Long Path Support in der Windows-Registrierung."""
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SYSTEM\CurrentControlSet\Control\FileSystem", 0, winreg.KEY_WRITE)
        winreg.SetValueEx(key, "LongPathsEnabled", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        print("Long Path Support aktiviert. Bitte starte den Computer neu, falls dies der erste Lauf ist.")
        print("Drücke Enter, um fortzufahren, oder starte den Computer neu und führe das Skript erneut aus.")
        input()
    except Exception as e:
        print(f"Fehler beim Aktivieren von Long Path Support: {e}")
        print("Manuelle Anleitung: Öffne regedit, gehe zu HKEY_LOCAL_MACHINE\\SYSTEM\\CurrentControlSet\\Control\\FileSystem, erstelle/setze LongPathsEnabled auf 1.")
        input("Drücke Enter, um fortzufahren...")

def create_virtual_env(env_path="C:\\facebot_env"):
    """Erstellt eine virtuelle Umgebung."""
    env_path = Path(env_path)
    if env_path.exists():
        print(f"Virtuelle Umgebung {env_path} existiert bereits. Überspringe Erstellung.")
    else:
        print(f"Erstelle virtuelle Umgebung in {env_path}...")
        run_command(f'python -m venv "{env_path}"')
    return env_path

def activate_virtual_env(env_path):
    """Gibt den Befehl zum Aktivieren der virtuellen Umgebung zurück."""
    activate_script = env_path / "Scripts" / "activate.bat"
    return f'"{activate_script}" && '

def update_pip(env_path):
    """Aktualisiert pip in der virtuellen Umgebung."""
    print("Aktualisiere pip...")
    activate = activate_virtual_env(env_path)
    run_command(f'{activate} python -m pip install --upgrade pip')

def install_dependencies(env_path):
    """Installiert alle Abhängigkeiten."""
    print("Installiere Abhängigkeiten...")
    activate = activate_virtual_env(env_path)
    
    # Zuerst NumPy in kompatibler Version installieren
    print("Behebe NumPy-Kompatibilität...")
    run_command(f'{activate} pip uninstall numpy -y', check=False)
    run_command(f'{activate} pip install numpy==1.26.4')
    
    # Alle anderen Abhängigkeiten
    dependencies = [
        "psutil",
        "requests",
        "selenium",
        "webdriver-manager",
        "pyautogui",
        "pillow",
        "opencv-python==4.8.1.78",  # Spezifische Version für Stabilität
        "pytesseract",
        "speechrecognition",
        "pyaudio",
        "elevenlabs",
        "gtts",
        "pygame",
        "spacy"
    ]
    run_command(f'{activate} pip install {" ".join(dependencies)}')
    
    # Versuche pipwin und pyaudio, falls pyaudio fehlschlägt
    print("Prüfe pyaudio-Installation...")
    run_command(f'{activate} pip install pipwin', check=False)
    run_command(f'{activate} pipwin install pyaudio', check=False)

def install_spacy_model(env_path, model="de_core_news_lg"):
    """Installiert das spaCy-Modell."""
    print(f"Installiere spaCy-Modell {model}...")
    activate = activate_virtual_env(env_path)
    result = run_command(f'{activate} python -m spacy download {model}', check=False)
    if result and result.returncode != 0:
        print(f"Fehler beim Installieren von {model}. Versuche kleineres Modell de_core_news_md...")
        run_command(f'{activate} python -m spacy download de_core_news_md')
        return "de_core_news_md"
    return model

def add_to_path(path_to_add):
    """Fügt einen Pfad zur Umgebungsvariablen PATH hinzu."""
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, "Environment", 0, winreg.KEY_READ | winreg.KEY_WRITE)
        current_path, _ = winreg.QueryValueEx(key, "Path")
        if path_to_add not in current_path:
            new_path = f"{current_path};{path_to_add}"
            winreg.SetValueEx(key, "Path", 0, winreg.REG_EXPAND_SZ, new_path)
            print(f"{path_to_add} zu PATH hinzugefügt.")
        else:
            print(f"{path_to_add} bereits in PATH.")
        winreg.CloseKey(key)
    except Exception as e:
        print(f"Fehler beim Hinzufügen zu PATH: {e}")
        print(f"Manuell hinzufügen: setx PATH \"%PATH%;{path_to_add}\" /M")

def check_tesseract():
    """Prüft, ob Tesseract installiert ist, und fügt es zum PATH hinzu."""
    tesseract_path = r"C:\Program Files\Tesseract-OCR"
    tesseract_exe = os.path.join(tesseract_path, "tesseract.exe")
    if os.path.exists(tesseract_exe):
        print("Tesseract gefunden.")
        add_to_path(tesseract_path)
        return True
    else:
        print("Tesseract-OCR nicht gefunden!")
        print("Bitte lade Tesseract von https://github.com/UB-Mannheim/tesseract/wiki herunter und installiere es in C:\\Program Files\\Tesseract-OCR.")
        print("Führe das Skript nach der Installation erneut aus.")
        input("Drücke Enter, um fortzufahren...")
        return False

def test_installation(env_path):
    """Testet, ob alle Bibliotheken importiert werden können."""
    print("Teste Installation...")
    test_code = """
import tkinter
import psutil
import requests
import selenium
import webdriver_manager
import pyautogui
from PIL import Image
import cv2
import pytesseract
import speech_recognition
import pyaudio
from elevenlabs import ElevenLabs
from gtts import gTTS
import pygame
import spacy
import numpy
print("Alle Bibliotheken erfolgreich importiert!")
print("NumPy-Version:", numpy.__version__)
"""
    test_file = "test_libs.py"
    with open(test_file, "w", encoding="utf-8") as f:
        f.write(test_code)
    
    activate = activate_virtual_env(env_path)
    result = run_command(f'{activate} python {test_file}', check=False)
    if result and result.returncode == 0:
        print("Installation erfolgreich!")
    else:
        print("Fehler bei der Installation. Überprüfe die Fehlermeldungen oben.")
    os.remove(test_file)

def main():
    print("=== FaceBot Abhängigkeiten Installation ===")
    
    # Python-Version prüfen
    check_python_version()
    
    # Long Path Support aktivieren
    enable_long_paths()
    
    # Virtuelle Umgebung erstellen
    env_path = create_virtual_env("C:\\facebot_env")
    
    # pip aktualisieren
    update_pip(env_path)
    
    # Abhängigkeiten installieren
    install_dependencies(env_path)
    
    # spaCy-Modell installieren
    spacy_model = install_spacy_model(env_path)
    
    # Scripts-Ordner zu PATH hinzufügen
    scripts_path = env_path / "Scripts"
    add_to_path(str(scripts_path))
    
    # Tesseract prüfen und zu PATH hinzufügen
    tesseract_installed = check_tesseract()
    
    # Installation testen
    test_installation(env_path)
    
    # Abschlussmeldung
    print("\n=== Installation abgeschlossen ===")
    if tesseract_installed:
        print("Alle Abhängigkeiten sollten installiert sein!")
        print("Nächste Schritte:")
        print("1. Stelle sicher, dass facebot.py in C:\\Users\\xByYu\\Documents gespeichert ist.")
        print("2. Erstelle captcha_checkbox.png und discord_input.png (siehe Anweisungen).")
        print("3. Optional: Füge deinen ElevenLabs API-Schlüssel in facebot.py hinzu.")
        print("4. Aktiviere die virtuelle Umgebung:")
        print(f"   C:\\facebot_env\\Scripts\\activate")
        print("5. Starte FaceBot:")
        print("   python facebot.py")
    else:
        print("Tesseract fehlt. Installiere Tesseract und führe das Skript erneut aus.")
    print("\nFalls Fehler auftreten, notiere die Fehlermeldung und kontaktiere den Support.")
    input("Drücke Enter, um zu beenden...")

if __name__ == "__main__":
    main()