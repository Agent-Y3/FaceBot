import os
import sys
import subprocess
import urllib.request
import shutil
import argparse
import importlib.util

# Konfiguration
FACEBOT_DIR = r"C:\Users\xByYu\Documents\FaceBot"
PYTHON_MIN_VERSION = (3, 8)
REQUIRED_MODULES = [
    "selenium",
    "webdriver-manager",
    "gtts",
    "pygame",
    "pyaudio",
    "pywin32",
    "fuzzywuzzy",
    "python-Levenshtein",
    "cryptography"
]
WINSCP_PATH = r"C:\Program Files (x86)\WinSCP\WinSCP.exe"
PUTTY_PATH = r"C:\Program Files\PuTTY\putty.exe"

def print_status(message, color="green"):
    """Gibt eine Nachricht mit Farbe aus (simuliert PowerShell-Farben)."""
    colors = {"green": "\033[92m", "yellow": "\033[93m", "red": "\033[91m", "reset": "\033[0m"}
    print(f"{colors.get(color, '')}{message}{colors['reset']}")

def check_command(command):
    """Prüft, ob ein Befehl verfügbar ist."""
    return shutil.which(command) is not None

def check_tkinter():
    """Prüft, ob tkinter verfügbar ist."""
    try:
        importlib.util.find_spec("tkinter")
        print_status("tkinter ist verfügbar.")
        return True
    except ImportError:
        print_status("tkinter nicht gefunden. Stelle sicher, dass deine Python-Installation tkinter enthält (normalerweise standardmäßig enthalten).", "red")
        return False

def main(facebot_script_path=""):
    # Initialisiere install_python
    install_python = False

    # 1. Verzeichnis erstellen
    print_status(f"Prüfe FaceBot-Verzeichnis: {FACEBOT_DIR}")
    if not os.path.exists(FACEBOT_DIR):
        os.makedirs(FACEBOT_DIR)
        print_status(f"Verzeichnis {FACEBOT_DIR} erstellt.")
    else:
        print_status(f"Verzeichnis {FACEBOT_DIR} existiert bereits.")

    # 2. Python prüfen
    print_status("Prüfe Python-Installation...")
    python_version = None
    if check_command("python"):
        try:
            result = subprocess.run(["python", "--version"], capture_output=True, text=True, check=True)
            python_version = tuple(map(int, result.stdout.split()[1].split(".")))
            if python_version >= PYTHON_MIN_VERSION:
                print_status(f"Python {'.'.join(map(str, python_version))} ist installiert.")
            else:
                print_status(f"Python-Version {'.'.join(map(str, python_version))} ist zu alt. Mindestens {'.'.join(map(str, PYTHON_MIN_VERSION))} erforderlich.", "yellow")
                install_python = True
        except subprocess.CalledProcessError:
            print_status("Fehler beim Abrufen der Python-Version.", "red")
            install_python = True
    else:
        print_status("Python nicht gefunden.", "yellow")
        install_python = True

    if install_python:
        print_status("Python-Installation wird nicht automatisch unterstützt. Bitte installiere Python 3.10 manuell von: https://www.python.org/downloads/", "red")
        sys.exit(1)

    # 3. tkinter prüfen
    print_status("Prüfe tkinter...")
    if not check_tkinter():
        sys.exit(1)

    # 4. pip aktualisieren
    print_status("Aktualisiere pip...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip", "--quiet"], check=True)
        print_status("pip erfolgreich aktualisiert.")
    except subprocess.CalledProcessError:
        print_status("Fehler beim Aktualisieren von pip.", "red")
        sys.exit(1)

    # 5. Visual C++ Build Tools können nicht direkt installiert werden
    print_status("Hinweis: Für pyaudio sind Visual C++ Build Tools erforderlich. Lade sie bei Bedarf von: https://aka.ms/vs/16/release/vs_buildtools.exe", "yellow")

    # 6. Python-Module installieren
    print_status("Installiere Python-Module...")
    for module in REQUIRED_MODULES:
        print_status(f"Prüfe/Installiere {module}...")
        try:
            result = subprocess.run([sys.executable, "-m", "pip", "show", module], capture_output=True, text=True)
            if result.returncode == 0:
                print_status(f"{module} ist bereits installiert.")
            else:
                subprocess.run([sys.executable, "-m", "pip", "install", module, "--quiet"], check=True)
                print_status(f"{module} erfolgreich installiert.")
        except subprocess.CalledProcessError:
            print_status(f"Fehler beim Installieren von {module}.", "red")
            sys.exit(1)

    # 7. WinSCP prüfen
    print_status("Prüfe WinSCP...")
    if os.path.exists(WINSCP_PATH):
        print_status(f"WinSCP ist installiert unter: {WINSCP_PATH}")
    else:
        print_status("WinSCP nicht gefunden. Für SFTP-Funktionen empfohlen. Lade es herunter: https://winscp.net/eng/download.php", "yellow")

    # 8. PuTTY prüfen
    print_status("Prüfe PuTTY...")
    if os.path.exists(PUTTY_PATH):
        print_status(f"PuTTY ist installiert unter: {PUTTY_PATH}")
    else:
        print_status("PuTTY nicht gefunden. Für SSH-Funktionen empfohlen. Lade es herunter: https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html", "yellow")

    # 9. facebot.py kopieren
    if facebot_script_path and os.path.exists(facebot_script_path):
        print_status(f"Kopiere facebot.py nach {FACEBOT_DIR}...")
        try:
            shutil.copy(facebot_script_path, os.path.join(FACEBOT_DIR, "facebot.py"))
            print_status("facebot.py erfolgreich kopiert.")
        except Exception as e:
            print_status(f"Fehler beim Kopieren von facebot.py: {e}", "red")
            sys.exit(1)
    else:
        print_status(f"Kein facebot.py-Pfad angegeben oder Datei nicht gefunden. Kopiere facebot.py manuell nach {FACEBOT_DIR}.", "yellow")

    # 10. Abschluss
    print_status("Installation abgeschlossen!", "green")
    print_status("So startest du FaceBot:")
    print_status(f"1. Navigiere zu: cd {FACEBOT_DIR}")
    print_status("2. Starte den Bot: python facebot.py")
    print_status("Falls Probleme auftreten, überprüfe die Internetverbindung und die Installation von WinSCP/PuTTY.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Installiert FaceBot-Abhängigkeiten.")
    parser.add_argument("--facebot-script-path", default="", help="Pfad zu facebot.py")
    args = parser.parse_args()
    main(args.facebot_script_path)