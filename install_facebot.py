import os
import sys
import subprocess
import shutil
import argparse
import importlib.util

# Configuration
FACEBOT_DIR = os.path.expandvars(r"%userprofile%\Documents\FaceBot")
PYTHON_MIN_VERSION = (3, 8)
REQUIRED_MODULES = [
    "selenium",
    "webdriver-manager",
    "gtts",
    "pygame",
    "sounddevice",
    "numpy",
    "spacy",
    "customtkinter",
    "python-dotenv",
    "pyautogui",
    "pywin32",
    "fuzzywuzzy",
    "python-Levenshtein",
    "cryptography",
    "speechrecognition"
]
WINSCP_PATH = os.getenv("WINSCP_PATH", r"C:\Program Files (x86)\WinSCP\WinSCP.exe")
PUTTY_PATH = os.getenv("PUTTY_PATH", r"C:\Program Files\PuTTY\putty.exe")

def print_status(message, color="green"):
    colors = {"green": "\033[92m", "yellow": "\033[93m", "red": "\033[91m", "reset": "\033[0m"}
    print(f"{colors.get(color, '')}{message}{colors['reset']}")

def check_command(command):
    return shutil.which(command) is not None

def check_tkinter():
    try:
        importlib.util.find_spec("tkinter")
        print_status("tkinter is available.")
        return True
    except ImportError:
        print_status("tkinter not found. Install it with: 'pip install tk' or ensure your Python installation includes tkinter (usually included on Windows).", "red")
        return False

def install_spacy_model(model_name="en_core_web_sm"):
    """Installiert das spaCy-Modell mit Wiederholungslogik."""
    for attempt in range(3):
        try:
            print_status(f"Versuche, spaCy-Modell {model_name} zu installieren (Versuch {attempt + 1}/3)...")
            result = subprocess.run([sys.executable, "-m", "spacy", "download", model_name], capture_output=True, text=True, timeout=300)
            if result.returncode == 0:
                print_status(f"spaCy-Modell {model_name} erfolgreich installiert.")
                return True
            else:
                print_status(f"Fehler beim Installieren von {model_name}: {result.stderr}", "yellow")
        except subprocess.TimeoutExpired:
            print_status(f"Timeout beim Installieren von {model_name}. Überprüfe die Internetverbindung.", "yellow")
        except subprocess.CalledProcessError as e:
            print_status(f"Fehler beim Installieren von {model_name}: {e}", "yellow")
        time.sleep(2)
    print_status(f"Fehler: Konnte {model_name} nicht installieren. Installiere es manuell mit 'python -m spacy download {model_name}'.", "red")
    return False

def main(facebot_script_path=""):
    install_python = False

    print_status(f"Prüfe FaceBot-Verzeichnis: {FACEBOT_DIR}")
    if not os.path.exists(FACEBOT_DIR):
        os.makedirs(FACEBOT_DIR)
        print_status(f"Verzeichnis {FACEBOT_DIR} erstellt.")
    else:
        print_status(f"Verzeichnis {FACEBOT_DIR} existiert bereits.")

    print_status("Prüfe Python-Installation...")
    python_version = None
    if check_command("python"):
        try:
            result = subprocess.run(["python", "--version"], capture_output=True, text=True, check=True)
            version_str = result.stdout.split()[1]
            python_version = tuple(map(int, version_str.split(".")))
            if python_version >= PYTHON_MIN_VERSION:
                print_status(f"Python {version_str} ist installiert.")
            else:
                print_status(f"Python-Version {version_str} ist zu alt. Mindestens Python {'.'.join(map(str, PYTHON_MIN_VERSION))} erforderlich.", "yellow")
                install_python = True
        except (subprocess.CalledProcessError, ValueError):
            print_status("Fehler beim Abrufen der Python-Version.", "red")
            install_python = True
    else:
        print_status("Python nicht gefunden.", "yellow")
        install_python = True

    if install_python:
        print_status("Bitte installiere Python 3.8 oder höher manuell von: https://www.python.org/downloads/. Stelle sicher, dass 'Add Python to PATH' während der Installation aktiviert ist.", "red")
        sys.exit(1)

    print_status("Prüfe tkinter...")
    if not check_tkinter():
        print_status("Versuche, tkinter zu installieren...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "tk", "--quiet"], check=True)
            print_status("tkinter erfolgreich installiert.")
        except subprocess.CalledProcessError:
            print_status("Fehler beim Installieren von tkinter. Stelle sicher, dass tkinter in deiner Python-Installation enthalten ist oder installiere es manuell.", "red")
            sys.exit(1)

    print_status("Aktualisiere pip...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip", "--quiet"], check=True)
        print_status("pip erfolgreich aktualisiert.")
    except subprocess.CalledProcessError:
        print_status("Fehler beim Aktualisieren von pip. Stelle sicher, dass du eine aktive Internetverbindung hast.", "red")
        sys.exit(1)

    print_status("Hinweis: Visual C++ Build Tools können für einige Module (z.B. sounddevice) erforderlich sein. Falls die Installation fehlschlägt, lade sie von: https://aka.ms/vs/17/release/vs_BuildTools.exe herunter", "yellow")

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
            print_status(f"Fehler beim Installieren von {module}. Überprüfe deine Internetverbindung oder installiere manuell mit 'pip install {module}'.", "red")
            sys.exit(1)

    print_status("Installiere spaCy-Sprachmodell (en_core_web_sm)...")
    if not install_spacy_model("en_core_web_sm"):
        print_status("Installation von spaCy-Modell fehlgeschlagen. FaceBot kann ohne dieses Modell nicht ausgeführt werden.", "red")
        sys.exit(1)

    print_status("Prüfe WinSCP...")
    if os.path.exists(WINSCP_PATH):
        print_status(f"WinSCP ist installiert unter: {WINSCP_PATH}")
    else:
        print_status("WinSCP nicht gefunden. Für SFTP-Funktionen empfohlen. Lade es von: https://winscp.net/eng/download.php herunter", "yellow")

    print_status("Prüfe PuTTY...")
    if os.path.exists(PUTTY_PATH):
        print_status(f"PuTTY ist installiert unter: {PUTTY_PATH}")
    else:
        print_status("PuTTY nicht gefunden. Für SSH-Funktionen empfohlen. Lade es von: https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html herunter", "yellow")

    if facebot_script_path and os.path.exists(facebot_script_path):
        print_status(f"Kopiere facebot.py nach {FACEBOT_DIR}...")
        try:
            shutil.copy(facebot_script_path, os.path.join(FACEBOT_DIR, "facebot.py"))
            print_status("facebot.py erfolgreich kopiert.")
        except Exception as e:
            print_status(f"Fehler beim Kopieren von facebot.py: {e}. Stelle sicher, dass die Datei zugänglich ist und das Zielverzeichnis beschreibbar ist.", "red")
            sys.exit(1)
    else:
        print_status(f"Kein gültiger facebot.py-Pfad angegeben oder Datei nicht gefunden. Kopiere facebot.py manuell nach {FACEBOT_DIR} oder gib den korrekten Pfad mit --facebot-script-path an.", "yellow")

    print_status("Erstelle .env-Datei für FaceBot-Konfiguration...")
    env_path = os.path.join(FACEBOT_DIR, ".env")
    if not os.path.exists(env_path):
        try:
            with open(env_path, "w") as f:
                f.write(
                    f"WINSCP_PATH={WINSCP_PATH}\n"
                    f"PUTTY_PATH={PUTTY_PATH}\n"
                    f"BASE_SEARCH_DIR={os.path.expandvars(r'%userprofile%')}\n"
                    f"SPOTIFY_SEARCH_URL=https://open.spotify.com/search/{{}}\n"
                    f"LETA_SEARCH_URL=https://leta.mullvad.net/search?q={{}}&engine=brave\n"
                    f"DISCORD_LOGIN_URL=https://discord.com/login\n"
                    f"TRACKLIST_CSS=div[data-testid='tracklist-row']\n"
                    f"DISCORD_EMAIL_CSS=input[name='email']\n"
                    f"DISCORD_PASSWORD_CSS=input[name='password']\n"
                    f"DISCORD_SUBMIT_CSS=button[type='submit']\n"
                    f"DISCORD_MESSAGE_CSS=div[role='textbox']\n"
                    f"CONFIG_FILE={os.path.join(os.path.expanduser('~'), '.facebot_config.json')}\n"
                    f"ENCRYPTION_KEY_FILE={os.path.join(os.path.expanduser('~'), '.facebot_key')}\n"
                    f"ENABLE_SPEECH=False\n"
                    f"SPEECH_LANGUAGE=de\n"
                    f"ENABLE_LISTENING=True\n"
                    f"DISCORD_EMAIL=\n"
                    f"DISCORD_PASSWORD=\n"
                    f"SERVER_HOST=\n"
                    f"SERVER_USERNAME=\n"
                    f"SERVER_PASSWORD=\n"
                    f"SERVER_KEY_PATH=\n"
                )
            print_status(f".env-Datei erstellt unter {env_path}. Bitte aktualisiere sie mit deinen Discord- und Server-Zugangsdaten.")
        except Exception as e:
            print_status(f"Fehler beim Erstellen der .env-Datei: {e}. Erstelle sie manuell in {FACEBOT_DIR} mit den erforderlichen Umgebungsvariablen.", "red")
            sys.exit(1)
    else:
        print_status(f".env-Datei existiert bereits unter {env_path}.")

    print_status("Installation erfolgreich abgeschlossen!", "green")
    print_status("So startest du FaceBot:")
    print_status(f"1. Navigiere zu: cd {FACEBOT_DIR}")
    print_status("2. Aktualisiere die .env-Datei mit deinen Zugangsdaten (z.B. DISCORD_EMAIL, SERVER_HOST)")
    print_status("3. Starte den Bot: python facebot.py")
    print_status("Falls Probleme auftreten, stelle sicher, dass WinSCP/PuTTY installiert sind, die .env-Datei konfiguriert ist und die Internetverbindung funktioniert.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Installiert FaceBot-Abhängigkeiten.")
    parser.add_argument("--facebot-script-path", default="", help="Pfad zu facebot.py")
    args = parser.parse_args()
    main(args.facebot_script_path)