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

def main(facebot_script_path=""):
    install_python = False

    print_status(f"Checking FaceBot directory: {FACEBOT_DIR}")
    if not os.path.exists(FACEBOT_DIR):
        os.makedirs(FACEBOT_DIR)
        print_status(f"Directory {FACEBOT_DIR} created.")
    else:
        print_status(f"Directory {FACEBOT_DIR} already exists.")

    print_status("Checking Python installation...")
    python_version = None
    if check_command("python"):
        try:
            result = subprocess.run(["python", "--version"], capture_output=True, text=True, check=True)
            version_str = result.stdout.split()[1]
            python_version = tuple(map(int, version_str.split(".")))
            if python_version >= PYTHON_MIN_VERSION:
                print_status(f"Python {version_str} is installed.")
            else:
                print_status(f"Python version {version_str} is too old. At least Python {'.'.join(map(str, PYTHON_MIN_VERSION))} is required.", "yellow")
                install_python = True
        except (subprocess.CalledProcessError, ValueError):
            print_status("Error retrieving Python version.", "red")
            install_python = True
    else:
        print_status("Python not found.", "yellow")
        install_python = True

    if install_python:
        print_status("Please install Python 3.8 or higher manually from: https://www.python.org/downloads/. Ensure 'Add Python to PATH' is checked during installation.", "red")
        sys.exit(1)

    print_status("Checking tkinter...")
    if not check_tkinter():
        print_status("Attempting to install tkinter...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "tk", "--quiet"], check=True)
            print_status("tkinter successfully installed.")
        except subprocess.CalledProcessError:
            print_status("Failed to install tkinter. Ensure your Python installation includes tkinter or install it manually.", "red")
            sys.exit(1)

    print_status("Updating pip...")
    try:
        subprocess.run([sys.executable, "-m", "pip", "install", "--upgrade", "pip", "--quiet"], check=True)
        print_status("pip successfully updated.")
    except subprocess.CalledProcessError:
        print_status("Error updating pip. Ensure you have an active internet connection.", "red")
        sys.exit(1)

    print_status("Note: Visual C++ Build Tools may be required for some modules (e.g., sounddevice). If installation fails, download them from: https://aka.ms/vs/17/release/vs_BuildTools.exe", "yellow")

    print_status("Installing Python modules...")
    for module in REQUIRED_MODULES:
        print_status(f"Checking/Installing {module}...")
        try:
            result = subprocess.run([sys.executable, "-m", "pip", "show", module], capture_output=True, text=True)
            if result.returncode == 0:
                print_status(f"{module} is already installed.")
            else:
                subprocess.run([sys.executable, "-m", "pip", "install", module, "--quiet"], check=True)
                print_status(f"{module} successfully installed.")
        except subprocess.CalledProcessError:
            print_status(f"Error installing {module}. Check your internet connection or install manually with 'pip install {module}'.", "red")
            sys.exit(1)

    print_status("Installing spaCy language model (en_core_web_sm)...")
    try:
        result = subprocess.run([sys.executable, "-m", "spacy", "download", "en_core_web_sm"], capture_output=True, text=True)
        if result.returncode == 0:
            print_status("spaCy model en_core_web_sm successfully installed.")
        else:
            print_status("Error installing spaCy model. Install it manually with: 'python -m spacy download en_core_web_sm'.", "red")
            sys.exit(1)
    except subprocess.CalledProcessError:
        print_status("Error installing spaCy model. Ensure internet connection and try: 'python -m spacy download en_core_web_sm'.", "red")
        sys.exit(1)

    print_status("Checking WinSCP...")
    if os.path.exists(WINSCP_PATH):
        print_status(f"WinSCP is installed at: {WINSCP_PATH}")
    else:
        print_status("WinSCP not found. Recommended for SFTP functions. Download it from: https://winscp.net/eng/download.php", "yellow")

    print_status("Checking PuTTY...")
    if os.path.exists(PUTTY_PATH):
        print_status(f"PuTTY is installed at: {PUTTY_PATH}")
    else:
        print_status("PuTTY not found. Recommended for SSH functions. Download it from: https://www.chiark.greenend.org.uk/~sgtatham/putty/latest.html", "yellow")

    if facebot_script_path and os.path.exists(facebot_script_path):
        print_status(f"Copying facebot.py to {FACEBOT_DIR}...")
        try:
            shutil.copy(facebot_script_path, os.path.join(FACEBOT_DIR, "facebot.py"))
            print_status("facebot.py successfully copied.")
        except Exception as e:
            print_status(f"Error copying facebot.py: {e}. Ensure the file is accessible and the destination is writable.", "red")
            sys.exit(1)
    else:
        print_status(f"No valid facebot.py path provided or file not found. Copy facebot.py manually to {FACEBOT_DIR} or provide the correct path with --facebot-script-path.", "yellow")

    print_status("Creating .env file for FaceBot configuration...")
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
            print_status(f".env file created at {env_path}. Please update it with your Discord and server credentials.")
        except Exception as e:
            print_status(f"Error creating .env file: {e}. Create it manually in {FACEBOT_DIR} with required environment variables.", "red")
            sys.exit(1)
    else:
        print_status(f".env file already exists at {env_path}.")

    print_status("Installation completed successfully!", "green")
    print_status("How to start FaceBot:")
    print_status(f"1. Navigate to: cd {FACEBOT_DIR}")
    print_status("2. Update the .env file with your credentials (e.g., DISCORD_EMAIL, SERVER_HOST)")
    print_status("3. Start the bot: python facebot.py")
    print_status("If issues occur, ensure WinSCP/PuTTY are installed, the .env file is configured, and check your internet connection.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Installs FaceBot dependencies.")
    parser.add_argument("--facebot-script-path", default="", help="Path to facebot.py")
    args = parser.parse_args()
    main(args.facebot_script_path)