import tkinter as tk
import customtkinter as ctk
import os
import json
import logging
import asyncio
import sounddevice as sd
import numpy as np
import speech_recognition as sr
from dataclasses import dataclass
from typing import Optional, Dict, Tuple, List, Callable
from urllib.parse import quote, urlparse
from pydantic import BaseModel, ValidationError
from cryptography.fernet import Fernet
import base64
import secrets
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import win32gui
import win32con
import winreg
import time
import subprocess
import shutil
from fuzzywuzzy import fuzz
import mimetypes
import spacy
from gtts import gTTS
import io
import threading
import pyautogui
from dotenv import load_dotenv
import re
import queue
import pygame

pygame.mixer.init()
load_dotenv()
pyautogui.FAILSAFE = False
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class FaceBotError(Exception):
    pass

class ConfigError(FaceBotError):
    pass

class BrowserError(FaceBotError):
    pass

class SpeechError(FaceBotError):
    pass

class UIError(FaceBotError):
    pass

class CommandError(FaceBotError):
    pass

def handle_errors(method: Callable) -> Callable:
    async def async_wrapper(self, *args, **kwargs):
        try:
            return await method(self, *args, **kwargs)
        except (ConfigError, BrowserError, SpeechError, UIError, CommandError) as e:
            logger = kwargs.get('logger') or getattr(self, 'logger', logging.getLogger("FaceBot"))
            logger.log_message(f"{method.__name__.replace('_', ' ').title()} failed: {str(e)}") if hasattr(logger, 'log_message') else logger.error(f"{method.__name__.replace('_', ' ').title()} failed: {str(e)}")
            return None
        except Exception as e:
            logger = kwargs.get('logger') or getattr(self, 'logger', logging.getLogger("FaceBot"))
            logger.log_message(f"Unexpected error in {method.__name__.replace('_', ' ').title()}: {str(e)}") if hasattr(logger, 'log_message') else logger.error(f"Unexpected error in {method.__name__.replace('_', ' ').title()}: {str(e)}")
            return None
    def sync_wrapper(self, *args, **kwargs):
        try:
            return method(self, *args, **kwargs)
        except (ConfigError, BrowserError, SpeechError, UIError, CommandError) as e:
            logger = kwargs.get('logger') or getattr(self, 'logger', logging.getLogger("FaceBot"))
            logger.log_message(f"{method.__name__.replace('_', ' ').title()} failed: {str(e)}") if hasattr(logger, 'log_message') else logger.error(f"{method.__name__.replace('_', ' ').title()} failed: {str(e)}")
            return None
        except Exception as e:
            logger = kwargs.get('logger') or getattr(self, 'logger', logging.getLogger("FaceBot"))
            logger.log_message(f"Unexpected error in {method.__name__.replace('_', ' ').title()}: {str(e)}") if hasattr(logger, 'log_message') else logger.error(f"Unexpected error in {method.__name__.replace('_', ' ').title()}: {str(e)}")
            return None
    return async_wrapper if asyncio.iscoroutinefunction(method) else sync_wrapper

class Config(BaseModel):
    winscp_path: str = os.getenv("WINSCP_PATH", r"C:\Program Files (x86)\WinSCP\WinSCP.exe")
    putty_path: str = os.getenv("PUTTY_PATH", r"C:\Program Files\PuTTY\putty.exe")
    base_search_dir: str = os.getenv("BASE_SEARCH_DIR", os.path.expandvars(r"%userprofile%"))
    spotify_search_url: str = os.getenv("SPOTIFY_SEARCH_URL", "https://open.spotify.com/search/{}")
    leta_search_url: str = os.getenv("LETA_SEARCH_URL", "https://leta.mullvad.net/search?q={}&engine=brave")
    discord_login_url: str = os.getenv("DISCORD_LOGIN_URL", "https://discord.com/login")
    tracklist_css: str = os.getenv("TRACKLIST_CSS", "div[data-testid='tracklist-row']")
    discord_email_css: str = os.getenv("DISCORD_EMAIL_CSS", "input[name='email']")
    discord_password_css: str = os.getenv("DISCORD_PASSWORD_CSS", "input[name='password']")
    discord_submit_css: str = os.getenv("DISCORD_SUBMIT_CSS", "button[type='submit']")
    discord_message_css: str = os.getenv("DISCORD_MESSAGE_CSS", "div[role='textbox']")
    config_file: str = os.getenv("CONFIG_FILE", os.path.join(os.path.expanduser("~"), ".facebot_config.json"))
    encryption_key_file: str = os.getenv("ENCRYPTION_KEY_FILE", os.path.join(os.path.expanduser("~"), ".facebot_key"))
    enable_speech: bool = bool(os.getenv("ENABLE_SPEECH", "False").lower() == "true")
    speech_language: str = os.getenv("SPEECH_LANGUAGE", "de")
    enable_listening: bool = bool(os.getenv("ENABLE_LISTENING", "True").lower() == "true")
    discord_email: str = os.getenv("DISCORD_EMAIL", "")
    discord_password: str = os.getenv("DISCORD_PASSWORD", "")
    server_host: str = os.getenv("SERVER_HOST", "")
    server_username: str = os.getenv("SERVER_USERNAME", "")
    server_password: str = os.getenv("SERVER_PASSWORD", "")
    server_key_path: str = os.getenv("SERVER_KEY_PATH", "")

@dataclass
class Context:
    last_application: Optional[str] = None
    last_action: Optional[str] = None
    user_preferences: Dict[str, int] = None

    def __post_init__(self):
        self.user_preferences = self.user_preferences or {"music": 0, "browser": 0, "document": 0}

class ConfigManager:
    def __init__(self, logger):
        self.config = Config()
        self.fernet = None
        self.server_config = None
        self.logger = logger
        self._setup_encryption()
        self._load_config()

    @handle_errors
    def _setup_encryption(self):
        if os.path.exists(self.config.encryption_key_file):
            with open(self.config.encryption_key_file, 'rb') as f:
                key = f.read()
        else:
            key = secrets.token_bytes(32)
            with open(self.config.encryption_key_file, 'wb') as f:
                f.write(key)
        self.fernet = Fernet(base64.urlsafe_b64encode(key))

    @handle_errors
    def _encrypt_data(self, data: str) -> str:
        if not self.fernet or not data:
            return data
        return base64.b64encode(self.fernet.encrypt(data.encode())).decode()

    @handle_errors
    def _decrypt_data(self, data: str) -> str:
        if not self.fernet or not data:
            return data
        return self.fernet.decrypt(base64.b64decode(data)).decode()

    @handle_errors
    def _load_config(self):
        if not os.path.exists(self.config.config_file):
            self._save_config()
        try:
            with open(self.config.config_file, 'r') as f:
                config_data = json.load(f)
            self.server_config = config_data.get('server_config', {})
            self.server_config['password'] = self._decrypt_data(self.server_config.get('password', self.config.server_password))
            self.config.discord_email = self._decrypt_data(config_data.get('discord_email', self.config.discord_email))
            self.config.discord_password = self._decrypt_data(config_data.get('discord_password', self.config.discord_password))
            self.config.enable_speech = config_data.get('speech_enabled', self.config.enable_speech)
            self.config.enable_listening = config_data.get('enable_listening', self.config.enable_listening)
        except Exception as e:
            raise ConfigError(f"Konfiguration konnte nicht geladen werden: {e}. Verwende Standardwerte.")

    @handle_errors
    def _save_config(self):
        config_data = {
            'server_config': self.server_config or {
                'host': self.config.server_host,
                'username': self.config.server_username,
                'password': self._encrypt_data(self.config.server_password),
                'key_path': self.config.server_key_path
            },
            'discord_email': self._encrypt_data(self.config.discord_email),
            'discord_password': self._encrypt_data(self.config.discord_password),
            'speech_enabled': self.config.enable_speech,
            'enable_listening': self.config.enable_listening
        }
        with open(self.config.config_file, 'w') as f:
            json.dump(config_data, f, indent=4)

class BrowserManager:
    def __init__(self, logger):
        self.driver = None
        self.browser_name = None
        self.logger = logger
        self.session_cache = {}

    def __enter__(self):
        self.initialize()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.driver:
            self.driver.quit()
            self.driver = None

    @handle_errors
    def _get_default_browser(self) -> str:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice") as key:
            prog_id = winreg.QueryValueEx(key, "ProgId")[0]
        browser_map = {
            "Chrome": "chrome",
            "Firefox": "firefox",
            "Edge": "edge",
            "IE": "edge",
            "Opera": "opera"
        }
        return browser_map.get(prog_id.split('.')[0], "chrome")

    @handle_errors
    def initialize(self):
        if self.browser_name in self.session_cache:
            self.driver = self.session_cache[self.browser_name]
            return
        self.browser_name = self._get_default_browser()
        for attempt in range(2):
            try:
                if self.browser_name == "firefox":
                    service = FirefoxService(GeckoDriverManager().install())
                    self.driver = webdriver.Firefox(service=service)
                elif self.browser_name == "edge":
                    options = webdriver.EdgeOptions()
                    options.add_argument("--disable-features=OptimizationHints")
                    options.add_argument("--no-sandbox")
                    options.add_argument("--disable-gpu")
                    service = EdgeService(EdgeChromiumDriverManager().install())
                    self.driver = webdriver.Edge(service=service, options=options)
                elif self.browser_name == "opera":
                    options = webdriver.ChromeOptions()
                    options.binary_location = shutil.which("opera.exe") or r"C:\Program Files\Opera\opera.exe"
                    service = ChromeService(ChromeDriverManager().install())
                    self.driver = webdriver.Chrome(service=service, options=options)
                else:
                    service = ChromeService(ChromeDriverManager().install())
                    self.driver = webdriver.Chrome(service=service)
                self.driver.maximize_window()
                self.session_cache[self.browser_name] = self.driver
                break
            except Exception as e:
                if attempt == 1:
                    raise BrowserError(f"Initialisierung von {self.browser_name} fehlgeschlagen: {e}. Browserfunktionen deaktiviert.")
                self.browser_name = "chrome"

    @handle_errors
    async def navigate_to_url(self, url: str, browser: Optional[str], context: Context):
        parsed_url = urlparse(url)
        if not parsed_url.scheme or not parsed_url.netloc:
            raise BrowserError(f"Ungültige URL '{url}'")
        if not self.driver:
            self.initialize()
            if not self.driver:
                raise BrowserError("Browser konnte nicht initialisiert werden")
        browser = browser or self.browser_name or "chrome"
        self.logger.log_message(f"Navigiere zu '{url}' in {browser}...")
        context.last_action = "goto"
        context.user_preferences["browser"] += 1
        await self._focus_application(browser)
        for attempt in range(2):
            try:
                self.driver.get(url)
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                self.logger.log_message(f"Navigation zu '{url}' abgeschlossen!")
                return
            except Exception as e:
                if attempt == 1:
                    raise BrowserError(f"Fehler beim Navigieren zu '{url}': {e}")

    @handle_errors
    async def _focus_application(self, app_name: str):
        app_map = {
            "edge": "Microsoft Edge",
            "microsoft edge": "Microsoft Edge",
            "chrome": "Google Chrome",
            "firefox": "Firefox",
            "opera": "Opera",
            "word": "Microsoft Word",
            "excel": "Microsoft Excel",
            "notepad": "Notepad",
            "winscp": "WinSCP",
            "discord": "Discord"
        }
        window_title = app_map.get(app_name.lower(), app_name)
        def enum_windows_callback(hwnd, results):
            window_text = win32gui.GetWindowText(hwnd).lower()
            if window_title.lower() in window_text or app_name.lower() in window_text and win32gui.IsWindowVisible(hwnd):
                results.append(hwnd)
        handles = []
        win32gui.EnumWindows(enum_windows_callback, handles)
        if handles:
            hwnd = handles[0]
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
            self.logger.log_message(f"Anwendung '{app_name}' fokussiert.")
            return True
        self.logger.log_message(f"Anwendung '{app_name}' nicht gefunden. Starte sie...")
        await self._open_file_or_program(app_name)
        return True

    @handle_errors
    async def _open_file_or_program(self, target: str):
        self.logger.log_message(f"Öffne '{target}'...")
        program_map = {
            "microsoft edge": "msedge",
            "edge": "msedge",
            "chrome": "chrome",
            "firefox": "firefox",
            "opera": "opera",
            "word": "winword",
            "excel": "excel",
            "notepad": "notepad",
            "winscp": "WinSCP",
            "discord": "Discord"
        }
        target_lower = target.lower().strip()
        executable = program_map.get(target_lower)
        if executable:
            program_path = shutil.which(f"{executable}.exe")
            if program_path:
                pyautogui.hotkey("win", "r")
                pyautogui.write(f"{executable}.exe")
                pyautogui.press("enter")
                self.logger.log_message(f"Programm '{target}' gestartet!")
                return
        program_path = shutil.which(target)
        if program_path:
            pyautogui.hotkey("win", "r")
            pyautogui.write(program_path)
            pyautogui.press("enter")
            self.logger.log_message(f"Programm '{target}' gestartet!")
            return
        if os.path.isabs(target) and os.path.exists(target):
            os.startfile(target)
            self.logger.log_message(f"'{target}' geöffnet!")
            return
        for root, _, files in os.walk(self.config.base_search_dir, maxdepth=3):
            if target in files:
                file_path = os.path.join(root, target)
                os.startfile(file_path)
                self.logger.log_message(f"Datei '{file_path}' geöffnet!")
                return
        suggestions = self._suggest_alternatives(target)
        if suggestions:
            suggestion_text = "\n".join([f"- {name}: {reason}" for name, reason in suggestions])
            self.logger.log_message(f"'{target}' nicht gefunden. Meintest du:\n{suggestion_text}\nSage z.B. 'Öffne {suggestions[0][0]}'.")
        else:
            self.logger.log_message(f"'{target}' nicht gefunden. Gib einen gültigen Pfad oder ein Programm an.")

    @handle_errors
    def _suggest_alternatives(self, target: str) -> List[Tuple[str, str]]:
        suggestions = []
        target_lower = target.lower()
        program_map = {
            "microsoft edge": "msedge.exe",
            "edge": "msedge.exe",
            "chrome": "chrome.exe",
            "firefox": "firefox.exe",
            "opera": "opera.exe",
            "word": "winword.exe",
            "excel": "excel.exe",
            "notepad": "notepad.exe",
            "winscp": "WinSCP.exe",
            "discord": "Discord.exe"
        }
        for name in program_map.keys():
            score = fuzz.ratio(target_lower, name.lower())
            if score > 85:
                suggestions.append((name, "ähnlicher Name"))
        if "." in target:
            mime_type, _ = mimetypes.guess_type(target)
            if mime_type:
                if mime_type.startswith("text"):
                    suggestions.append(("notepad", "Textdatei"))
                    if shutil.which("winword.exe"):
                        suggestions.append(("word", "Textdatei"))
                elif mime_type.startswith("image"):
                    suggestions.append(("msedge.exe", "Bilddatei"))
                elif mime_type.startswith("application/vnd"):
                    if shutil.which("excel.exe"):
                        suggestions.append(("excel", "Tabellendokument"))
        for name, exe in program_map.items():
            if shutil.which(exe) and name not in [s[0] for s in suggestions]:
                suggestions.append((name, "auf dem System verfügbar"))
        suggestions = sorted(suggestions, key=lambda x: fuzz.ratio(target_lower, x[0]), reverse=True)[:3]
        return suggestions

class SpeechManager:
    def __init__(self, config: Config, logger):
        self.config = config
        self.recognizer = sr.Recognizer()
        self.listening = False
        self.audio_queue = queue.Queue()
        self.logger = logger
        self.nlp = spacy.load("en_core_web_sm")

    @handle_errors
    async def _speak(self, text: str):
        if not self.config.enable_speech:
            return
        tts = gTTS(text=text, lang=self.config.speech_language)
        audio_file = io.BytesIO()
        tts.write_to_fp(audio_file)
        audio_file.seek(0)
        pygame.mixer.music.load(audio_file)
        pygame.mixer.music.play()
        while pygame.mixer.music.get_busy():
            await asyncio.sleep(0.1)

    @handle_errors
    async def _update_audio_indicator(self, canvas: tk.Canvas, bars: list):
        sample_rate = 44100
        while self.listening:
            try:
                audio = sd.rec(int(0.1 * sample_rate), samplerate=sample_rate, channels=1)
                sd.wait()
                rms = np.sqrt(np.mean(audio**2))
                height = min(max(int(rms * 500), 5), 25)
                for i, bar in enumerate(bars):
                    canvas.coords(bar, 10 + i * 20, 25, 25 + i * 20, 25 - height + (i % 2) * 5)
                canvas.update()
                await asyncio.sleep(0.1)
            except Exception:
                pass

    @handle_errors
    async def listen(self, bot: 'FaceBot'):
        error_count = 0
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
                self.recognizer.dynamic_energy_threshold = True
                self.recognizer.energy_threshold = 300
                self.logger.log_message("Spracherkennung aktiv. Sprich deutlich, z.B. 'Öffne Edge'.")
                threading.Thread(target=self._update_audio_indicator, args=(bot.logger.indicator_canvas, bot.logger.bars), daemon=True).start()
                while self.listening:
                    try:
                        audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=15)
                        audio_duration = len(audio.frame_data) / (audio.sample_rate * audio.sample_width)
                        if audio_duration < 0.5:
                            error_count = 0
                            continue
                        for attempt in range(2):
                            try:
                                command = self.recognizer.recognize_google(audio, language="en-US")
                                self.logger.log_message(f"Du hast gesagt: {command}")
                                bot.root.after(0, bot.process_command, None, command)
                                error_count = 0
                                break
                            except sr.RequestError as e:
                                if attempt == 1:
                                    raise SpeechError(f"Fehler bei der Spracherkennung: {e}. Überprüfe die Internetverbindung.")
                            except sr.UnknownValueError:
                                error_count += 1
                                if error_count >= 3:
                                    self.logger.log_message("Mehrere unklare Eingaben. Sprich lauter oder überprüfe das Mikrofon.")
                                    error_count = 0
                                break
                    except sr.WaitTimeoutError:
                        error_count = 0
                    except Exception as e:
                        raise SpeechError(f"Fehler bei der Spracherkennung: {e}.")
                    await asyncio.sleep(0.1)
        except Exception as e:
            raise SpeechError(f"Fehler beim Starten der Spracherkennung: {e}. Spracherkennung deaktiviert.")
        finally:
            self.listening = False
            bot.logger.listen_button.configure(text="Mikrofon")
            for bar in bot.logger.bars:
                bot.logger.indicator_canvas.coords(bar, 10 + bot.logger.bars.index(bar) * 20, 25, 25 + bot.logger.bars.index(bar) * 20, 25)
            self.logger.log_message("Mikrofon ausgeschaltet.")

class UIManager:
    def __init__(self, root: ctk.CTk, bot: 'FaceBot'):
        self.root = root
        self.bot = bot
        self.chat_area = None
        self.input_field = None
        self.send_button = None
        self.listen_button = None
        self.config_button = None
        self.indicator_canvas = None
        self.bars = []
        self._setup_ui()

    @handle_errors
    def _setup_ui(self):
        self.root.title("FaceBot")
        self.root.geometry("600x500")
        self.chat_area = ctk.CTkTextbox(self.root, wrap="word", width=580, height=300, state="disabled")
        self.chat_area.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
        input_frame = ctk.CTkFrame(self.root)
        input_frame.grid(row=1, column=0, columnspan=3, padx=10, pady=5, sticky="ew")
        self.input_field = ctk.CTkEntry(input_frame, placeholder_text="Befehl eingeben...")
        self.input_field.grid(row=0, column=0, sticky="ew", padx=(0, 5))
        self.input_field.bind("<Return>", self.bot.process_command)
        input_frame.grid_columnconfigure(0, weight=1)
        self.send_button = ctk.CTkButton(input_frame, text="Senden", command=self.bot.process_command)
        self.send_button.grid(row=0, column=1, padx=5)
        self.listen_button = ctk.CTkButton(input_frame, text="Mikrofon", command=self.bot.toggle_listening)
        self.listen_button.grid(row=0, column=2, padx=5)
        self.config_button = ctk.CTkButton(input_frame, text="Einstellungen", command=self.bot._open_config_ui)
        self.config_button.grid(row=0, column=3, padx=5)
        self.indicator_canvas = ctk.CTkCanvas(self.root, width=100, height=30, bg='black', highlightthickness=0)
        self.indicator_canvas.grid(row=2, column=0, columnspan=3, pady=5)
        for i in range(5):
            x = 10 + i * 20
            bar = self.indicator_canvas.create_rectangle(x, 25, x + 15, 25, fill='green')
            self.bars.append(bar)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

    def log_message(self, message: str):
        def update_gui():
            try:
                self.chat_area.configure(state="normal")
                self.chat_area.insert("end", f"{message}\n")
                self.chat_area.configure(state="disabled")
                self.chat_area.see("end")
                self.root.update()
            except Exception as e:
                logging.getLogger("FaceBot").error(f"UI-Aktualisierung fehlgeschlagen: {e}")
        self.root.after(0, update_gui)
        logging.getLogger("FaceBot").info(message)
        asyncio.create_task(self.bot.speech_manager._speak(message))

class CommandRegistry:
    def __init__(self, logger):
        self.commands = {}
        self.logger = logger
        self.nlp = spacy.load("en_core_web_sm")

    def register(self, intent: str, keywords: List[str]):
        def decorator(func: Callable):
            self.commands[intent] = {"func": func, "keywords": keywords}
            return func
        return decorator

    @handle_errors
    def parse(self, command: str, context: Context) -> Tuple[Optional[str], Dict[str, str]]:
        doc = self.nlp(command.lower().strip())
        intent = None
        params = {}
        max_score = 0
        for intent_name, cmd_data in self.commands.items():
            for keyword in cmd_data["keywords"]:
                score = fuzz.ratio(keyword, command.lower())
                if score > max_score and score > 85:
                    max_score = score
                    intent = intent_name
                    remaining = command.lower()
                    for kw in cmd_data["keywords"]:
                        remaining = remaining.replace(kw, "").strip()
                    if intent == "search":
                        params["search_term"] = remaining.replace("for", "").strip()
                        for token in doc:
                            if token.text.lower() in ["chrome", "firefox", "edge", "opera"]:
                                params["browser"] = token.text.lower()
                    elif intent == "goto":
                        for token in doc:
                            if token.like_url:
                                params["url"] = token.text
                                for t in doc:
                                    if t.text.lower() in ["chrome", "firefox", "edge", "opera"]:
                                        params["browser"] = t.text.lower()
                                break
                    elif intent in ["open", "play", "close", "maximize", "save", "upload", "task"]:
                        params["target"] = remaining
                    elif intent == "write":
                        params["text"] = remaining.replace("in word", "").replace("in excel", "").replace("in notepad", "").strip()
                        for token in doc:
                            if token.text.lower() in ["word", "excel", "notepad"]:
                                params["app"] = token.text.lower()
                    elif intent == "discord":
                        for i, token in enumerate(doc):
                            if token.text.lower() == "to":
                                params["target"] = doc[i+1].text if i+1 < len(doc) else ""
                                params["message"] = " ".join(t.text for t in doc[i+2:]) if i+2 < len(doc) else ""
                                break
        if intent in ["play", "search", "goto"]:
            context.user_preferences["music" if intent == "play" else "browser"] += 1
        elif intent in ["write", "save"] and params.get("app") in ["word", "excel"]:
            context.user_preferences["document"] += 1
        return intent, params

class FaceBot:
    def __init__(self, root: ctk.CTk):
        self.root = root
        self.config = Config()
        self.context = Context()
        self.logger = UIManager(root, self)
        self.browser_manager = BrowserManager(self.logger)
        self.speech_manager = SpeechManager(self.config, self.logger)
        self.config_manager = ConfigManager(self.logger)
        self.command_registry = CommandRegistry(self.logger)
        self._setup_logger()
        self._register_commands()
        try:
            with self.browser_manager:
                self.context.last_application = self.browser_manager.browser_name
                self.logger.log_message(f"Okay, ich bin bereit! Verwende Browser: {self.browser_manager.browser_name or 'Unbekannt'}. Sage z.B. 'Öffne Edge', 'Suche nach xAI'.")
        except BrowserError as e:
            self.logger.log_message(f"Browserinitialisierung fehlgeschlagen: {e}. Browserfunktionen deaktiviert.")
            self.browser_manager.browser_name = "chrome"

    def _setup_logger(self):
        logger = logging.getLogger("FaceBot")
        logger.setLevel(logging.DEBUG)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(handler)

    def _register_commands(self):
        @self.command_registry.register("open", ["open", "start", "launch", "öffnen", "starte"])
        async def open_command(self, params):
            target = params.get("target")
            if not target:
                raise CommandError("Was soll geöffnet werden? Sage z.B. 'Öffne Edge'.")
            await self.browser_manager._open_file_or_program(target)

        @self.command_registry.register("play", ["play", "music", "spiele", "musik"])
        async def play_command(self, params):
            song_name = params.get("target")
            if not song_name:
                raise CommandError("Welchen Song soll ich abspielen? Sage z.B. 'Spiele Shape of You'.")
            await self._play_spotify_song(song_name)

        @self.command_registry.register("search", ["search", "google", "find", "suche", "finde"])
        async def search_command(self, params):
            search_term = params.get("search_term")
            browser = params.get("browser", self.browser_manager.browser_name)
            if not search_term:
                raise CommandError("Wonach soll ich suchen? Sage z.B. 'Suche nach xAI'.")
            await self._search_leta(search_term, browser)

        @self.command_registry.register("goto", ["go to", "goto", "navigate to", "gehe zu"])
        async def goto_command(self, params):
            url = params.get("url")
            browser = params.get("browser", self.browser_manager.browser_name)
            if not url:
                raise CommandError("Welche Website soll ich öffnen? Sage z.B. 'Gehe zu https://check24.de'.")
            await self.browser_manager.navigate_to_url(url, browser, self.context)

        @self.command_registry.register("close", ["close", "quit", "exit", "schließen", "beenden"])
        async def close_command(self, params):
            app = params.get("target", self.context.last_application)
            if not app:
                raise CommandError("Keine Anwendung zum Schließen angegeben.")
            await self.browser_manager._focus_application(app)
            pyautogui.hotkey("alt", "f4")
            self.logger.log_message(f"Anwendung '{app}' geschlossen.")

        @self.command_registry.register("maximize", ["maximize", "enlarge", "maximieren"])
        async def maximize_command(self, params):
            app = params.get("target", self.context.last_application)
            if not app:
                raise CommandError("Keine Anwendung zum Maximieren angegeben.")
            await self.browser_manager._focus_application(app)
            pyautogui.hotkey("win", "up")
            self.logger.log_message(f"Anwendung '{app}' maximiert.")

        @self.command_registry.register("write", ["write", "type", "input", "schreiben", "eingeben"])
        async def write_command(self, params):
            text = params.get("text")
            app = params.get("app", self.context.last_application)
            if not app:
                raise CommandError("Keine Anwendung zum Schreiben angegeben.")
            await self.browser_manager._focus_application(app)
            pyautogui.write(text)
            self.logger.log_message(f"Text '{text}' in {app} geschrieben.")

        @self.command_registry.register("save", ["save", "store", "speichern"])
        async def save_command(self, params):
            app = params.get("target", self.context.last_application)
            if not app:
                raise CommandError("Keine Anwendung zum Speichern angegeben.")
            await self.browser_manager._focus_application(app)
            pyautogui.hotkey("ctrl", "s")
            self.logger.log_message(f"Dokument in '{app}' gespeichert.")

        @self.command_registry.register("click", ["click", "klicken"])
        async def click_command(self, params):
            await self._perform_click()

        @self.command_registry.register("upload", ["upload", "send file", "hochladen", "datei senden"])
        async def upload_command(self, params):
            file_name = params.get("target")
            if not file_name:
                raise CommandError("Welche Datei soll hochgeladen werden? Sage z.B. 'Hochladen document.txt'.")
            if not self.config_manager.server_config:
                self.logger.log_message("Keine Serverdaten. Öffne die Einstellungen.")
                self._open_config_ui()
                if not self.config_manager.server_config:
                    raise ConfigError("Serverkonfiguration erforderlich")
            await self._upload_file(file_name)

        @self.command_registry.register("discord", ["discord", "message", "send", "nachricht", "senden"])
        async def discord_command(self, params):
            target = params.get("target")
            message = params.get("message")
            if not target or not message:
                raise CommandError("Sage mir, an wen und was ich senden soll, z.B. 'Sende an @user Hallo'.")
            if not self.config.discord_email or not self.config.discord_password:
                self.logger.log_message("Keine Discord-Zugangsdaten. Öffne die Einstellungen.")
                self._open_config_ui()
                if not self.config.discord_email or not self.config.discord_password:
                    raise ConfigError("Discord-Zugangsdaten erforderlich")
            await self._send_discord_message(target, message)

        @self.command_registry.register("winscp", ["winscp", "server", "sftp"])
        async def winscp_command(self, params):
            if not self.config_manager.server_config:
                self.logger.log_message("Keine Serverdaten. Öffne die Einstellungen.")
                self._open_config_ui()
                if not self.config_manager.server_config:
                    raise ConfigError("Serverkonfiguration erforderlich")
            await self._start_winscp()

        @self.command_registry.register("putty", ["putty", "ssh", "terminal"])
        async def putty_command(self, params):
            if not self.config_manager.server_config:
                self.logger.log_message("Keine Serverdaten. Öffne die Einstellungen.")
                self._open_config_ui()
                if not self.config_manager.server_config:
                    raise ConfigError("Serverkonfiguration erforderlich")
            await self._start_putty()

        @self.command_registry.register("task", ["task", "do", "execute", "aufgabe", "ausführen"])
        async def task_command(self, params):
            task = params.get("target")
            if not task:
                raise CommandError("Was soll ich tun? Sage z.B. 'Suche nach xAI'.")
            await self._execute_task(task)

        @self.command_registry.register("help", ["help", "commands", "hilfe", "befehle"])
        async def help_command(self, params):
            self.logger.log_message("Ich kann Folgendes tun:\n- Programme öffnen: 'Öffne Edge'\n- Suchen: 'Suche nach xAI'\n- Websites: 'Gehe zu https://check24.de'\n- Musik: 'Spiele Shape of You'\n- Hochladen: 'Hochladen document.txt'\n- Discord: 'Sende an @user Hallo'\n- Server: 'Starte WinSCP', 'Starte PuTTY'\n- Schreiben: 'Schreibe in Word Hallo'\n- Beenden: 'Beenden'\n- Hilfe: 'Hilfe'")

        @self.command_registry.register("exit", ["exit", "quit", "stop", "beenden", "stoppen"])
        async def exit_command(self, params):
            self.logger.log_message("Fahre herunter. Tschüss!")
            self.speech_manager.listening = False
            if self.browser_manager.driver:
                self.browser_manager.driver.quit()
            self.root.quit()

    @handle_errors
    def _open_config_ui(self):
        config_window = ctk.CTkToplevel(self.root)
        config_window.title("FaceBot Einstellungen")
        config_window.geometry("400x550")
        ctk.CTkLabel(config_window, text="Serverkonfiguration", font=("Arial", 12, "bold")).pack(pady=10)
        ctk.CTkLabel(config_window, text="Host (IP/Hostname):").pack()
        host_entry = ctk.CTkEntry(config_window, placeholder_text="Host eingeben")
        host_entry.pack()
        host_entry.insert(0, self.config_manager.server_config.get('host', '') if self.config_manager.server_config else '')
        ctk.CTkLabel(config_window, text="Benutzername:").pack()
        username_entry = ctk.CTkEntry(config_window, placeholder_text="Benutzername eingeben")
        username_entry.pack()
        username_entry.insert(0, self.config_manager.server_config.get('username', '') if self.config_manager.server_config else '')
        ctk.CTkLabel(config_window, text="Passwort (optional, wenn Schlüssel verwendet):").pack()
        password_entry = ctk.CTkEntry(config_window, show="*", placeholder_text="Passwort eingeben")
        password_entry.pack()
        password_entry.insert(0, self.config_manager.server_config.get('password', '') if self.config_manager.server_config else '')
        ctk.CTkLabel(config_window, text="Schlüsselpfad (.ppk, optional):").pack()
        key_path_entry = ctk.CTkEntry(config_window, placeholder_text="Schlüsselpfad eingeben")
        key_path_entry.pack()
        key_path_entry.insert(0, self.config_manager.server_config.get('key_path', '') if self.config_manager.server_config else '')
        ctk.CTkLabel(config_window, text="Discord-Konfiguration", font=("Arial", 12, "bold")).pack(pady=10)
        ctk.CTkLabel(config_window, text="E-Mail:").pack()
        discord_email_entry = ctk.CTkEntry(config_window, placeholder_text="E-Mail eingeben")
        discord_email_entry.pack()
        discord_email_entry.insert(0, self.config.discord_email)
        ctk.CTkLabel(config_window, text="Passwort:").pack()
        discord_password_entry = ctk.CTkEntry(config_window, show="*", placeholder_text="Passwort eingeben")
        discord_password_entry.pack()
        discord_password_entry.insert(0, self.config.discord_password)
        ctk.CTkLabel(config_window, text="Allgemeine Einstellungen", font=("Arial", 12, "bold")).pack(pady=10)
        speech_var = ctk.BooleanVar(value=self.config.enable_speech)
        ctk.CTkCheckBox(config_window, text="Sprachausgabe aktivieren", variable=speech_var).pack()
        listening_var = ctk.BooleanVar(value=self.config.enable_listening)
        ctk.CTkCheckBox(config_window, text="Spracherkennung aktivieren", variable=listening_var).pack()
        def save():
            try:
                host = self._sanitize_input(host_entry.get().strip())
                username = self._sanitize_input(username_entry.get().strip())
                password = self._sanitize_input(password_entry.get().strip())
                key_path = self._sanitize_input(key_path_entry.get().strip())
                if not host or not re.match(r"^[a-zA-Z0-9.-]+$", host):
                    raise ConfigError("Ungültiges Host-Format")
                if not username or not re.match(r"^[a-zA-Z0-9_-]+$", username):
                    raise ConfigError("Ungültiges Benutzername-Format")
                if key_path and not os.path.exists(key_path):
                    raise ConfigError("Ungültiger Schlüsselpfad")
                if host and username:
                    self.config_manager.server_config = {
                        "host": host,
                        "username": username,
                        "password": password,
                        "key_path": key_path
                    }
                self.config.discord_email = self._sanitize_input(discord_email_entry.get().strip())
                self.config.discord_password = self._sanitize_input(discord_password_entry.get().strip())
                if not self.config.discord_email or not re.match(r"[^@]+@[^@]+\.[^@]+", self.config.discord_email):
                    raise ConfigError("Ungültige Discord-E-Mail")
                self.config.enable_speech = speech_var.get()
                self.config.enable_listening = listening_var.get()
                self.config_manager._save_config()
                self.logger.log_message("Einstellungen gespeichert.")
                if not self.config.enable_listening and self.speech_manager.listening:
                    self.toggle_listening()
                config_window.destroy()
            except Exception as e:
                self.logger.log_message(f"Fehler beim Speichern der Einstellungen: {e}")
        ctk.CTkButton(config_window, text="Speichern", command=save).pack(pady=20)
        config_window.transient(self.root)
        config_window.grab_set()

    @handle_errors
    async def toggle_listening(self):
        if not self.config.enable_listening:
            raise SpeechError("Spracherkennung in den Einstellungen deaktiviert.")
        if not self.speech_manager.listening:
            try:
                sr.Microphone()
                self.speech_manager.listening = True
                self.logger.listen_button.configure(text="Mikrofon stoppen")
                asyncio.create_task(self.speech_manager.listen(self))
                self.logger.log_message("Höre zu...")
            except Exception as e:
                raise SpeechError(f"Mikrofonfehler: {e}. Verwende Texteingabe.")
        else:
            self.speech_manager.listening = False
            self.logger.listen_button.configure(text="Mikrofon")
            self.logger.log_message("Mikrofon ausgeschaltet.")

    @handle_errors
    def _sanitize_input(self, cmd: str) -> str:
        cmd = re.sub(r'[<>|;&$]', '', cmd)
        cmd = cmd.strip()
        if len(cmd) > 500:
            cmd = cmd[:500]
        return cmd

    @handle_errors
    async def process_command(self, event: Optional[tk.Event] = None, command: Optional[str] = None):
        cmd = command if command else self.logger.input_field.get().strip()
        if not command:
            self.logger.input_field.delete(0, "end")
        if not cmd:
            return
        cmd = self._sanitize_input(cmd)
        self.logger.log_message(f"Du: {cmd}")
        cmd = re.sub(r'^facebot[,]?[\s]*(hey\s)?', '', cmd, flags=re.IGNORECASE).strip().lower()
        intent, params = self.command_registry.parse(cmd, self.context)
        if not intent:
            raise CommandError(f"Ich habe '{cmd}' nicht verstanden. Sage z.B. 'Öffne Edge', 'Suche nach xAI', oder 'Hilfe'.")
        if intent in self.command_registry.commands:
            await self.command_registry.commands[intent]["func"](self, params)
        else:
            raise CommandError(f"Befehl '{cmd}' nicht verstanden. Sage z.B. 'Öffne Edge' oder 'Hilfe'.")

    @handle_errors
    async def _execute_task(self, task: str):
        self.logger.log_message(f"Arbeite an: '{task}'...")
        self.context.last_action = "task"
        task_lower = task.lower().strip()
        steps = re.split(r"\s+and\s+|,\s*", task_lower)
        for step in steps:
            step = step.strip()
            if not step:
                continue
            intent, params = self.command_registry.parse(step, self.context)
            if not intent:
                raise CommandError(f"Schritt '{step}' nicht verstanden. Sage z.B. 'Öffne Edge'.")
            if intent in self.command_registry.commands:
                await self.command_registry.commands[intent]["func"](self, params)
            else:
                raise CommandError(f"Schritt '{step}' nicht verstanden.")

    @handle_errors
    async def _perform_click(self):
        pyautogui.click()
        self.logger.log_message("Klick ausgeführt.")

    @handle_errors
    async def _start_winscp(self):
        if not os.path.exists(self.config.winscp_path):
            raise ConfigError(f"WinSCP nicht gefunden unter: '{self.config.winscp_path}'.")
        self.logger.log_message("Starte WinSCP und verbinde mit Server...")
        self.context.last_action = "winscp"
        cmd = [self.config.winscp_path, f"sftp://{self.config_manager.server_config['username']}@{self.config_manager.server_config['host']}"]
        if self.config_manager.server_config["key_path"]:
            cmd.append(f"/privatekey={self.config_manager.server_config['key_path']}")
        process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.root.after(10000, lambda: self._check_process_output(process, "WinSCP"))

    @handle_errors
    async def _start_putty(self):
        if not os.path.exists(self.config.putty_path):
            raise ConfigError(f"PuTTY nicht gefunden unter: '{self.config.putty_path}'.")
        self.logger.log_message("Starte PuTTY und verbinde mit Server...")
        self.context.last_action = "putty"
        cmd = [self.config.putty_path, "-ssh", f"{self.config_manager.server_config['username']}@{self.config_manager.server_config['host']}"]
        if self.config_manager.server_config["key_path"]:
            cmd.extend(["-i", self.config_manager.server_config["key_path"]])
        process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        self.root.after(10000, lambda: self._check_process_output(process, "PuTTY"))

    @handle_errors
    async def _check_process_output(self, process: subprocess.Popen, name: str):
        stdout, stderr = process.communicate(timeout=5)
        if stderr:
            raise ConfigError(f"Fehler bei {name}: {stderr.decode()}")
        self.logger.log_message(f"{name} erfolgreich gestartet!")

    @handle_errors
    async def _upload_file(self, file_name: str):
        file_name = self._sanitize_input(file_name)
        if os.path.isabs(file_name) and os.path.exists(file_name):
            file_path = file_name
        else:
            file_path = None
            for root, _, files in os.walk(self.config.base_search_dir, maxdepth=3):
                if file_name in files:
                    file_path = os.path.join(root, file_name)
                    break
        if not file_path or not os.path.exists(file_path):
            raise CommandError(f"Datei '{file_name}' nicht gefunden.")
        if not os.path.exists(self.config.winscp_path):
            raise ConfigError(f"WinSCP nicht gefunden unter: '{self.config.winscp_path}'.")
        self.logger.log_message(f"Lade '{file_path}' auf den Server hoch...")
        self.context.last_action = "upload"
        script_path = os.path.join(os.path.expanduser("~"), "upload_script.txt")
        if self.config_manager.server_config["key_path"]:
            script_content = (
                f'open sftp://{self.config_manager.server_config["username"]}@{self.config_manager.server_config["host"]} -privatekey="{self.config_manager.server_config["key_path"]}"\n'
                f'put "{file_path}" /root/\n'
                f'exit'
            )
        else:
            script_content = (
                f'open sftp://{self.config_manager.server_config["username"]}@{self.config_manager.server_config["host"]}\n'
                f'put "{file_path}" /root/\n'
                f'exit'
            )
        with open(script_path, "w", encoding="utf-8") as f:
            f.write(script_content)
        cmd = [self.config.winscp_path, "/script", script_path]
        if not self.config_manager.server_config["key_path"] and self.config_manager.server_config["password"]:
            cmd.extend(["/parameter", self._sanitize_input(self.config_manager.server_config["password"])])
        subprocess.run(cmd, shell=False, check=True)
        os.remove(script_path)
        self.logger.log_message("Datei erfolgreich hochgeladen!")

    @handle_errors
    async def _play_spotify_song(self, song_name: str):
        song_name = self._sanitize_input(song_name)
        with self.browser_manager:
            if not self.browser_manager.driver:
                raise BrowserError("Fehler beim Initialisieren des Browsers")
            self.logger.log_message(f"Spiele '{song_name}' auf Spotify...")
            self.context.last_action = "play"
            self.context.user_preferences["music"] += 1
            encoded_song = quote(song_name)
            search_url = self.config.spotify_search_url.format(encoded_song)
            self.browser_manager.driver.get(search_url)
            WebDriverWait(self.browser_manager.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            try:
                first_result = WebDriverWait(self.browser_manager.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, self.config.tracklist_css))
                )
                first_result.click()
                self.logger.log_message(f"'{song_name}' wird abgespielt! Überprüfe die Spotify-Anmeldung, falls es nicht startet.")
            except Exception as e:
                raise BrowserError(f"Fehler beim Abspielen des Songs: {e}. Bist du bei Spotify angemeldet?")

    @handle_errors
    async def _search_leta(self, search_term: str, browser: Optional[str]):
        search_term = self._sanitize_input(search_term)
        with self.browser_manager:
            if not self.browser_manager.driver:
                raise BrowserError("Fehler beim Initialisieren des Browsers")
            browser = browser or self.browser_manager.browser_name or "chrome"
            self.logger.log_message(f"Suche nach '{search_term}' auf Mullvad Leta (Brave) in {browser}...")
            self.context.last_action = "search"
            self.context.user_preferences["browser"] += 1
            encoded_term = quote(search_term)
            search_url = self.config.leta_search_url.format(encoded_term)
            await self.browser_manager._open_file_or_program(browser)
            await self.browser_manager._focus_application(browser)
            self.browser_manager.driver.get(search_url)
            WebDriverWait(self.browser_manager.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            self.logger.log_message(f"Suche nach '{search_term}' abgeschlossen!")

    @handle_errors
    async def _send_discord_message(self, target: str, message: str):
        target = self._sanitize_input(target)
        message = self._sanitize_input(message)
        with self.browser_manager:
            if not self.browser_manager.driver:
                raise BrowserError("Fehler beim Initialisieren des Browsers")
            self.logger.log_message(f"Sende Nachricht an '{target}' auf Discord...")
            self.context.last_action = "discord"
            self.browser_manager.driver.get(self.config.discord_login_url)
            WebDriverWait(self.browser_manager.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            try:
                email_field = WebDriverWait(self.browser_manager.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, self.config.discord_email_css))
                )
                password_field = self.browser_manager.driver.find_element(By.CSS_SELECTOR, self.config.discord_password_css)
                email_field.send_keys(self._sanitize_input(self.config.discord_email))
                password_field.send_keys(self._sanitize_input(self.config.discord_password))
                self.browser_manager.driver.find_element(By.CSS_SELECTOR, self.config.discord_submit_css).click()
                self.logger.log_message("Melde mich bei Discord an...")
                WebDriverWait(self.browser_manager.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, self.config.discord_message_css)))
            except Exception as e:
                raise BrowserError(f"Fehler beim Anmelden bei Discord: {e}. Melde dich manuell an oder überprüfe die Zugangsdaten.")
            try:
                message_field = WebDriverWait(self.browser_manager.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, self.config.discord_message_css))
                )
                message_field.send_keys(f"@{target} {message}")
                message_field.send_keys(Keys.RETURN)
                self.logger.log_message(f"Nachricht an '{target}' gesendet!")
            except Exception as e:
                raise BrowserError(f"Fehler beim Senden der Nachricht: {e}. Stelle sicher, dass der Discord-Kanal aktiv ist.")

    @handle_errors
    async def run(self):
        loop = asyncio.get_event_loop()
        loop.run_until_complete(self.root.mainloop())

if __name__ == "__main__":
    try:
        root = ctk.CTk()
        bot = FaceBot(root)
        asyncio.run(bot.run())
    except Exception as e:
        print(f"Fehler beim Starten des Bots: {e}")