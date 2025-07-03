import tkinter as tk
import tkinter.scrolledtext as scrolledtext
import tkinter.messagebox as messagebox
import os
import json
import logging
from dataclasses import dataclass
from typing import Optional, Dict, Tuple, List
from urllib.parse import quote, urlparse
import re
from cryptography.fernet import Fernet
import base64
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
import speech_recognition as sr
import pyaudio
import struct
import math
from gtts import gTTS
import pygame
import io
import threading

pygame.mixer.init()

@dataclass
class Config:
    winscp_path: str = r"C:\Program Files (x86)\WinSCP\WinSCP.exe"
    putty_path: str = r"C:\Program Files\PuTTY\putty.exe"
    base_search_dir: str = os.path.expandvars(r"%userprofile%")
    spotify_search_url: str = "https://open.spotify.com/search/{}"
    leta_search_url: str = "https://leta.mullvad.net/search?q={}&engine=brave"
    discord_login_url: str = "https://discord.com/login"
    tracklist_css: str = "div[data-testid='tracklist-row']"
    discord_email_css: str = "input[name='email']"
    discord_password_css: str = "input[name='password']"
    discord_submit_css: str = "button[type='submit']"
    discord_message_css: str = "div[role='textbox']"
    config_file: str = os.path.join(os.path.expanduser("~"), ".facebot_config.json")
    encryption_key_file: str = os.path.join(os.path.expanduser("~"), ".facebot_key")
    enable_speech: bool = False
    speech_language: str = "de"
    enable_listening: bool = True

@dataclass
class Context:
    last_application: Optional[str] = None
    last_action: Optional[str] = None
    user_preferences: Dict[str, int] = None

    def __post_init__(self):
        self.user_preferences = self.user_preferences or {"music": 0, "browser": 0, "document": 0}

class ConfigManager:
    def __init__(self, config: Config):
        self.config = config
        self.fernet = None
        self.server_config = None
        self._setup_encryption()
        self._load_config()

    def _setup_encryption(self):
        try:
            if os.path.exists(self.config.encryption_key_file):
                with open(self.config.encryption_key_file, 'rb') as f:
                    key = f.read()
            else:
                key = Fernet.generate_key()
                with open(self.config.encryption_key_file, 'wb') as f:
                    f.write(key)
            self.fernet = Fernet(key)
        except Exception as e:
            logging.getLogger("FaceBot").error(f"Error setting up encryption: {e}")
            self.fernet = None

    def _encrypt_data(self, data: str) -> str:
        if not self.fernet or not data:
            return data
        return base64.b64encode(self.fernet.encrypt(data.encode())).decode()

    def _decrypt_data(self, data: str) -> str:
        if not self.fernet or not data:
            return data
        try:
            return self.fernet.decrypt(base64.b64decode(data)).decode()
        except Exception:
            return data

    def _load_config(self):
        try:
            if os.path.exists(self.config.config_file):
                with open(self.config.config_file, 'r') as f:
                    config_data = json.load(f)
                self.server_config = config_data.get('server_config', {})
                self.server_config['password'] = self._decrypt_data(self.server_config.get('password', ''))
                self.config.discord_email = self._decrypt_data(config_data.get('discord_email', ''))
                self.config.discord_password = self._decrypt_data(config_data.get('discord_password', ''))
                self.config.enable_speech = config_data.get('speech_enabled', self.config.enable_speech)
                self.config.enable_listening = config_data.get('enable_listening', self.config.enable_listening)
        except Exception as e:
            logging.getLogger("FaceBot").error(f"Error loading configuration: {e}")

    def _save_config(self):
        try:
            config_data = {
                'server_config': self.server_config or {},
                'discord_email': self._encrypt_data(self.config.discord_email),
                'discord_password': self._encrypt_data(self.config.discord_password),
                'speech_enabled': self.config.enable_speech,
                'enable_listening': self.config.enable_listening
            }
            config_data['server_config']['password'] = self._encrypt_data(self.server_config.get('password', '') if self.server_config else '')
            with open(self.config.config_file, 'w') as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            logging.getLogger("FaceBot").error(f"Error saving configuration: {e}")

class BrowserManager:
    def __init__(self):
        self.driver = None
        self.browser_name = None

    def _get_default_browser(self) -> str:
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice") as key:
                prog_id = winreg.QueryValueEx(key, "ProgId")[0]
            if "Chrome" in prog_id:
                return "chrome"
            elif "Firefox" in prog_id:
                return "firefox"
            elif "Edge" in prog_id or "IE" in prog_id:
                return "edge"
            elif "Opera" in prog_id:
                return "opera"
            return "chrome"
        except Exception:
            return "chrome"

    def initialize(self):
        self.browser_name = self._get_default_browser()
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
        except Exception as e:
            logging.getLogger("FaceBot").error(f"Error starting {self.browser_name}: {e}. Trying Chrome...")
            try:
                service = ChromeService(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service)
                self.driver.maximize_window()
                self.browser_name = "chrome"
            except Exception as e2:
                logging.getLogger("FaceBot").error(f"Error starting Chrome: {e2}. Browser functions disabled.")
                self.driver = None
                self.browser_name = "chrome"

    def navigate_to_url(self, url: str, browser: Optional[str], logger, context: Context):
        try:
            parsed_url = urlparse(url)
            if not parsed_url.scheme or not parsed_url.netloc:
                logger.log_message(f"Invalid URL '{url}'. Provide a valid URL.")
                return
            if not self.driver:
                logger.log_message("No browser available. Restarting browser.")
                self.initialize()
                if not self.driver:
                    return
            browser = browser or self.browser_name or "chrome"
            logger.log_message(f"Navigating to '{url}' in {browser}...")
            context.last_action = "goto"
            context.user_preferences["browser"] += 1
            self._focus_application(browser, logger)
            self.driver.get(url)
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            logger.log_message(f"Navigation to '{url}' completed!")
        except Exception as e:
            logger.log_message(f"Error navigating to '{url}': {e}. Check URL and browser.")

    def _focus_application(self, app_name: str, logger):
        try:
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
                logger.log_message(f"Application '{app_name}' focused.")
                return True
            logger.log_message(f"Application '{app_name}' not found. Starting it...")
            self._open_file_or_program(app_name, logger)
            return True
        except Exception as e:
            logger.log_message(f"Error focusing '{app_name}': {e}. Ensure application installed.")
            return False

    def _open_file_or_program(self, target: str, logger):
        try:
            logger.log_message(f"Opening '{target}'...")
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
            target_lower = target.lower().strip()
            executable = program_map.get(target_lower)
            if executable:
                program_path = shutil.which(executable)
                if program_path:
                    subprocess.Popen([program_path], shell=False)
                    logger.log_message(f"Program '{target}' started!")
                    return
            program_path = shutil.which(target)
            if program_path:
                subprocess.Popen([program_path], shell=False)
                logger.log_message(f"Program '{target}' started!")
                return
            if os.path.isabs(target) and os.path.exists(target):
                os.startfile(target)
                logger.log_message(f"'{target}' opened!")
                return
            for root, _, files in os.walk(Config().base_search_dir):
                if target in files:
                    file_path = os.path.join(root, target)
                    os.startfile(file_path)
                    logger.log_message(f"File '{file_path}' opened!")
                    return
            suggestions = self._suggest_alternatives(target, logger)
            if suggestions:
                suggestion_text = "\n".join([f"- {name}: {reason}" for name, reason in suggestions])
                logger.log_message(f"'{target}' not found. Did you mean:\n{suggestion_text}\nSay e.g., 'Open {suggestions[0][0]}'.")
            else:
                logger.log_message(f"'{target}' not found. Provide valid path or program.")
        except Exception as e:
            logger.log_message(f"Error opening '{target}': {e}. Check path or program.")

    def _suggest_alternatives(self, target: str, logger) -> List[Tuple[str, str]]:
        try:
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
                    suggestions.append((name, 90, "similar name"))
            if "." in target:
                mime_type, _ = mimetypes.guess_type(target)
                if mime_type:
                    if mime_type.startswith("text"):
                        suggestions.append(("notepad", 90, "text file"))
                        if shutil.which("winword.exe"):
                            suggestions.append(("word", 85, "text file"))
                    elif mime_type.startswith("image"):
                        suggestions.append(("msedge.exe", 85, "image file"))
                    elif mime_type.startswith("application/vnd"):
                        if shutil.which("excel.exe"):
                            suggestions.append(("excel", 85, "spreadsheet"))
            for name, exe in program_map.items():
                if shutil.which(exe) and name not in [s[0] for s in suggestions]:
                    suggestions.append((name, 80, "available on system"))
            suggestions = sorted(suggestions, key=lambda x: x[1], reverse=True)[:3]
            return [(name, reason) for name, _, reason in suggestions]
        except Exception as e:
            logger.log_message(f"Error generating suggestions: {e}")
            return []

class SpeechManager:
    def __init__(self, config: Config):
        self.config = config
        self.recognizer = sr.Recognizer()
        self.listening = False
        self.audio_thread = None

    def _speak(self, text: str):
        if not self.config.enable_speech:
            return
        try:
            def play_audio():
                tts = gTTS(text=text, lang=self.config.speech_language)
                audio_file = io.BytesIO()
                tts.write_to_fp(audio_file)
                audio_file.seek(0)
                pygame.mixer.music.load(audio_file)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    time.sleep(0.1)
            threading.Thread(target=play_audio, daemon=True).start()
        except Exception as e:
            logging.getLogger("FaceBot").error(f"Error in speech output: {e}")

    def _update_audio_indicator(self, stream: pyaudio.Stream, canvas: tk.Canvas, bars: list):
        CHUNK = 1024
        while self.listening:
            try:
                data = stream.read(CHUNK, exception_on_overflow=False)
                rms = math.sqrt(abs(sum([(x / 32768) ** 2 for x in struct.unpack(f"{CHUNK * 2}h", data)]) / CHUNK))
                height = min(max(int(rms * 50), 5), 25)
                for i, bar in enumerate(bars):
                    canvas.coords(bar, 10 + i * 20, 25, 25 + i * 20, 25 - height + (i % 2) * 5)
                canvas.update()
            except Exception:
                pass

    def listen(self, bot: 'FaceBot'):
        stream = None
        error_count = 0
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
                self.recognizer.dynamic_energy_threshold = True
                self.recognizer.energy_threshold = 300
                stream = pyaudio.PyAudio().open(format=pyaudio.paInt16, channels=1, rate=44100, input=True, frames_per_buffer=1024)
                threading.Thread(target=self._update_audio_indicator, args=(stream, bot.indicator_canvas, bot.bars), daemon=True).start()
                bot.log_message("Speech recognition active. Speak clearly, e.g., 'Open Edge'.")
                while self.listening:
                    try:
                        audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=15)
                        audio_duration = len(audio.frame_data) / (audio.sample_rate * audio.sample_width)
                        if audio_duration < 0.5:
                            error_count = 0
                            continue
                        command = self.recognizer.recognize_google(audio, language="en-US")
                        bot.log_message(f"You said: {command}")
                        bot.root.after(0, bot.process_command, None, command)
                        error_count = 0
                    except sr.WaitTimeoutError:
                        error_count = 0
                    except sr.UnknownValueError:
                        error_count += 1
                        if error_count >= 3:
                            bot.log_message("Multiple unclear inputs. Speak louder or check microphone.")
                            error_count = 0
                    except sr.RequestError as e:
                        bot.log_message(f"Speech recognition error: {e}. Check internet.")
                        error_count = 0
                    except Exception as e:
                        bot.log_message(f"Unknown speech recognition error: {e}. Try again.")
                        error_count = 0
        except Exception as e:
            bot.log_message(f"Error starting speech recognition: {e}. Speech disabled.")
            self.listening = False
            bot.listen_button.config(text="Microphone")
        finally:
            if stream:
                stream.stop_stream()
                stream.close()
                pyaudio.PyAudio().terminate()

class UIManager:
    def __init__(self, root: tk.Tk, bot: 'FaceBot'):
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

    def _setup_ui(self):
        self.root.title("FaceBot")
        self.chat_area = scrolledtext.ScrolledText(self.root, wrap=tk.WORD, width=60, height=20, state='disabled')
        self.chat_area.pack(padx=10, pady=10)
        input_frame = tk.Frame(self.root)
        input_frame.pack(padx=10, pady=5, fill=tk.X)
        self.input_field = tk.Entry(input_frame)
        self.input_field.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.input_field.bind("<Return>", self.bot.process_command)
        self.send_button = tk.Button(input_frame, text="Send", command=self.bot.process_command)
        self.send_button.pack(side=tk.RIGHT, padx=5)
        self.listen_button = tk.Button(input_frame, text="Microphone", command=self.bot.toggle_listening)
        self.listen_button.pack(side=tk.RIGHT)
        self.config_button = tk.Button(input_frame, text="Settings", command=self.bot._open_config_ui)
        self.config_button.pack(side=tk.RIGHT, padx=5)
        self.indicator_canvas = tk.Canvas(self.root, width=100, height=30, bg='black')
        self.indicator_canvas.pack(pady=5)
        for i in range(5):
            x = 10 + i * 20
            bar = self.indicator_canvas.create_rectangle(x, 25, x + 15, 25, fill='green')
            self.bars.append(bar)

    def log_message(self, message: str):
        def update_gui():
            self.chat_area.configure(state='normal')
            self.chat_area.insert(tk.END, f"{message}\n")
            self.chat_area.configure(state='disabled')
            self.chat_area.see(tk.END)
            self.root.update()
        self.root.after(0, update_gui)
        logging.getLogger("FaceBot").info(message)
        self.bot.speech_manager._speak(message)

class CommandParser:
    def parse(self, command: str, context: Context) -> Tuple[Optional[str], Dict[str, str]]:
        try:
            intent = None
            params = {}
            patterns = {
                "open": r"^(?:open|start|launch)\s+(.+)$",
                "play": r"^(?:play|music)\s+(.+)$",
                "search": r"^(?:search|google|find)\s*(?:for)?\s+(.+?)(?:\s+in\s+([a-zA-Z\s]+))?\s*$",
                "goto": r"^(?:go\s+to|goto|navigate\s+to|gehe\s+zu)\s+(https?://[^\s]+)(?:\s+in\s+([a-zA-Z\s]+))?\s*$",
                "close": r"^(?:close|quit|exit)\s+(.+)$",
                "maximize": r"^(?:maximize|enlarge)\s+(.+)$",
                "write": r"^(?:write|type|input)\s+(.+?)(?:\s+in\s+(word|excel|notepad))?$",
                "save": r"^(?:save|store)\s+(.+)$",
                "click": r"^(?:click)$",
                "upload": r"^(?:upload|send\s+file)\s+(.+)$",
                "discord": r"^(?:discord|message|send)\s+to\s+(.+?)\s+(.+)$",
                "winscp": r"^(?:winscp|server|sftp)$",
                "putty": r"^(?:putty|ssh|terminal)$",
                "task": r"^(?:task|do|execute)\s+(.+)$",
                "help": r"^(?:help|commands)$",
                "exit": r"^(?:exit|quit|stop)$"
            }
            command_lower = command.lower().strip()
            for intent_name, pattern in patterns.items():
                match = re.match(pattern, command_lower)
                if match:
                    intent = intent_name
                    if intent == "search":
                        params["search_term"] = match.group(1).strip()
                        params["browser"] = match.group(2).strip() if len(match.groups()) > 1 and match.group(2) else None
                    elif intent == "goto":
                        params["url"] = match.group(1).strip()
                        params["browser"] = match.group(2).strip() if len(match.groups()) > 1 and match.group(2) else None
                    elif intent in ["open", "play", "close", "maximize", "save", "upload", "task"]:
                        params["target"] = match.group(1).strip()
                    elif intent == "write":
                        params["text"] = match.group(1).strip()
                        params["app"] = match.group(2) if len(match.groups()) > 1 else None
                    elif intent == "discord":
                        params["target"] = match.group(1).strip()
                        params["message"] = match.group(2).strip()
                    break
            if not intent:
                intent_keywords = {
                    "open": ["open", "start", "launch"],
                    "play": ["play", "music"],
                    "search": ["search", "google", "find"],
                    "goto": ["go to", "goto", "navigate to", "gehe zu"],
                    "close": ["close", "quit", "exit"],
                    "maximize": ["maximize", "enlarge"],
                    "write": ["write", "type", "input"],
                    "save": ["save", "store"],
                    "click": ["click"],
                    "upload": ["upload", "send file"],
                    "discord": ["discord", "message", "send"],
                    "winscp": ["winscp", "server", "sftp"],
                    "putty": ["putty", "ssh", "terminal"],
                    "task": ["task", "do", "execute"],
                    "help": ["help", "commands"],
                    "exit": ["exit", "quit", "stop"]
                }
                for token in command_lower.split():
                    for key, keywords in intent_keywords.items():
                        if token in keywords or any(kw in command_lower for kw in keywords):
                            intent = key
                            params["target"] = command_lower.replace(token, "").replace("for", "").strip()
                            break
                    if intent:
                        break
            if intent in ["play", "search", "goto"]:
                context.user_preferences["music" if intent == "play" else "browser"] += 1
            elif intent in ["write", "save"] and params.get("app") in ["word", "excel"]:
                context.user_preferences["document"] += 1
            return intent, params
        except Exception as e:
            logging.getLogger("FaceBot").error(f"Error parsing command: {e}")
            return None, {}

class FaceBot:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.config = Config()
        self.context = Context()
        self.logger = UIManager(root, self)
        self.browser_manager = BrowserManager()
        self.speech_manager = SpeechManager(self.config)
        self.config_manager = ConfigManager(self.config)
        self.command_parser = CommandParser()
        self._setup_logger()
        self.browser_manager.initialize()
        self.context.last_application = self.browser_manager.browser_name
        self.logger.log_message(f"Okay, I'm ready! Using browser: {self.browser_manager.browser_name or 'Unknown'}. Say e.g., 'Open Edge', 'Search for xAI'.")

    def _setup_logger(self):
        logger = logging.getLogger("FaceBot")
        logger.setLevel(logging.DEBUG)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(handler)

    def _open_config_ui(self):
        try:
            config_window = tk.Toplevel(self.root)
            config_window.title("FaceBot Settings")
            config_window.geometry("400x550")
            tk.Label(config_window, text="Server Configuration", font=("Arial", 12, "bold")).pack(pady=10)
            tk.Label(config_window, text="Host (IP/Hostname):").pack()
            host_entry = tk.Entry(config_window)
            host_entry.pack()
            host_entry.insert(0, self.config_manager.server_config.get('host', '') if self.config_manager.server_config else '')
            tk.Label(config_window, text="Username:").pack()
            username_entry = tk.Entry(config_window)
            username_entry.pack()
            username_entry.insert(0, self.config_manager.server_config.get('username', '') if self.config_manager.server_config else '')
            tk.Label(config_window, text="Password (optional if key is used):").pack()
            password_entry = tk.Entry(config_window, show="*")
            password_entry.pack()
            password_entry.insert(0, self.config_manager.server_config.get('password', '') if self.config_manager.server_config else '')
            tk.Label(config_window, text="Key Path (.ppk, optional):").pack()
            key_path_entry = tk.Entry(config_window)
            key_path_entry.pack()
            key_path_entry.insert(0, self.config_manager.server_config.get('key_path', '') if self.config_manager.server_config else '')
            tk.Label(config_window, text="Discord Configuration", font=("Arial", 12, "bold")).pack(pady=10)
            tk.Label(config_window, text="Email:").pack()
            discord_email_entry = tk.Entry(config_window)
            discord_email_entry.pack()
            discord_email_entry.insert(0, self.config.discord_email)
            tk.Label(config_window, text="Password:").pack()
            discord_password_entry = tk.Entry(config_window, show="*")
            discord_password_entry.pack()
            discord_password_entry.insert(0, self.config.discord_password)
            tk.Label(config_window, text="General Settings", font=("Arial", 12, "bold")).pack(pady=10)
            speech_var = tk.BooleanVar(value=self.config.enable_speech)
            tk.Checkbutton(config_window, text="Enable Speech Output", variable=speech_var).pack()
            listening_var = tk.BooleanVar(value=self.config.enable_listening)
            tk.Checkbutton(config_window, text="Enable Speech Recognition", variable=listening_var).pack()
            def save():
                host = host_entry.get().strip()
                username = username_entry.get().strip()
                password = password_entry.get().strip()
                key_path = key_path_entry.get().strip()
                if host and username:
                    self.config_manager.server_config = {
                        "host": host,
                        "username": username,
                        "password": password,
                        "key_path": key_path
                    }
                self.config.discord_email = discord_email_entry.get().strip()
                self.config.discord_password = discord_password_entry.get().strip()
                self.config.enable_speech = speech_var.get()
                self.config.enable_listening = listening_var.get()
                self.config_manager._save_config()
                self.logger.log_message("Settings saved.")
                if not self.config.enable_listening and self.speech_manager.listening:
                    self.toggle_listening()
                config_window.destroy()
            tk.Button(config_window, text="Save", command=save).pack(pady=20)
            config_window.transient(self.root)
            config_window.grab_set()
        except Exception as e:
            self.logger.log_message(f"Error opening settings: {e}")

    def toggle_listening(self):
        try:
            if not self.config.enable_listening:
                self.logger.log_message("Speech recognition disabled in settings.")
                return
            if not self.speech_manager.listening:
                try:
                    sr.Microphone()
                    self.speech_manager.listening = True
                    self.logger.listen_button.config(text="Stop Microphone")
                    self.speech_manager.audio_thread = threading.Thread(target=self.speech_manager.listen, args=(self,), daemon=True)
                    self.speech_manager.audio_thread.start()
                    self.logger.log_message("Hearing...")
                except Exception as e:
                    self.logger.log_message(f"Microphone error: {e}. Use text input.")
            else:
                self.speech_manager.listening = False
                self.logger.listen_button.config(text="Microphone")
                if self.speech_manager.audio_thread:
                    self.speech_manager.audio_thread.join(timeout=1)
                    self.speech_manager.audio_thread = None
                for bar in self.logger.bars:
                    self.logger.indicator_canvas.coords(bar, 10 + self.logger.bars.index(bar) * 20, 25, 25 + self.logger.bars.index(bar) * 20, 25)
                self.logger.log_message("Microphone turned off.")
        except Exception as e:
            self.logger.log_message(f"Error toggling microphone: {e}. Use text input.")

    def _sanitize_input(self, cmd: str) -> str:
        cmd = re.sub(r'[<>|;&$]', '', cmd)
        cmd = cmd.strip()
        if len(cmd) > 500:
            cmd = cmd[:500]
        return cmd

    def process_command(self, event: Optional[tk.Event] = None, command: Optional[str] = None):
        try:
            cmd = command if command else self.logger.input_field.get().strip()
            if not command:
                self.logger.input_field.delete(0, tk.END)
            if not cmd:
                return
            cmd = self._sanitize_input(cmd)
            self.logger.log_message(f"You: {cmd}")
            cmd = re.sub(r'^facebot[,]?[\s]*(hey\s)?', '', cmd, flags=re.IGNORECASE).strip().lower()
            intent, params = self.command_parser.parse(cmd, self.context)
            if not intent:
                self.logger.log_message(f"I didn't understand '{cmd}'. Say e.g., 'Open Edge', 'Search for xAI', or 'Help'.")
                return
            if intent == "exit":
                self.logger.log_message("Shutting down. Bye!")
                self.speech_manager.listening = False
                if self.browser_manager.driver:
                    self.browser_manager.driver.quit()
                self.root.quit()
            elif intent == "help":
                self.logger.log_message("I can do:\n- Open programs: 'Open Edge'\n- Search: 'Search for xAI'\n- Websites: 'Go to https://check24.de'\n- Music: 'Play Shape of You'\n- Upload: 'Upload document.txt'\n- Discord: 'Send to @user Hello'\n- Server: 'Start WinSCP', 'Start PuTTY'\n- Write: 'Write in Word Hello'\n- Exit: 'Exit'\n- Help: 'Help'")
            elif intent == "click":
                self._perform_click()
            elif intent == "winscp":
                if not self.config_manager.server_config:
                    self.logger.log_message("No server data. Open settings.")
                    self._open_config_ui()
                    if not self.config_manager.server_config:
                        return
                self._start_winscp()
            elif intent == "putty":
                if not self.config_manager.server_config:
                    self.logger.log_message("No server data. Open settings.")
                    self._open_config_ui()
                    if not self.config_manager.server_config:
                        return
                self._start_putty()
            elif intent == "upload":
                file_name = params.get("target")
                if not file_name:
                    self.logger.log_message("Which file to upload? Say e.g., 'Upload document.txt'.")
                    return
                if not self.config_manager.server_config:
                    self.logger.log_message("No server data. Open settings.")
                    self._open_config_ui()
                    if not self.config_manager.server_config:
                        return
                self._upload_file(file_name)
            elif intent == "discord":
                target = params.get("target")
                message = params.get("message")
                if not target or not message:
                    self.logger.log_message("Tell me who to send to and what, e.g., 'Send to @user Hello'.")
                    return
                if not self.config.discord_email or not self.config.discord_password:
                    self.logger.log_message("No Discord credentials. Open settings.")
                    self._open_config_ui()
                    if not self.config.discord_email or not self.config.discord_password:
                        return
                self._send_discord_message(target, message)
            elif intent == "play":
                song_name = params.get("target")
                if not song_name:
                    self.logger.log_message("Which song to play? Say e.g., 'Play Shape of You'.")
                    return
                self._play_spotify_song(song_name)
            elif intent == "search":
                search_term = params.get("search_term")
                browser = params.get("browser", self.browser_manager.browser_name)
                if not search_term:
                    self.logger.log_message("What to search for? Say e.g., 'Search for xAI'.")
                    return
                self._search_leta(search_term, browser)
            elif intent == "goto":
                url = params.get("url")
                browser = params.get("browser", self.browser_manager.browser_name)
                if not url:
                    self.logger.log_message("What website to go to? Say e.g., 'Go to https://check24.de'.")
                    return
                self.browser_manager.navigate_to_url(url, browser, self.logger, self.context)
            elif intent == "open":
                target = params.get("target")
                if not target:
                    self.logger.log_message("What to open? Say e.g., 'Open Edge'.")
                    return
                self.browser_manager._open_file_or_program(target, self.logger)
            elif intent == "task":
                task = params.get("target", cmd)
                if not task:
                    self.logger.log_message("What to do? Say e.g., 'Search for xAI'.")
                    return
                self._execute_task(task)
            else:
                self.logger.log_message(f"Command '{cmd}' not understood. Say e.g., 'Open Edge' or 'Help'.")
        except Exception as e:
            self.logger.log_message(f"Error processing '{cmd}': {e}. Try another command.")

    def _execute_task(self, task: str):
        try:
            self.logger.log_message(f"Working on: '{task}'...")
            self.context.last_action = "task"
            task_lower = task.lower().strip()
            steps = re.split(r"\s+and\s+|,\s*", task_lower)
            for step in steps:
                step = step.strip()
                if not step:
                    continue
                intent, params = self.command_parser.parse(step, self.context)
                if not intent:
                    self.logger.log_message(f"Step '{step}' not understood. Say e.g., 'Open Edge'.")
                    continue
                if intent == "open" and "tab" in step:
                    browser = params.get("target", self.browser_manager.browser_name)
                    self.browser_manager._focus_application(browser, self.logger)
                    self.root.after(100, lambda: subprocess.run(["start", ""], shell=True))
                    self.logger.log_message(f"New tab opened in {browser}!")
                elif intent == "search":
                    search_term = params.get("search_term", step.replace("search", "").replace("for", "").strip())
                    browser = params.get("browser", self.browser_manager.browser_name)
                    if not search_term:
                        self.logger.log_message("No search term provided.")
                        continue
                    self._search_leta(search_term, browser)
                elif intent == "goto":
                    url = params.get("url", step.replace("go to", "").replace("goto", "").replace("navigate to", "").strip())
                    browser = params.get("browser", self.browser_manager.browser_name)
                    if not url:
                        self.logger.log_message("No URL provided.")
                        continue
                    self.browser_manager.navigate_to_url(url, browser, self.logger, self.context)
                elif intent == "close":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.logger.log_message("No application specified to close.")
                        continue
                    self.browser_manager._focus_application(app, self.logger)
                    self.root.after(100, lambda: subprocess.run(["taskkill", "/IM", app + ".exe", "/F"], shell=True, capture_output=True))
                    self.logger.log_message(f"Application '{app}' closed.")
                elif intent == "maximize":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.logger.log_message("No application specified to maximize.")
                        continue
                    self.browser_manager._focus_application(app, self.logger)
                    self.logger.log_message(f"Application '{app}' maximized.")
                elif intent == "open":
                    program = params.get("target")
                    if not program:
                        self.logger.log_message("What to open?")
                        continue
                    self.browser_manager._open_file_or_program(program, self.logger)
                elif intent == "write":
                    text = params.get("text", step.replace("write", "").replace("type", "").strip())
                    app = params.get("app", self.context.last_application)
                    if not app:
                        self.logger.log_message("No application specified to write.")
                        continue
                    self.browser_manager._focus_application(app, self.logger)
                    self.root.after(100, lambda: subprocess.run(["powershell", "-Command", f"Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('{text}')"], shell=True))
                    self.logger.log_message(f"Text '{text}' written in {app}.")
                elif intent == "save":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.logger.log_message("No application specified to save.")
                        continue
                    self.browser_manager._focus_application(app, self.logger)
                    self.root.after(100, lambda: subprocess.run(["powershell", "-Command", "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('^s')"], shell=True))
                    self.logger.log_message(f"Document saved in '{app}'.")
                else:
                    self.logger.log_message(f"Step '{step}' not understood.")
        except Exception as e:
            self.logger.log_message(f"Error executing '{task}': {e}.")

    def _perform_click(self):
        try:
            subprocess.run(["powershell", "-Command", "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('{LEFT}')"], shell=True)
            self.logger.log_message("Click performed.")
        except Exception as e:
            self.logger.log_message(f"Error performing click: {e}.")

    def _start_winscp(self):
        try:
            if not os.path.exists(self.config.winscp_path):
                self.logger.log_message(f"WinSCP not found at '{self.config.winscp_path}'.")
                return
            self.logger.log_message("Starting WinSCP and connecting to server...")
            self.context.last_action = "winscp"
            cmd = [self.config.winscp_path, f"sftp://{self.config_manager.server_config['username']}@{self.config_manager.server_config['host']}"]
            if self.config_manager.server_config["key_path"]:
                cmd.append(f"/privatekey={self.config_manager.server_config['key_path']}")
            process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.root.after(10000, lambda: self._check_process_output(process, "WinSCP"))
        except Exception as e:
            self.logger.log_message(f"Error starting WinSCP: {e}.")

    def _start_putty(self):
        try:
            if not os.path.exists(self.config.putty_path):
                self.logger.log_message(f"PuTTY not found at '{self.config.putty_path}'.")
                return
            self.logger.log_message("Starting PuTTY and connecting to server...")
            self.context.last_action = "putty"
            cmd = [self.config.putty_path, "-ssh", f"{self.config_manager.server_config['username']}@{self.config_manager.server_config['host']}"]
            if self.config_manager.server_config["key_path"]:
                cmd.extend(["-i", self.config_manager.server_config["key_path"]])
            process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.root.after(10000, lambda: self._check_process_output(process, "PuTTY"))
        except Exception as e:
            self.logger.log_message(f"Error starting PuTTY: {e}.")

    def _check_process_output(self, process: subprocess.Popen, name: str):
        try:
            stdout, stderr = process.communicate(timeout=5)
            if stderr:
                self.logger.log_message(f"Error with {name}: {stderr.decode()}")
            else:
                self.logger.log_message(f"{name} started successfully!")
        except Exception as e:
            self.logger.log_message(f"Error checking {name}: {e}")

    def _upload_file(self, file_name: str):
        try:
            if os.path.isabs(file_name) and os.path.exists(file_name):
                file_path = file_name
            else:
                file_path = None
                for root, _, files in os.walk(self.config.base_search_dir):
                    if file_name in files:
                        file_path = os.path.join(root, file_name)
                        break
            if not file_path or not os.path.exists(file_path):
                self.logger.log_message(f"File '{file_name}' not found.")
                return
            if not os.path.exists(self.config.winscp_path):
                self.logger.log_message(f"WinSCP not found at '{self.config.winscp_path}'.")
                return
            self.logger.log_message(f"Uploading '{file_path}' to server...")
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
                cmd.extend(["/parameter", self.config_manager.server_config["password"]])
            subprocess.run(cmd, shell=False, check=True)
            os.remove(script_path)
            self.logger.log_message("File uploaded successfully!")
        except Exception as e:
            self.logger.log_message(f"Error uploading: {e}.")

    def _play_spotify_song(self, song_name: str):
        try:
            if not self.browser_manager.driver:
                self.logger.log_message("No browser available. Restarting browser.")
                self.browser_manager.initialize()
                if not self.browser_manager.driver:
                    return
            self.logger.log_message(f"Playing '{song_name}' on Spotify...")
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
                self.logger.log_message(f"'{song_name}' is playing! Check Spotify login if it doesn't start.")
            except Exception as e:
                self.logger.log_message(f"Error playing: {e}. Are you logged into Spotify?")
        except Exception as e:
            self.logger.log_message(f"Error opening Spotify: {e}. Check internet and browser.")

    def _search_leta(self, search_term: str, browser: Optional[str]):
        try:
            if not self.browser_manager.driver:
                self.logger.log_message("No browser available. Restarting browser.")
                self.browser_manager.initialize()
                if not self.browser_manager.driver:
                    return
            browser = browser or self.browser_manager.browser_name or "chrome"
            self.logger.log_message(f"Searching for '{search_term}' on Mullvad Leta (Brave) in {browser}...")
            self.context.last_action = "search"
            self.context.user_preferences["browser"] += 1
            encoded_term = quote(search_term)
            search_url = self.config.leta_search_url.format(encoded_term)
            self.browser_manager._open_file_or_program(browser, self.logger)
            self.browser_manager._focus_application(browser, self.logger)
            self.browser_manager.driver.get(search_url)
            WebDriverWait(self.browser_manager.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            self.logger.log_message(f"Search for '{search_term}' completed!")
        except Exception as e:
            self.logger.log_message(f"Error searching on Leta: {e}. Check internet and browser.")

    def _send_discord_message(self, target: str, message: str):
        try:
            if not self.browser_manager.driver:
                self.logger.log_message("No browser available. Restarting browser.")
                self.browser_manager.initialize()
                if not self.browser_manager.driver:
                    return
            self.logger.log_message(f"Sending message to '{target}' on Discord...")
            self.context.last_action = "discord"
            self.browser_manager.driver.get(self.config.discord_login_url)
            WebDriverWait(self.browser_manager.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            try:
                email_field = WebDriverWait(self.browser_manager.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, self.config.discord_email_css))
                )
                password_field = self.browser_manager.driver.find_element(By.CSS_SELECTOR, self.config.discord_password_css)
                email_field.send_keys(self.config.discord_email)
                password_field.send_keys(self.config.discord_password)
                self.browser_manager.driver.find_element(By.CSS_SELECTOR, self.config.discord_submit_css).click()
                self.logger.log_message("Logging into Discord...")
                WebDriverWait(self.browser_manager.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, self.config.discord_message_css)))
            except Exception as e:
                self.logger.log_message(f"Error logging into Discord: {e}. Log in manually or check credentials.")
                return
            try:
                message_field = WebDriverWait(self.browser_manager.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, self.config.discord_message_css))
                )
                message_field.send_keys(f"@{target} {message}")
                message_field.send_keys(Keys.RETURN)
                self.logger.log_message(f"Message sent to '{target}'!")
            except Exception as e:
                self.logger.log_message(f"Error sending message: {e}. Ensure Discord channel active.")
        except Exception as e:
            self.logger.log_message(f"Error in Discord operation: {e}. Check internet and browser.")

    def run(self):
        try:
            self.root.mainloop()
        except Exception as e:
            self.logger.log_message(f"Error running: {e}")
            if self.browser_manager.driver:
                self.browser_manager.driver.quit()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        bot = FaceBot(root)
        bot.run()
    except Exception as e:
        print(f"Error starting bot: {e}")