import tkinter as tk
from tkinter import scrolledtext, messagebox
import os
import subprocess
import shutil
from urllib.parse import quote, urlparse
import logging
from typing import Optional, Dict, Tuple, List
from dataclasses import dataclass
import speech_recognition as sr
import pyaudio
import threading
import struct
import math
from gtts import gTTS
import pygame
import io
import re
import json
from pathlib import Path
from fuzzywuzzy import fuzz
import mimetypes
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

pygame.mixer.init()

@dataclass
class Config:
    """Configuration for FaceBot."""
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
    """Context for user actions and preferences."""
    last_application: Optional[str] = None
    last_action: Optional[str] = None
    user_preferences: Dict[str, int] = None

    def __post_init__(self):
        self.user_preferences = {"music": 0, "browser": 0, "document": 0}

class FaceBot:
    def __init__(self, root: tk.Tk):
        """Initializes the FaceBot."""
        self.root = root
        self.root.title("FaceBot")
        self.driver = None
        self.browser_name = None
        self.server_config = None
        self.config = Config()
        self.context = Context()
        self.logger = self._setup_logger()
        self.recognizer = sr.Recognizer()
        self.listening = False
        self.audio_thread = None
        self.fernet = None
        self.speech_enabled = self.config.enable_speech
        
        self.chat_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=20, state='disabled')
        self.chat_area.pack(padx=10, pady=10)
        
        self.input_frame = tk.Frame(root)
        self.input_frame.pack(padx=10, pady=5, fill=tk.X)
        
        self.input_field = tk.Entry(self.input_frame)
        self.input_field.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.input_field.bind("<Return>", self.process_command)
        
        self.send_button = tk.Button(self.input_frame, text="Send", command=self.process_command)
        self.send_button.pack(side=tk.RIGHT, padx=5)
        
        self.listen_button = tk.Button(self.input_frame, text="Microphone", command=self.toggle_listening)
        self.listen_button.pack(side=tk.RIGHT)
        
        self.config_button = tk.Button(self.input_frame, text="Settings", command=self._open_config_ui)
        self.config_button.pack(side=tk.RIGHT, padx=5)
        
        self.indicator_canvas = tk.Canvas(root, width=100, height=30, bg='black')
        self.indicator_canvas.pack(pady=5)
        self.bars = []
        for i in range(5):
            x = 10 + i * 20
            bar = self.indicator_canvas.create_rectangle(x, 25, x + 15, 25, fill='green')
            self.bars.append(bar)
        
        self._setup_encryption()
        self._load_config()
        self._initialize_browser()
        
        self.log_message(f"Okay, I'm ready! Using browser: {self.browser_name or 'Unknown'}. Say e.g., 'Open Edge', 'Search for xAI', or 'Go to https://check24.de'.")

    def _setup_logger(self) -> logging.Logger:
        """Sets up the logger."""
        logger = logging.getLogger("FaceBot")
        logger.setLevel(logging.DEBUG)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(handler)
        return logger

    def _setup_encryption(self) -> None:
        """Sets up encryption."""
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
            self.logger.error(f"Error setting up encryption: {e}")
            self.fernet = None

    def _encrypt_data(self, data: str) -> str:
        """Encrypts data."""
        if not self.fernet or not data:
            return data
        return base64.b64encode(self.fernet.encrypt(data.encode())).decode()

    def _decrypt_data(self, data: str) -> str:
        """Decrypts data."""
        if not self.fernet or not data:
            return data
        try:
            return self.fernet.decrypt(base64.b64decode(data)).decode()
        except Exception:
            return data

    def _load_config(self) -> None:
        """Loads configuration data."""
        try:
            if os.path.exists(self.config.config_file):
                with open(self.config.config_file, 'r') as f:
                    config_data = json.load(f)
                self.server_config = config_data.get('server_config', {})
                self.server_config['password'] = self._decrypt_data(self.server_config.get('password', ''))
                self.config.discord_email = self._decrypt_data(config_data.get('discord_email', ''))
                self.config.discord_password = self._decrypt_data(config_data.get('discord_password', ''))
                self.speech_enabled = config_data.get('speech_enabled', self.config.enable_speech)
                self.config.enable_listening = config_data.get('enable_listening', self.config.enable_listening)
        except Exception as e:
            self.logger.error(f"Error loading configuration: {e}")

    def _save_config(self) -> None:
        """Saves configuration data."""
        try:
            config_data = {
                'server_config': self.server_config or {},
                'discord_email': self._encrypt_data(self.config.discord_email),
                'discord_password': self._encrypt_data(self.config.discord_password),
                'speech_enabled': self.speech_enabled,
                'enable_listening': self.config.enable_listening
            }
            config_data['server_config']['password'] = self._encrypt_data(self.server_config.get('password', '') if self.server_config else '')
            with open(self.config.config_file, 'w') as f:
                json.dump(config_data, f, indent=4)
        except Exception as e:
            self.logger.error(f"Error saving configuration: {e}")

    def _speak(self, text: str) -> None:
        """Speaks the text if enabled."""
        if not self.speech_enabled:
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
            self.logger.error(f"Error in speech output: {e}")

    def _update_audio_indicator(self, stream: pyaudio.Stream) -> None:
        """Updates the audio indicator."""
        CHUNK = 1024
        while self.listening:
            try:
                data = stream.read(CHUNK, exception_on_overflow=False)
                rms = math.sqrt(abs(sum([(x / 32768) ** 2 for x in struct.unpack(f"{CHUNK * 2}h", data)]) / CHUNK))
                height = min(max(int(rms * 50), 5), 25)
                for i, bar in enumerate(self.bars):
                    self.indicator_canvas.coords(bar, 10 + i * 20, 25, 25 + i * 20, 25 - height + (i % 2) * 5)
                self.root.update()
            except Exception:
                pass

    def _listen_for_commands(self) -> None:
        """Listens for voice commands."""
        stream = None
        error_count = 0
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
                self.recognizer.dynamic_energy_threshold = True
                self.recognizer.energy_threshold = 300
                stream = pyaudio.PyAudio().open(format=pyaudio.paInt16, channels=1, rate=44100, input=True, frames_per_buffer=1024)
                threading.Thread(target=self._update_audio_indicator, args=(stream,), daemon=True).start()
                
                self.log_message("Speech recognition active. Speak clearly, e.g., 'Open Edge'. Disable in settings if issues occur.")
                while self.listening:
                    try:
                        audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=15)
                        audio_duration = len(audio.frame_data) / (audio.sample_rate * audio.sample_width)
                        if audio_duration < 0.5:
                            error_count = 0
                            continue
                        command = self.recognizer.recognize_google(audio, language="en-US")
                        self.log_message(f"You said: {command}")
                        self.root.after(0, self.process_command, None, command)
                        error_count = 0
                    except sr.WaitTimeoutError:
                        error_count = 0
                    except sr.UnknownValueError:
                        error_count += 1
                        if error_count >= 3:
                            self.log_message("Multiple unclear inputs. Speak louder or check the microphone. Disable speech recognition in settings if needed.")
                            error_count = 0
                    except sr.RequestError as e:
                        self.log_message(f"Speech recognition error: {e}. Check your internet connection.")
                        error_count = 0
                    except Exception as e:
                        self.log_message(f"Unknown speech recognition error: {e}. Try again.")
                        error_count = 0
        except Exception as e:
            self.log_message(f"Error starting speech recognition: {e}. Speech recognition disabled. Use text input.")
            self.listening = False
            self.listen_button.config(text="Microphone")
        finally:
            if stream:
                stream.stop_stream()
                stream.close()
                pyaudio.PyAudio().terminate()

    def toggle_listening(self) -> None:
        """Toggles the microphone on/off."""
        try:
            if not self.config.enable_listening:
                self.log_message("Speech recognition is disabled in settings. Enable it to use the microphone.")
                return
            if not self.listening:
                try:
                    sr.Microphone()
                    self.listening = True
                    self.listen_button.config(text="Stop Microphone")
                    self.audio_thread = threading.Thread(target=self._listen_for_commands, daemon=True)
                    self.audio_thread.start()
                    self.log_message("Microphone is on. Tell me what to do!")
                except Exception as e:
                    self.log_message(f"Error: Microphone not available ({e}). Use text input.")
            else:
                self.listening = False
                self.listen_button.config(text="Microphone")
                if self.audio_thread:
                    self.audio_thread.join(timeout=1)
                    self.audio_thread = None
                for bar in self.bars:
                    self.indicator_canvas.coords(bar, 10 + self.bars.index(bar) * 20, 25, 25 + self.bars.index(bar) * 20, 25)
                self.log_message("Microphone turned off.")
        except Exception as e:
            self.log_message(f"Error toggling microphone: {e}. Use text input.")

    def _open_config_ui(self) -> None:
        """Opens the configuration UI."""
        try:
            config_window = tk.Toplevel(self.root)
            config_window.title("FaceBot Settings")
            config_window.geometry("400x550")
            
            tk.Label(config_window, text="Server Configuration", font=("Arial", 12, "bold")).pack(pady=10)
            tk.Label(config_window, text="Host (IP/Hostname):").pack()
            host_entry = tk.Entry(config_window)
            host_entry.pack()
            host_entry.insert(0, self.server_config.get('host', '') if self.server_config else '')
            
            tk.Label(config_window, text="Username:").pack()
            username_entry = tk.Entry(config_window)
            username_entry.pack()
            username_entry.insert(0, self.server_config.get('username', '') if self.server_config else '')
            
            tk.Label(config_window, text="Password (optional if key is used):").pack()
            password_entry = tk.Entry(config_window, show="*")
            password_entry.pack()
            password_entry.insert(0, self.server_config.get('password', '') if self.server_config else '')
            
            tk.Label(config_window, text="Key Path (.ppk, optional):").pack()
            key_path_entry = tk.Entry(config_window)
            key_path_entry.pack()
            key_path_entry.insert(0, self.server_config.get('key_path', '') if self.server_config else '')
            
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
            speech_var = tk.BooleanVar(value=self.speech_enabled)
            tk.Checkbutton(config_window, text="Enable Speech Output", variable=speech_var).pack()
            listening_var = tk.BooleanVar(value=self.config.enable_listening)
            tk.Checkbutton(config_window, text="Enable Speech Recognition", variable=listening_var).pack()
            
            def save():
                host = host_entry.get().strip()
                username = username_entry.get().strip()
                password = password_entry.get().strip()
                key_path = key_path_entry.get().strip()
                if host and username:
                    self.server_config = {
                        "host": host,
                        "username": username,
                        "password": password,
                        "key_path": key_path
                    }
                self.config.discord_email = discord_email_entry.get().strip()
                self.config.discord_password = discord_password_entry.get().strip()
                self.speech_enabled = speech_var.get()
                self.config.enable_listening = listening_var.get()
                self._save_config()
                self.log_message("Settings saved.")
                if not self.config.enable_listening and self.listening:
                    self.toggle_listening()
                config_window.destroy()
            
            tk.Button(config_window, text="Save", command=save).pack(pady=20)
            config_window.transient(self.root)
            config_window.grab_set()
        except Exception as e:
            self.log_message(f"Error opening settings: {e}")

    def _get_default_browser(self) -> str:
        """Determines the default browser."""
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

    def _initialize_browser(self) -> None:
        """Initializes the web browser."""
        try:
            self.browser_name = self._get_default_browser()
            self.log_message(f"Default browser: {self.browser_name.capitalize()}.")
            
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
            self.context.last_application = self.browser_name
        except Exception as e:
            self.log_message(f"Error starting {self.browser_name}: {e}. Trying Chrome...")
            try:
                service = ChromeService(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service)
                self.driver.maximize_window()
                self.browser_name = "chrome"
                self.context.last_application = "chrome"
            except Exception as e2:
                self.log_message(f"Error starting Chrome: {e2}. Browser functions are disabled.")
                self.driver = None
                self.browser_name = "chrome"

    def log_message(self, message: str) -> None:
        """Logs a message."""
        try:
            def update_gui():
                self.chat_area.configure(state='normal')
                self.chat_area.insert(tk.END, f"{message}\n")
                self.chat_area.configure(state='disabled')
                self.chat_area.see(tk.END)
                self.root.update()
            self.root.after(0, update_gui)
            self.logger.info(message)
            self._speak(message)
        except Exception as e:
            self.logger.error(f"Error logging: {e}")

    def _focus_application(self, app_name: str) -> bool:
        """Focuses an application."""
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
                self.log_message(f"Application '{app_name}' focused.")
                return True
            
            self.log_message(f"Application '{app_name}' not found. Starting it...")
            self._open_file_or_program(app_name)
            return True
        except Exception as e:
            self.log_message(f"Error focusing '{app_name}': {e}. Ensure the application is installed.")
            return False

    def _parse_intent(self, command: str) -> Tuple[Optional[str], Dict[str, str]]:
        """Parses the intent of a command with regex fallback."""
        try:
            intent = None
            params = {}
            self.logger.debug(f"Parsing command: '{command}'")
            
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
            self.logger.debug(f"Normalized command: '{command_lower}'")
            for intent_name, pattern in patterns.items():
                match = re.match(pattern, command_lower)
                self.logger.debug(f"Testing pattern '{intent_name}': {'Match' if match else 'No match'}")
                if match:
                    intent = intent_name
                    if intent == "search":
                        params["search_term"] = match.group(1).strip()
                        params["browser"] = match.group(2).strip() if len(match.groups()) > 1 and match.group(2) else None
                        self.logger.debug(f"Search params: term='{params['search_term']}', browser='{params['browser']}'")
                    elif intent == "goto":
                        params["url"] = match.group(1).strip()
                        params["browser"] = match.group(2).strip() if len(match.groups()) > 1 and match.group(2) else None
                        self.logger.debug(f"Goto params: url='{params['url']}', browser='{params['browser']}'")
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
                            self.logger.debug(f"Fallback intent: '{intent}', target='{params['target']}'")
                            break
                    if intent:
                        break
            
            if intent in ["play", "search", "goto"]:
                self.context.user_preferences["music" if intent == "play" else "browser"] += 1
            elif intent in ["write", "save"] and params.get("app") in ["word", "excel"]:
                self.context.user_preferences["document"] += 1
            
            return intent, params
        except Exception as e:
            self.logger.error(f"Error parsing command: {e}. Try a clearer command.")
            return None, {}

    def _execute_task(self, task: str) -> None:
        """Executes a task."""
        try:
            self.log_message(f"Working on: '{task}'...")
            self.context.last_action = "task"
            task_lower = task.lower().strip()
            steps = re.split(r"\s+and\s+|,\s*", task_lower)
            
            for step in steps:
                step = step.strip()
                if not step:
                    continue
                intent, params = self._parse_intent(step)
                
                if not intent:
                    self.log_message(f"Step '{step}' not understood. Say e.g., 'Open Edge' or 'Search for xAI'.")
                    continue
                
                if intent == "open" and "tab" in step:
                    browser = params.get("target", self.browser_name)
                    self._focus_application(browser)
                    self.root.after(100, lambda: subprocess.run(["start", ""], shell=True))
                    self.log_message(f"New tab opened in {browser}!")
                
                elif intent == "search":
                    search_term = params.get("search_term", step.replace("search", "").replace("for", "").strip())
                    browser = params.get("browser", self.browser_name)
                    if not search_term:
                        self.log_message("No search term provided. What should I search for?")
                        continue
                    self._search_leta(search_term, browser)
                
                elif intent == "goto":
                    url = params.get("url", step.replace("go to", "").replace("goto", "").replace("navigate to", "").strip())
                    browser = params.get("browser", self.browser_name)
                    if not url:
                        self.log_message("No URL provided. What website should I go to?")
                        continue
                    self._navigate_to_url(url, browser)
                
                elif intent == "close":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.log_message("No application specified. What should I close?")
                        continue
                    self._focus_application(app)
                    self.root.after(100, lambda: subprocess.run(["taskkill", "/IM", app + ".exe", "/F"], shell=True, capture_output=True))
                    self.log_message(f"Application '{app}' closed.")
                
                elif intent == "maximize":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.log_message("No application specified. What should I maximize?")
                        continue
                    self._focus_application(app)
                    self.log_message(f"Application '{app}' maximized.")
                
                elif intent == "open":
                    program = params.get("target")
                    if not program:
                        self.log_message("What should I open?")
                        continue
                    self._open_file_or_program(program)
                
                elif intent == "write":
                    text = params.get("text", step.replace("write", "").replace("type", "").strip())
                    app = params.get("app", self.context.last_application)
                    if not app:
                        self.log_message("No application specified. Where should I write?")
                        continue
                    self._focus_application(app)
                    self.root.after(100, lambda: subprocess.run(["powershell", "-Command", f"Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('{text}')"], shell=True))
                    self.log_message(f"Text '{text}' written in {app}.")
                
                elif intent == "save":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.log_message("No application specified. What should I save?")
                        continue
                    self._focus_application(app)
                    self.root.after(100, lambda: subprocess.run(["powershell", "-Command", "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('^s')"], shell=True))
                    self.log_message(f"Document saved in '{app}'.")
                
                else:
                    self.log_message(f"Step '{step}' not understood. Say e.g., 'Open Edge' or 'Search for xAI'.")
        except Exception as e:
            self.log_message(f"Error executing '{task}': {e}. Try another command.")

    def _sanitize_input(self, cmd: str) -> str:
        """Sanitizes user input."""
        cmd = re.sub(r'[<>|;&$]', '', cmd)
        cmd = cmd.strip()
        if len(cmd) > 500:
            cmd = cmd[:500]
        return cmd

    def process_command(self, event: Optional[tk.Event] = None, command: Optional[str] = None) -> None:
        """Processes a command."""
        try:
            cmd = command if command else self.input_field.get().strip()
            if not command:
                self.input_field.delete(0, tk.END)
            
            if not cmd:
                return
            
            cmd = self._sanitize_input(cmd)
            self.log_message(f"You: {cmd}")
            
            cmd = re.sub(r'^facebot[,]?[\s]*(hey\s)?', '', cmd, flags=re.IGNORECASE).strip().lower()
            
            intent, params = self._parse_intent(cmd)
            
            if not intent:
                self.log_message(f"I didn't understand '{cmd}'. Say e.g., 'Open Edge', 'Search for xAI', 'Go to https://check24.de', or 'Help'.")
                return
            
            if intent == "exit":
                self.log_message("Okay, I'm shutting down. Bye!")
                self.listening = False
                if self.driver:
                    self.driver.quit()
                self.root.quit()
            elif intent == "help":
                self.log_message("I can do the following:\n- Open programs: 'Open Edge'\n- Search: 'Search for xAI'\n- Go to websites: 'Go to https://check24.de'\n- Play music: 'Play Shape of You'\n- Upload files: 'Upload document.txt'\n- Discord messages: 'Send to @user Hello'\n- Server: 'Start WinSCP', 'Start PuTTY'\n- Write: 'Write in Word Hello'\n- Exit: 'Exit'\n- Help: 'Help'")
            elif intent == "click":
                self._perform_click()
            elif intent == "winscp":
                if not self.server_config:
                    self.log_message("No server data. Open settings to enter them.")
                    self._open_config_ui()
                    if not self.server_config:
                        return
                self._start_winscp()
            elif intent == "putty":
                if not self.server_config:
                    self.log_message("No server data. Open settings to enter them.")
                    self._open_config_ui()
                    if not self.server_config:
                        return
                self._start_putty()
            elif intent == "upload":
                file_name = params.get("target")
                if not file_name:
                    self.log_message("Which file should I upload? Say e.g., 'Upload document.txt'.")
                    return
                if not self.server_config:
                    self.log_message("No server data. Open settings to enter them.")
                    self._open_config_ui()
                    if not self.server_config:
                        return
                self._upload_file(file_name)
            elif intent == "discord":
                target = params.get("target")
                message = params.get("message")
                if not target or not message:
                    self.log_message("Tell me who to send to and what, e.g., 'Send to @user Hello'.")
                    return
                if not self.config.discord_email or not self.config.discord_password:
                    self.log_message("No Discord credentials. Open settings to enter them.")
                    self._open_config_ui()
                    if not self.config.discord_email or not self.config.discord_password:
                        return
                self._send_discord_message(target, message)
            elif intent == "play":
                song_name = params.get("target")
                if not song_name:
                    self.log_message("Which song should I play? Say e.g., 'Play Shape of You'.")
                    return
                self._play_spotify_song(song_name)
            elif intent == "search":
                search_term = params.get("search_term")
                browser = params.get("browser", self.browser_name)
                if not search_term:
                    self.log_message("What should I search for? Say e.g., 'Search for xAI'.")
                    return
                self._search_leta(search_term, browser)
            elif intent == "goto":
                url = params.get("url")
                browser = params.get("browser", self.browser_name)
                if not url:
                    self.log_message("What website should I go to? Say e.g., 'Go to https://check24.de'.")
                    return
                self._navigate_to_url(url, browser)
            elif intent == "open":
                target = params.get("target")
                if not target:
                    self.log_message("What should I open? Say e.g., 'Open Edge'.")
                    return
                self._open_file_or_program(target)
            elif intent == "task":
                task = params.get("target", cmd)
                if not task:
                    self.log_message("What should I do? Say e.g., 'Search for xAI'.")
                    return
                self._execute_task(task)
            else:
                self.log_message(f"Command '{cmd}' not understood. Say e.g., 'Open Edge' or 'Help'.")
        except Exception as e:
            self.log_message(f"Error processing '{cmd}': {e}. Try another command.")

    def _navigate_to_url(self, url: str, browser: Optional[str]) -> None:
        """Navigates to a specified URL."""
        try:
            # Validate URL
            parsed_url = urlparse(url)
            if not parsed_url.scheme or not parsed_url.netloc:
                self.log_message(f"Invalid URL '{url}'. Please provide a valid URL, e.g., 'https://check24.de'.")
                return
            
            if not self.driver:
                self.log_message("No browser available. Restarting a browser.")
                self._initialize_browser()
                if not self.driver:
                    return
            
            browser = browser or self.browser_name or "chrome"
            self.log_message(f"Navigating to '{url}' in {browser}...")
            self.context.last_action = "goto"
            self.context.user_preferences["browser"] += 1
            
            self._open_file_or_program(browser)
            self._focus_application(browser)
            self.driver.get(url)
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            self.log_message(f"Navigation to '{url}' completed!")
        except Exception as e:
            self.log_message(f"Error navigating to '{url}': {e}. Check the URL, internet connection, and browser.")

    def _perform_click(self) -> None:
        """Performs a mouse click."""
        try:
            subprocess.run(["powershell", "-Command", "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('{LEFT}')"], shell=True)
            self.log_message("Click performed.")
        except Exception as e:
            self.log_message(f"Error performing click: {e}. Ensure the GUI is accessible.")

    def _start_winscp(self) -> None:
        """Starts WinSCP."""
        try:
            if not os.path.exists(self.config.winscp_path):
                self.log_message(f"WinSCP not found at '{self.config.winscp_path}'. Install WinSCP or update the path in settings.")
                return
            
            self.log_message("Starting WinSCP and connecting to the server...")
            self.context.last_action = "winscp"
            cmd = [self.config.winscp_path, f"sftp://{self.server_config['username']}@{self.server_config['host']}"]
            if self.server_config["key_path"]:
                cmd.append(f"/privatekey={self.server_config['key_path']}")
            process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.root.after(10000, lambda: self._check_process_output(process, "WinSCP"))
        except Exception as e:
            self.log_message(f"Error starting WinSCP: {e}. Check server data and WinSCP installation.")

    def _start_putty(self) -> None:
        """Starts PuTTY."""
        try:
            if not os.path.exists(self.config.putty_path):
                self.log_message(f"PuTTY not found at '{self.config.putty_path}'. Install PuTTY or update the path in settings.")
                return
            
            self.log_message("Starting PuTTY and connecting to the server...")
            self.context.last_action = "putty"
            cmd = [self.config.putty_path, "-ssh", f"{self.server_config['username']}@{self.server_config['host']}"]
            if self.server_config["key_path"]:
                cmd.extend(["-i", self.server_config["key_path"]])
            process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.root.after(10000, lambda: self._check_process_output(process, "PuTTY"))
        except Exception as e:
            self.log_message(f"Error starting PuTTY: {e}. Check server data and PuTTY installation.")

    def _check_process_output(self, process: subprocess.Popen, name: str) -> None:
        """Checks the output of a process."""
        try:
            stdout, stderr = process.communicate(timeout=5)
            if stderr:
                self.log_message(f"Error with {name}: {stderr.decode()}")
            else:
                self.log_message(f"{name} started successfully!")
        except Exception as e:
            self.log_message(f"Error checking {name}: {e}")

    def _upload_file(self, file_name: str) -> None:
        """Uploads a file to the server."""
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
                self.log_message(f"File '{file_name}' not found. Provide a valid path or filename.")
                return
            
            if not os.path.exists(self.config.winscp_path):
                self.log_message(f"WinSCP not found at '{self.config.winscp_path}'. Install WinSCP.")
                return
            
            self.log_message(f"Uploading '{file_path}' to the server...")
            self.context.last_action = "upload"
            script_path = os.path.join(os.path.expanduser("~"), "upload_script.txt")
            if self.server_config["key_path"]:
                script_content = (
                    f'open sftp://{self.server_config["username"]}@{self.server_config["host"]} -privatekey="{self.server_config["key_path"]}"\n'
                    f'put "{file_path}" /root/\n'
                    f'exit'
                )
            else:
                script_content = (
                    f'open sftp://{self.server_config["username"]}@{self.server_config["host"]}\n'
                    f'put "{file_path}" /root/\n'
                    f'exit'
                )
            with open(script_path, "w", encoding="utf-8") as f:
                f.write(script_content)
            
            cmd = [self.config.winscp_path, "/script", script_path]
            if not self.server_config["key_path"] and self.server_config["password"]:
                cmd.extend(["/parameter", self.server_config["password"]])
            subprocess.run(cmd, shell=False, check=True)
            os.remove(script_path)
            self.log_message("File uploaded successfully!")
        except Exception as e:
            self.log_message(f"Error uploading: {e}. Check the file, server data, and WinSCP installation.")

    def _play_spotify_song(self, song_name: str) -> None:
        """Plays a song on Spotify."""
        try:
            if not self.driver:
                self.log_message("No browser available. Restarting a browser.")
                self._initialize_browser()
                if not self.driver:
                    return
            self.log_message(f"Playing '{song_name}' on Spotify...")
            self.context.last_action = "play"
            self.context.user_preferences["music"] += 1
            
            encoded_song = quote(song_name)
            search_url = self.config.spotify_search_url.format(encoded_song)
            self.driver.get(search_url)
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            
            try:
                first_result = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, self.config.tracklist_css))
                )
                first_result.click()
                self.log_message(f"'{song_name}' is playing! If it doesn't start, check your Spotify login.")
            except Exception as e:
                self.log_message(f"Error playing: {e}. Are you logged into Spotify? Is the song available?")
        except Exception as e:
            self.log_message(f"Error opening Spotify: {e}. Check your internet connection and browser.")

    def _search_leta(self, search_term: str, browser: Optional[str]) -> None:
        """Searches using Mullvad Leta with Brave engine."""
        try:
            if not self.driver:
                self.log_message("No browser available. Restarting a browser.")
                self._initialize_browser()
                if not self.driver:
                    return
            browser = browser or self.browser_name or "chrome"
            self.log_message(f"Searching for '{search_term}' on Mullvad Leta (Brave) in {browser}...")
            self.context.last_action = "search"
            self.context.user_preferences["browser"] += 1
            
            encoded_term = quote(search_term)
            search_url = self.config.leta_search_url.format(encoded_term)
            self._open_file_or_program(browser)
            self._focus_application(browser)
            self.driver.get(search_url)
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            self.log_message(f"Search for '{search_term}' completed!")
        except Exception as e:
            self.log_message(f"Error searching on Leta: {e}. Check your internet connection and browser.")

    def _send_discord_message(self, target: str, message: str) -> None:
        """Sends a message on Discord."""
        try:
            if not self.driver:
                self.log_message("No browser available. Restarting a browser.")
                self._initialize_browser()
                if not self.driver:
                    return
            
            self.log_message(f"Sending message to '{target}' on Discord...")
            self.context.last_action = "discord"
            
            self.driver.get(self.config.discord_login_url)
            WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
            
            try:
                email_field = WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, self.config.discord_email_css))
                )
                password_field = self.driver.find_element(By.CSS_SELECTOR, self.config.discord_password_css)
                email_field.send_keys(self.config.discord_email)
                password_field.send_keys(self.config.discord_password)
                self.driver.find_element(By.CSS_SELECTOR, self.config.discord_submit_css).click()
                self.log_message("Logging into Discord...")
                WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, self.config.discord_message_css)))
            except Exception as e:
                self.log_message(f"Error logging into Discord: {e}. Log in manually or check your credentials.")
                return
            
            try:
                message_field = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, self.config.discord_message_css))
                )
                message_field.send_keys(f"@{target} {message}")
                message_field.send_keys(Keys.RETURN)
                self.log_message(f"Message sent to '{target}'!")
            except Exception as e:
                self.log_message(f"Error sending message: {e}. Ensure the Discord channel is active.")
        except Exception as e:
            self.log_message(f"Error in Discord operation: {e}. Check your internet connection and browser.")

    def _suggest_alternatives(self, target: str) -> List[Tuple[str, str]]:
        """Generates intelligent suggestions for a not-found target."""
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
                    weight = 1.0
                    if self.context.user_preferences.get("browser", 0) > 0 and name in ["edge", "chrome", "firefox", "opera", "microsoft edge"]:
                        weight += 0.3
                    if self.context.last_application == name:
                        weight += 0.4
                    suggestions.append((name, score * weight, "similar name (typo?)"))

            if "." in target:
                mime_type, _ = mimetypes.guess_type(target)
                if mime_type:
                    if mime_type.startswith("text"):
                        suggestions.append(("notepad", 90, "text file detected, Notepad suitable"))
                        if shutil.which("winword.exe"):
                            suggestions.append(("word", 85, "text file detected, Word suitable"))
                    elif mime_type.startswith("image"):
                        suggestions.append(("msedge.exe", 85, "image file detected, browser suitable"))
                    elif mime_type.startswith("application/vnd"):
                        if shutil.which("excel.exe"):
                            suggestions.append(("excel", 85, "spreadsheet detected, Excel suitable"))

            for name, exe in program_map.items():
                if shutil.which(exe) and name not in [s[0] for s in suggestions]:
                    weight = 0.8
                    if self.context.user_preferences.get(name, 0) > max(self.context.user_preferences.values(), default=0):
                        weight += 0.3
                    suggestions.append((name, weight * 80, "available on your system"))

            suggestions = sorted(suggestions, key=lambda x: x[1], reverse=True)[:3]
            return [(name, reason) for name, _, reason in suggestions]
        except Exception as e:
            self.log_message(f"Error generating suggestions: {e}")
            return []

    def _open_file_or_program(self, target: str) -> None:
        """Opens a file or program."""
        try:
            self.log_message(f"Opening '{target}'...")
            self.context.last_action = "open"
            self.context.last_application = target
            
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
                    self.log_message(f"Program '{target}' started!")
                    return
            
            program_path = shutil.which(target)
            if program_path:
                subprocess.Popen([program_path], shell=False)
                self.log_message(f"Program '{target}' started!")
                return
            
            if os.path.isabs(target) and os.path.exists(target):
                os.startfile(target)
                self.log_message(f"'{target}' opened!")
                return
            
            for root, _, files in os.walk(self.config.base_search_dir):
                if target in files:
                    file_path = os.path.join(root, target)
                    os.startfile(file_path)
                    self.log_message(f"File '{file_path}' opened!")
                    return
            
            suggestions = self._suggest_alternatives(target)
            if suggestions:
                suggestion_text = "\n".join([f"- {name}: {reason}" for name, reason in suggestions])
                self.log_message(f"'{target}' not found. Did you mean:\n{suggestion_text}\nSay e.g., 'Open {suggestions[0][0]}' or provide a valid path.")
            else:
                self.log_message(f"'{target}' not found. Provide a valid path or program name, e.g., 'Open Edge'.")
        except Exception as e:
            self.log_message(f"Error opening '{target}': {e}. Check the path or program availability.")

    def run(self) -> None:
        """Starts the main loop."""
        try:
            self.root.mainloop()
        except Exception as e:
            self.log_message(f"Error running: {e}")
            if self.driver:
                self.driver.quit()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        bot = FaceBot(root)
        bot.run()
    except Exception as e:
        print(f"Error starting the bot: {e}")