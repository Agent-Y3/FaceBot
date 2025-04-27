import tkinter as tk
from tkinter import scrolledtext, messagebox
import os
import subprocess
import shutil
from urllib.parse import quote
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
    """Konfiguration für FaceBot."""
    winscp_path: str = r"C:\Program Files (x86)\WinSCP\WinSCP.exe"
    putty_path: str = r"C:\Program Files\PuTTY\putty.exe"
    base_search_dir: str = os.path.expandvars(r"%userprofile%")
    spotify_search_url: str = "https://open.spotify.com/search/{}"
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
    """Kontext für Benutzeraktionen und Präferenzen."""
    last_application: Optional[str] = None
    last_action: Optional[str] = None
    user_preferences: Dict[str, int] = None

    def __post_init__(self):
        self.user_preferences = {"music": 0, "browser": 0, "document": 0}

class FaceBot:
    def __init__(self, root: tk.Tk):
        """Initialisiert den FaceBot."""
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
        
        self.send_button = tk.Button(self.input_frame, text="Senden", command=self.process_command)
        self.send_button.pack(side=tk.RIGHT, padx=5)
        
        self.listen_button = tk.Button(self.input_frame, text="Mikrofon", command=self.toggle_listening)
        self.listen_button.pack(side=tk.RIGHT)
        
        self.config_button = tk.Button(self.input_frame, text="Einstellungen", command=self._open_config_ui)
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
        
        self.log_message(f"Okay, ich bin bereit! Verwende Browser: {self.browser_name.capitalize()}. Sag z. B. 'Öffne Edge' oder 'Suche nach xAI'.")

    def _setup_logger(self) -> logging.Logger:
        """Konfiguriert den Logger."""
        logger = logging.getLogger("FaceBot")
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(handler)
        return logger

    def _setup_encryption(self) -> None:
        """Richtet die Verschlüsselung ein."""
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
            self.logger.error(f"Fehler bei der Verschlüsselungseinrichtung: {e}")
            self.fernet = None

    def _encrypt_data(self, data: str) -> str:
        """Verschlüsselt Daten."""
        if not self.fernet or not data:
            return data
        return base64.b64encode(self.fernet.encrypt(data.encode())).decode()

    def _decrypt_data(self, data: str) -> str:
        """Entschlüsselt Daten."""
        if not self.fernet or not data:
            return data
        try:
            return self.fernet.decrypt(base64.b64decode(data)).decode()
        except Exception:
            return data

    def _load_config(self) -> None:
        """Lädt Konfigurationsdaten."""
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
            self.logger.error(f"Fehler beim Laden der Konfiguration: {e}")

    def _save_config(self) -> None:
        """Speichert Konfigurationsdaten."""
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
            self.logger.error(f"Fehler beim Speichern der Konfiguration: {e}")

    def _speak(self, text: str) -> None:
        """Spricht den Text, wenn aktiviert."""
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
            self.logger.error(f"Fehler bei Sprachausgabe: {e}")

    def _update_audio_indicator(self, stream: pyaudio.Stream) -> None:
        """Aktualisiert den Audio-Indikator."""
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
        """Hört auf Sprachbefehle."""
        stream = None
        error_count = 0
        try:
            with sr.Microphone() as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=1)
                self.recognizer.dynamic_energy_threshold = True
                self.recognizer.energy_threshold = 300
                stream = pyaudio.PyAudio().open(format=pyaudio.paInt16, channels=1, rate=44100, input=True, frames_per_buffer=1024)
                threading.Thread(target=self._update_audio_indicator, args=(stream,), daemon=True).start()
                
                self.log_message("Spracherkennung aktiv. Sprich klar, z. B. 'Öffne Edge'. Deaktiviere in Einstellungen, falls Probleme auftreten.")
                while self.listening:
                    try:
                        audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=15)
                        audio_duration = len(audio.frame_data) / (audio.sample_rate * audio.sample_width)
                        if audio_duration < 0.5:
                            error_count = 0
                            continue
                        command = self.recognizer.recognize_google(audio, language="de-DE")
                        self.log_message(f"Du hast gesagt: {command}")
                        self.root.after(0, self.process_command, None, command)
                        error_count = 0
                    except sr.WaitTimeoutError:
                        error_count = 0
                    except sr.UnknownValueError:
                        error_count += 1
                        if error_count >= 3:
                            self.log_message("Mehrere unverständliche Eingaben. Sprich lauter oder überprüfe das Mikrofon. Deaktiviere die Spracherkennung in den Einstellungen, falls nötig.")
                            error_count = 0
                    except sr.RequestError as e:
                        self.log_message(f"Fehler bei der Spracherkennung: {e}. Überprüfe deine Internetverbindung.")
                        error_count = 0
                    except Exception as e:
                        self.log_message(f"Unbekannter Fehler bei der Spracherkennung: {e}. Versuche es erneut.")
                        error_count = 0
        except Exception as e:
            self.log_message(f"Fehler beim Starten der Spracherkennung: {e}. Spracherkennung deaktiviert. Verwende die Texteingabe.")
            self.listening = False
            self.listen_button.config(text="Mikrofon")
        finally:
            if stream:
                stream.stop_stream()
                stream.close()
                pyaudio.PyAudio().terminate()

    def toggle_listening(self) -> None:
        """Schaltet das Mikrofon ein/aus."""
        try:
            if not self.config.enable_listening:
                self.log_message("Spracherkennung ist in den Einstellungen deaktiviert. Aktiviere sie, um das Mikrofon zu nutzen.")
                return
            if not self.listening:
                try:
                    sr.Microphone()
                    self.listening = True
                    self.listen_button.config(text="Stop Mikrofon")
                    self.audio_thread = threading.Thread(target=self._listen_for_commands, daemon=True)
                    self.audio_thread.start()
                    self.log_message("Mikrofon ist an. Sag mir, was ich tun soll!")
                except Exception as e:
                    self.log_message(f"Fehler: Mikrofon nicht verfügbar ({e}). Verwende die Texteingabe.")
            else:
                self.listening = False
                self.listen_button.config(text="Mikrofon")
                if self.audio_thread:
                    self.audio_thread.join(timeout=1)
                    self.audio_thread = None
                for bar in self.bars:
                    self.indicator_canvas.coords(bar, 10 + self.bars.index(bar) * 20, 25, 25 + self.bars.index(bar) * 20, 25)
                self.log_message("Mikrofon ausgeschaltet.")
        except Exception as e:
            self.log_message(f"Fehler beim Umschalten des Mikrofons: {e}. Verwende die Texteingabe.")

    def _open_config_ui(self) -> None:
        """Öffnet die Konfigurations-UI."""
        try:
            config_window = tk.Toplevel(self.root)
            config_window.title("FaceBot Einstellungen")
            config_window.geometry("400x550")
            
            tk.Label(config_window, text="Server-Konfiguration", font=("Arial", 12, "bold")).pack(pady=10)
            tk.Label(config_window, text="Host (IP/Hostname):").pack()
            host_entry = tk.Entry(config_window)
            host_entry.pack()
            host_entry.insert(0, self.server_config.get('host', '') if self.server_config else '')
            
            tk.Label(config_window, text="Benutzername:").pack()
            username_entry = tk.Entry(config_window)
            username_entry.pack()
            username_entry.insert(0, self.server_config.get('username', '') if self.server_config else '')
            
            tk.Label(config_window, text="Passwort (optional, wenn Schlüssel):").pack()
            password_entry = tk.Entry(config_window, show="*")
            password_entry.pack()
            password_entry.insert(0, self.server_config.get('password', '') if self.server_config else '')
            
            tk.Label(config_window, text="Schlüsselpfad (.ppk, optional):").pack()
            key_path_entry = tk.Entry(config_window)
            key_path_entry.pack()
            key_path_entry.insert(0, self.server_config.get('key_path', '') if self.server_config else '')
            
            tk.Label(config_window, text="Discord-Konfiguration", font=("Arial", 12, "bold")).pack(pady=10)
            tk.Label(config_window, text="E-Mail:").pack()
            discord_email_entry = tk.Entry(config_window)
            discord_email_entry.pack()
            discord_email_entry.insert(0, self.config.discord_email)
            
            tk.Label(config_window, text="Passwort:").pack()
            discord_password_entry = tk.Entry(config_window, show="*")
            discord_password_entry.pack()
            discord_password_entry.insert(0, self.config.discord_password)
            
            tk.Label(config_window, text="Allgemeine Einstellungen", font=("Arial", 12, "bold")).pack(pady=10)
            speech_var = tk.BooleanVar(value=self.speech_enabled)
            tk.Checkbutton(config_window, text="Sprachausgabe aktivieren", variable=speech_var).pack()
            listening_var = tk.BooleanVar(value=self.config.enable_listening)
            tk.Checkbutton(config_window, text="Spracherkennung aktivieren", variable=listening_var).pack()
            
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
                self.log_message("Einstellungen gespeichert.")
                if not self.config.enable_listening and self.listening:
                    self.toggle_listening()
                config_window.destroy()
            
            tk.Button(config_window, text="Speichern", command=save).pack(pady=20)
            config_window.transient(self.root)
            config_window.grab_set()
        except Exception as e:
            self.log_message(f"Fehler beim Öffnen der Einstellungen: {e}")

    def _get_default_browser(self) -> str:
        """Ermittelt den Standardbrowser."""
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\Shell\Associations\UrlAssociations\http\UserChoice") as key:
                prog_id = winreg.QueryValueEx(key, "ProgId")[0]
            if "Chrome" in prog_id:
                return "chrome"
            elif "Firefox" in prog_id:
                return "firefox"
            elif "Edge" in prog_id or "IE" in prog_id:
                return "edge"
            return "chrome"
        except Exception:
            return "chrome"

    def _initialize_browser(self) -> None:
        """Initialisiert den Webbrowser."""
        try:
            self.browser_name = self._get_default_browser()
            self.log_message(f"Standardbrowser: {self.browser_name.capitalize()}.")
            
            if self.browser_name == "firefox":
                service = FirefoxService(GeckoDriverManager().install())
                self.driver = webdriver.Firefox(service=service)
            elif self.browser_name == "edge":
                service = EdgeService(EdgeChromiumDriverManager().install())
                self.driver = webdriver.Edge(service=service)
            else:
                service = ChromeService(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service)
            self.driver.maximize_window()
            self.context.last_application = self.browser_name
        except Exception as e:
            self.log_message(f"Fehler beim Starten von {self.browser_name}: {e}. Versuche Chrome...")
            try:
                service = ChromeService(ChromeDriverManager().install())
                self.driver = webdriver.Chrome(service=service)
                self.driver.maximize_window()
                self.browser_name = "chrome"
                self.context.last_application = "chrome"
            except Exception as e2:
                self.log_message(f"Fehler beim Starten von Chrome: {e2}. Browser-Funktionen sind deaktiviert.")
                self.driver = None

    def log_message(self, message: str) -> None:
        """Loggt eine Nachricht."""
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
            self.logger.error(f"Fehler beim Loggen: {e}")

    def _focus_application(self, app_name: str) -> bool:
        """Fokussiert eine Anwendung."""
        try:
            app_map = {
                "edge": "Microsoft Edge",
                "microsoft edge": "Microsoft Edge",
                "chrome": "Google Chrome",
                "firefox": "Firefox",
                "word": "Microsoft Word",
                "excel": "Microsoft Excel",
                "notepad": "Notepad",
                "winscp": "WinSCP",
                "discord": "Discord"
            }
            window_title = app_map.get(app_name.lower(), app_name)

            def enum_windows_callback(hwnd, results):
                if window_title.lower() in win32gui.GetWindowText(hwnd).lower() and win32gui.IsWindowVisible(hwnd):
                    results.append(hwnd)

            handles = []
            win32gui.EnumWindows(enum_windows_callback, handles)
            if handles:
                hwnd = handles[0]
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                win32gui.SetForegroundWindow(hwnd)
                self.log_message(f"Anwendung '{app_name}' fokussiert.")
                return True
            
            self.log_message(f"Anwendung '{app_name}' nicht gefunden. Starte sie...")
            self._open_file_or_program(app_name)
            return True
        except Exception as e:
            self.log_message(f"Fehler beim Fokussieren von '{app_name}': {e}. Stelle sicher, dass die Anwendung installiert ist.")
            return False

    def _parse_intent(self, command: str) -> Tuple[Optional[str], Dict[str, str]]:
        """Parst den Intent eines Befehls mit Regex-Fallback."""
        try:
            intent = None
            params = {}
            
            patterns = {
                "open": r"^(?:öffne|starte|mach\s+auf|open)\s+(.+)$",
                "play": r"^(?:spiele|spiel|play|musik)\s+(.+)$",
                "search": r"^(?:suche|google|find|such)\s*(?:nach)?\s+(.+?)(?:\s+in\s+(edge|microsoft\s+edge|chrome|firefox))?$",
                "close": r"^(?:schließe|close|beende|schließ)\s+(.+)$",
                "maximize": r"^(?:maximiere|maximize|vergrößere)\s+(.+)$",
                "write": r"^(?:schreibe|write|tippe|eingabe)\s+(.+?)(?:\s+in\s+(word|excel|notepad))?$",
                "save": r"^(?:speichere|save|sichern)\s+(.+)$",
                "click": r"^(?:klick|click|anklicken)$",
                "upload": r"^(?:upload|hochladen|lade\s+hoch|lade\s+datei)\s+(.+)$",
                "discord": r"^(?:discord|sende\s+nachricht)\s+an\s+(.+?)\s+(.+)$",
                "winscp": r"^(?:winscp|server|sftp)$",
                "putty": r"^(?:putty|ssh|terminal)$",
                "task": r"^(?:aufgabe|task|mache|erledige)\s+(.+)$",
                "help": r"^(?:hilfe|help|befehle)$",
                "exit": r"^(?:beenden|exit|schluss)$"
            }
            
            command_lower = command.lower().strip()
            for intent_name, pattern in patterns.items():
                match = re.match(pattern, command_lower)
                if match:
                    intent = intent_name
                    if intent == "search":
                        params["search_term"] = match.group(1).strip()
                        params["browser"] = match.group(2) if len(match.groups()) > 1 else None
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
                    "open": ["öffne", "starte", "mach auf", "open"],
                    "play": ["spiele", "spiel", "play", "musik"],
                    "search": ["suche", "google", "find", "such"],
                    "close": ["schließe", "close", "beende", "schließ"],
                    "maximize": ["maximiere", "maximize", "vergrößere"],
                    "write": ["schreibe", "write", "tippe", "eingabe"],
                    "save": ["speichere", "save", "sichern"],
                    "click": ["klick", "click", "anklicken"],
                    "upload": ["upload", "hochladen", "lade hoch", "lade datei"],
                    "discord": ["discord", "nachricht", "senden"],
                    "winscp": ["winscp", "server", "sftp"],
                    "putty": ["putty", "ssh", "terminal"],
                    "task": ["aufgabe", "task", "mache", "erledige"],
                    "help": ["hilfe", "help", "befehle"],
                    "exit": ["beenden", "exit", "schluss"]
                }
                for token in command_lower.split():
                    for key, keywords in intent_keywords.items():
                        if token in keywords:
                            intent = key
                            params["target"] = command_lower.replace(token, "").strip()
                            break
                    if intent:
                        break
            
            if intent in ["play", "search"]:
                self.context.user_preferences["music" if intent == "play" else "browser"] += 1
            elif intent in ["write", "save"] and params.get("app") in ["word", "excel"]:
                self.context.user_preferences["document"] += 1
            
            return intent, params
        except Exception as e:
            self.log_message(f"Fehler beim Parsen des Befehls: {e}. Versuche es mit einem klareren Befehl.")
            return None, {}

    def _execute_task(self, task: str) -> None:
        """Führt eine Aufgabe aus."""
        try:
            self.log_message(f"Mache: '{task}'...")
            self.context.last_action = "task"
            task_lower = task.lower().strip()
            steps = re.split(r"\s+und\s+|,\s*", task_lower)
            
            for step in steps:
                step = step.strip()
                if not step:
                    continue
                intent, params = self._parse_intent(step)
                
                if not intent:
                    self.log_message(f"Schritt '{step}' nicht verstanden. Sag z. B. 'Öffne Edge' oder 'Suche nach xAI'.")
                    continue
                
                if intent == "open" and "tab" in step:
                    browser = params.get("target", self.browser_name)
                    if browser not in ["edge", "chrome", "firefox", "microsoft edge"]:
                        browser = self.browser_name
                    self._focus_application(browser)
                    self.root.after(100, lambda: subprocess.run(["start", ""], shell=True))
                    self.log_message(f"Neuer Tab in {browser.capitalize()} geöffnet!")
                
                elif intent == "search":
                    search_term = params.get("search_term", step.replace("suche", "").replace("nach", "").strip())
                    browser = params.get("browser", self.browser_name)
                    if not search_term:
                        self.log_message("Kein Suchbegriff angegeben. Was soll ich suchen?")
                        continue
                    if browser not in ["edge", "microsoft edge", "chrome", "firefox"]:
                        browser = self.browser_name
                    if not self.driver:
                        self.log_message("Kein Browser verfügbar. Starte einen Browser neu.")
                        self._initialize_browser()
                        if not self.driver:
                            continue
                    self.log_message(f"Suche nach '{search_term}' in {browser.capitalize()}...")
                    self._open_file_or_program(browser)
                    self._focus_application(browser)
                    encoded_term = quote(search_term)
                    self.driver.get(f"https://www.google.com/search?q={encoded_term}")
                    WebDriverWait(self.driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    self.log_message(f"Suche nach '{search_term}' durchgeführt!")
                
                elif intent == "close":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.log_message("Keine Anwendung angegeben. Was soll ich schließen?")
                        continue
                    self._focus_application(app)
                    self.root.after(100, lambda: subprocess.run(["taskkill", "/IM", app + ".exe", "/F"], shell=True, capture_output=True))
                    self.log_message(f"Anwendung '{app}' geschlossen.")
                
                elif intent == "maximize":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.log_message("Keine Anwendung angegeben. Was soll ich maximieren?")
                        continue
                    self._focus_application(app)
                    self.log_message(f"Anwendung '{app}' maximiert.")
                
                elif intent == "open":
                    program = params.get("target")
                    if not program:
                        self.log_message("Was soll ich öffnen?")
                        continue
                    self._open_file_or_program(program)
                
                elif intent == "write":
                    text = params.get("text", step.replace("schreibe", "").replace("tippe", "").strip())
                    app = params.get("app", self.context.last_application)
                    if not app:
                        self.log_message("Keine Anwendung angegeben. Wohin soll ich schreiben?")
                        continue
                    self._focus_application(app)
                    self.root.after(100, lambda: subprocess.run(["powershell", "-Command", f"Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('{text}')"], shell=True))
                    self.log_message(f"Text '{text}' in {app} geschrieben.")
                
                elif intent == "save":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.log_message("Keine Anwendung angegeben. Was soll ich speichern?")
                        continue
                    self._focus_application(app)
                    self.root.after(100, lambda: subprocess.run(["powershell", "-Command", "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('^s')"], shell=True))
                    self.log_message(f"Dokument in '{app}' gespeichert.")
                
                else:
                    self.log_message(f"Schritt '{step}' nicht verstanden. Sag z. B. 'Öffne Edge' oder 'Suche nach xAI'.")
        except Exception as e:
            self.log_message(f"Fehler bei der Ausführung von '{task}': {e}. Versuche es mit einem anderen Befehl.")

    def _sanitize_input(self, cmd: str) -> str:
        """Bereinigt Benutzereingaben."""
        cmd = re.sub(r'[<>|;&$]', '', cmd)
        cmd = cmd.strip()
        if len(cmd) > 500:
            cmd = cmd[:500]
        return cmd

    def process_command(self, event: Optional[tk.Event] = None, command: Optional[str] = None) -> None:
        """Verarbeitet einen Befehl."""
        try:
            cmd = command if command else self.input_field.get().strip()
            if not command:
                self.input_field.delete(0, tk.END)
            
            if not cmd:
                return
            
            cmd = self._sanitize_input(cmd)
            self.log_message(f"Du: {cmd}")
            
            cmd = re.sub(r'^facebot[,]?[\s]*(hey\s)?', '', cmd, flags=re.IGNORECASE).strip().lower()
            
            intent, params = self._parse_intent(cmd)
            
            if not intent:
                self.log_message(f"Ich habe '{cmd}' nicht verstanden. Sag z. B. 'Öffne Edge', 'Suche nach xAI' oder 'Hilfe'.")
                return
            
            if intent == "exit":
                self.log_message("Okay, ich mache Schluss. Tschüss!")
                self.listening = False
                if self.driver:
                    self.driver.quit()
                self.root.quit()
            elif intent == "help":
                self.log_message("Ich kann folgendes:\n- Programme öffnen: 'Öffne Edge'\n- Suchen: 'Suche nach xAI'\n- Musik spielen: 'Spiele Shape of You'\n- Dateien hochladen: 'Lade dokument.txt hoch'\n- Discord-Nachrichten: 'Sende an @user Hallo'\n- Server: 'Starte WinSCP', 'Starte PuTTY'\n- Schreiben: 'Schreibe in Word Hallo'\n- Beenden: 'Beenden'\n- Hilfe: 'Hilfe'")
            elif intent == "click":
                self._perform_click()
            elif intent == "winscp":
                if not self.server_config:
                    self.log_message("Keine Server-Daten. Öffne die Einstellungen, um sie einzugeben.")
                    self._open_config_ui()
                    if not self.server_config:
                        return
                self._start_winscp()
            elif intent == "putty":
                if not self.server_config:
                    self.log_message("Keine Server-Daten. Öffne die Einstellungen, um sie einzugeben.")
                    self._open_config_ui()
                    if not self.server_config:
                        return
                self._start_putty()
            elif intent == "upload":
                file_name = params.get("target")
                if not file_name:
                    self.log_message("Welche Datei soll ich hochladen? Sag z. B. 'Lade dokument.txt hoch'.")
                    return
                if not self.server_config:
                    self.log_message("Keine Server-Daten. Öffne die Einstellungen, um sie einzugeben.")
                    self._open_config_ui()
                    if not self.server_config:
                        return
                self._upload_file(file_name)
            elif intent == "discord":
                target = params.get("target")
                message = params.get("message")
                if not target or not message:
                    self.log_message("Sag mir, an wen und was ich senden soll, z. B. 'Sende an @user Hallo'.")
                    return
                if not self.config.discord_email or not self.config.discord_password:
                    self.log_message("Keine Discord-Daten. Öffne die Einstellungen, um sie einzugeben.")
                    self._open_config_ui()
                    if not self.config.discord_email or not self.config.discord_password:
                        return
                self._send_discord_message(target, message)
            elif intent == "play":
                song_name = params.get("target")
                if not song_name:
                    self.log_message("Welches Lied soll ich spielen? Sag z. B. 'Spiele Shape of You'.")
                    return
                self._play_spotify_song(song_name)
            elif intent == "open":
                target = params.get("target")
                if not target:
                    self.log_message("Was soll ich öffnen? Sag z. B. 'Öffne Edge'.")
                    return
                self._open_file_or_program(target)
            elif intent == "task":
                task = params.get("target", cmd)
                if not task:
                    self.log_message("Was soll ich machen? Sag z. B. 'Suche nach xAI'.")
                    return
                self._execute_task(task)
            else:
                self.log_message(f"Befehl '{cmd}' nicht verstanden. Sag z. B. 'Öffne Edge' oder 'Hilfe'.")
        except Exception as e:
            self.log_message(f"Fehler beim Verarbeiten von '{cmd}': {e}. Versuche es mit einem anderen Befehl.")

    def _perform_click(self) -> None:
        """Führt einen Mausklick aus."""
        try:
            subprocess.run(["powershell", "-Command", "Add-Type -AssemblyName System.Windows.Forms; [System.Windows.Forms.SendKeys]::SendWait('{LEFT}')"], shell=True)
            self.log_message("Klick ausgeführt.")
        except Exception as e:
            self.log_message(f"Fehler beim Klicken: {e}. Stelle sicher, dass die GUI zugänglich ist.")

    def _start_winscp(self) -> None:
        """Startet WinSCP."""
        try:
            if not os.path.exists(self.config.winscp_path):
                self.log_message(f"WinSCP nicht gefunden unter '{self.config.winscp_path}'. Installiere WinSCP oder aktualisiere den Pfad in den Einstellungen.")
                return
            
            self.log_message("Starte WinSCP und verbinde mit dem Server...")
            self.context.last_action = "winscp"
            cmd = [self.config.winscp_path, f"sftp://{self.server_config['username']}@{self.server_config['host']}"]
            if self.server_config["key_path"]:
                cmd.append(f"/privatekey={self.server_config['key_path']}")
            process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.root.after(10000, lambda: self._check_process_output(process, "WinSCP"))
        except Exception as e:
            self.log_message(f"Fehler beim Starten von WinSCP: {e}. Überprüfe die Server-Daten und WinSCP-Installation.")

    def _start_putty(self) -> None:
        """Startet PuTTY."""
        try:
            if not os.path.exists(self.config.putty_path):
                self.log_message(f"PuTTY nicht gefunden unter '{self.config.putty_path}'. Installiere PuTTY oder aktualisiere den Pfad in den Einstellungen.")
                return
            
            self.log_message("Starte PuTTY und verbinde mit dem Server...")
            self.context.last_action = "putty"
            cmd = [self.config.putty_path, "-ssh", f"{self.server_config['username']}@{self.server_config['host']}"]
            if self.server_config["key_path"]:
                cmd.extend(["-i", self.server_config["key_path"]])
            process = subprocess.Popen(cmd, shell=False, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            self.root.after(10000, lambda: self._check_process_output(process, "PuTTY"))
        except Exception as e:
            self.log_message(f"Fehler beim Starten von PuTTY: {e}. Überprüfe die Server-Daten und PuTTY-Installation.")

    def _check_process_output(self, process: subprocess.Popen, name: str) -> None:
        """Überprüft die Ausgabe eines Prozesses."""
        try:
            stdout, stderr = process.communicate(timeout=5)
            if stderr:
                self.log_message(f"Fehler bei {name}: {stderr.decode()}")
            else:
                self.log_message(f"{name} erfolgreich gestartet!")
        except Exception as e:
            self.log_message(f"Fehler beim Überprüfen von {name}: {e}")

    def _upload_file(self, file_name: str) -> None:
        """Lädt eine Datei auf den Server hoch."""
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
                self.log_message(f"Datei '{file_name}' nicht gefunden. Gib einen gültigen Pfad oder Dateinamen an.")
                return
            
            if not os.path.exists(self.config.winscp_path):
                self.log_message(f"WinSCP nicht gefunden unter '{self.config.winscp_path}'. Installiere WinSCP.")
                return
            
            self.log_message(f"Lade '{file_path}' auf den Server...")
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
            self.log_message("Datei erfolgreich hochgeladen!")
        except Exception as e:
            self.log_message(f"Fehler beim Hochladen: {e}. Überprüfe die Datei, Server-Daten und WinSCP-Installation.")

    def _play_spotify_song(self, song_name: str) -> None:
        """Spielt ein Lied auf Spotify ab."""
        try:
            if not self.driver:
                self.log_message("Kein Browser verfügbar. Starte einen Browser neu.")
                self._initialize_browser()
                if not self.driver:
                    return
            self.log_message(f"Spiele '{song_name}' auf Spotify...")
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
                self.log_message(f"'{song_name}' wird abgespielt! Falls es nicht startet, überprüfe deine Spotify-Anmeldung.")
            except Exception as e:
                self.log_message(f"Fehler beim Abspielen: {e}. Bist du in Spotify eingeloggt? Ist das Lied verfügbar?")
        except Exception as e:
            self.log_message(f"Fehler beim Öffnen von Spotify: {e}. Überprüfe deine Internetverbindung und Browser.")

    def _send_discord_message(self, target: str, message: str) -> None:
        """Sendet eine Nachricht auf Discord."""
        try:
            if not self.driver:
                self.log_message("Kein Browser verfügbar. Starte einen Browser neu.")
                self._initialize_browser()
                if not self.driver:
                    return
            
            self.log_message(f"Sende Nachricht an '{target}' auf Discord...")
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
                self.log_message("Logge in Discord ein...")
                WebDriverWait(self.driver, 15).until(EC.presence_of_element_located((By.CSS_SELECTOR, self.config.discord_message_css)))
            except Exception as e:
                self.log_message(f"Fehler beim Discord-Login: {e}. Logge dich manuell ein oder überprüfe deine Zugangsdaten.")
                return
            
            try:
                message_field = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.CSS_SELECTOR, self.config.discord_message_css))
                )
                message_field.send_keys(f"@{target} {message}")
                message_field.send_keys(Keys.RETURN)
                self.log_message(f"Nachricht an '{target}' gesendet!")
            except Exception as e:
                self.log_message(f"Fehler beim Senden der Nachricht: {e}. Stelle sicher, dass der Discord-Kanal aktiv ist.")
        except Exception as e:
            self.log_message(f"Fehler beim Discord-Vorgang: {e}. Überprüfe deine Internetverbindung und Browser.")

    def _suggest_alternatives(self, target: str) -> List[Tuple[str, str]]:
        """Generiert intelligente Vorschläge für ein nicht gefundenes Ziel."""
        try:
            suggestions = []
            target_lower = target.lower()
            program_map = {
                "microsoft edge": "msedge.exe",
                "edge": "msedge.exe",
                "chrome": "chrome.exe",
                "firefox": "firefox.exe",
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
                    if self.context.user_preferences.get("browser", 0) > 0 and name in ["edge", "chrome", "firefox", "microsoft edge"]:
                        weight += 0.3
                    if self.context.last_application == name:
                        weight += 0.4
                    suggestions.append((name, score * weight, "ähnlicher Name (Tippfehler?)"))

            if "." in target:
                mime_type, _ = mimetypes.guess_type(target)
                if mime_type:
                    if mime_type.startswith("text"):
                        suggestions.append(("notepad", 90, "Textdatei erkannt, Notepad geeignet"))
                        if shutil.which("winword.exe"):
                            suggestions.append(("word", 85, "Textdatei erkannt, Word geeignet"))
                    elif mime_type.startswith("image"):
                        suggestions.append(("msedge.exe", 85, "Bilddatei erkannt, Browser geeignet"))
                    elif mime_type.startswith("application/vnd"):
                        if shutil.which("excel.exe"):
                            suggestions.append(("excel", 85, "Tabellendatei erkannt, Excel geeignet"))

            for name, exe in program_map.items():
                if shutil.which(exe) and name not in [s[0] for s in suggestions]:
                    weight = 0.8
                    if self.context.user_preferences.get(name, 0) > max(self.context.user_preferences.values(), default=0):
                        weight += 0.3
                    suggestions.append((name, weight * 80, "auf deinem System verfügbar"))

            suggestions = sorted(suggestions, key=lambda x: x[1], reverse=True)[:3]
            return [(name, reason) for name, _, reason in suggestions]
        except Exception as e:
            self.log_message(f"Fehler bei der Vorschlagserstellung: {e}")
            return []

    def _open_file_or_program(self, target: str) -> None:
        """Öffnet eine Datei oder ein Programm."""
        try:
            self.log_message(f"Öffne '{target}'...")
            self.context.last_action = "open"
            self.context.last_application = target
            
            program_map = {
                "microsoft edge": "msedge.exe",
                "edge": "msedge.exe",
                "chrome": "chrome.exe",
                "firefox": "firefox.exe",
                "word": "winword.exe",
                "excel": "excel.exe",
                "notepad": "notepad.exe",
                "winscp": "WinSCP.exe",
                "discord": "Discord.exe"
            }
            
            executable = program_map.get(target.lower())
            if executable:
                program_path = shutil.which(executable)
                if program_path:
                    subprocess.Popen([program_path], shell=False)
                    self.log_message(f"Programm '{target}' gestartet!")
                    return
            
            if os.path.isabs(target) and os.path.exists(target):
                os.startfile(target)
                self.log_message(f"'{target}' geöffnet!")
                return
            
            program_path = shutil.which(target)
            if program_path:
                subprocess.Popen([program_path], shell=False)
                self.log_message(f"Programm '{target}' gestartet!")
                return
            
            for root, _, files in os.walk(self.config.base_search_dir):
                if target in files:
                    file_path = os.path.join(root, target)
                    os.startfile(file_path)
                    self.log_message(f"Datei '{file_path}' geöffnet!")
                    return
            
            suggestions = self._suggest_alternatives(target)
            if suggestions:
                suggestion_text = "\n".join([f"- {name}: {reason}" for name, reason in suggestions])
                self.log_message(f"'{target}' nicht gefunden. Meintest du vielleicht:\n{suggestion_text}\nSag z. B. 'Öffne {suggestions[0][0]}' oder gib einen gültigen Pfad an.")
            else:
                self.log_message(f"'{target}' nicht gefunden. Gib einen gültigen Pfad oder Programmnamen an, z. B. 'Öffne Edge'.")
        except Exception as e:
            self.log_message(f"Fehler beim Öffnen von '{target}': {e}. Überprüfe den Pfad oder die Programmverfügbarkeit.")

    def run(self) -> None:
        """Startet die Hauptschleife."""
        try:
            self.root.mainloop()
        except Exception as e:
            self.log_message(f"Fehler beim Ausführen: {e}")
            if self.driver:
                self.driver.quit()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        bot = FaceBot(root)
        bot.run()
    except Exception as e:
        print(f"Fehler beim Starten des Bots: {e}")