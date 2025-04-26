import tkinter as tk
from tkinter import scrolledtext, messagebox
import os
import psutil
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
import re
import winreg
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
import spacy

@dataclass
class Config:
    winscp_path: str = r"C:\Program Files (x86)\WinSCP\WinSCP.exe"
    putty_path: str = r"C:\Program Files\PuTTY\putty.exe"
    base_search_dir: str = r"C:\Users\xByYu"
    discord_email: str = ""
    discord_password: str = ""
    spotify_search_url: str = "https://open.spotify.com/search/{}"
    discord_login_url: str = "https://discord.com/login"
    tracklist_xpath: str = "//div[@data-testid='tracklist-row']"
    discord_email_xpath: str = "//input[@name='email']"
    discord_password_xpath: str = "//input[@name='password']"
    discord_submit_xpath: str = "//button[@type='submit']"

@dataclass
class Context:
    last_application: Optional[str] = None
    last_action: Optional[str] = None
    user_preferences: Dict[str, int] = None

    def __post_init__(self):
        self.user_preferences = {"music": 0, "browser": 0, "document": 0}

class FaceBot:
    def __init__(self, root: tk.Tk):
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
        self.nlp = spacy.load("de_core_news_lg")
        
        pygame.mixer.init()
        
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
        
        self.indicator_canvas = tk.Canvas(root, width=100, height=30, bg='black')
        self.indicator_canvas.pack(pady=5)
        self.bars = []
        for i in range(5):
            x = 10 + i * 20
            bar = self.indicator_canvas.create_rectangle(x, 25, x + 15, 25, fill='green')
            self.bars.append(bar)
        
        self._initialize_browser()
        
        self.log_message(f"Okay, ich bin bereit! Verwende Browser: {self.browser_name.capitalize()}. Sag mir, was ich tun soll, z. B. ‚Öffne Edge‘ oder ‚Spiele Musik‘. Hinweis: Bildschirmerkennung ist deaktiviert.")

    def _setup_logger(self) -> logging.Logger:
        logger = logging.getLogger("FaceBot")
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(handler)
        return logger

    def _speak(self, text: str) -> None:
        def play_audio():
            try:
                tts = gTTS(text=text, lang="de")
                audio_file = io.BytesIO()
                tts.write_to_fp(audio_file)
                audio_file.seek(0)
                
                pygame.mixer.music.load(audio_file)
                pygame.mixer.music.play()
                while pygame.mixer.music.get_busy():
                    time.sleep(0.1)
            except Exception as e:
                self.logger.error(f"Fehler bei Sprachausgabe: {e}")
        
        threading.Thread(target=play_audio, daemon=True).start()

    def _update_audio_indicator(self, stream: pyaudio.Stream) -> None:
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
        with sr.Microphone() as source:
            self.recognizer.adjust_for_ambient_noise(source)
            audio_stream = pyaudio.PyAudio().open(format=pyaudio.paInt16, channels=1, rate=44100, input=True, frames_per_buffer=1024)
            threading.Thread(target=self._update_audio_indicator, args=(audio_stream,), daemon=True).start()
            
            while self.listening:
                try:
                    self.log_message("Ich höre zu, was willst du?")
                    audio = self.recognizer.listen(source, timeout=5, phrase_time_limit=10)
                    command = self.recognizer.recognize_google(audio, language="de-DE")
                    self.log_message(f"Du hast gesagt: {command}")
                    self.process_command(command=command)
                except sr.WaitTimeoutError:
                    pass
                except sr.UnknownValueError:
                    self.log_message("Sorry, ich hab dich nicht verstanden. Kannst du das klarer sagen?")
                except sr.RequestError as e:
                    self.log_message(f"Problem mit der Spracherkennung: {e}")
                except Exception as e:
                    self.log_message(f"Fehler bei der Spracherkennung: {e}")

    def toggle_listening(self) -> None:
        if not self.listening:
            self.listening = True
            self.listen_button.config(text="Stop Mikrofon")
            self.audio_thread = threading.Thread(target=self._listen_for_commands, daemon=True)
            self.audio_thread.start()
            self.log_message("Mikrofon ist an. Sag mir, was ich tun soll!")
        else:
            self.listening = False
            self.listen_button.config(text="Mikrofon")
            self.audio_thread.join(timeout=1)
            self.audio_thread = None
            for bar in self.bars:
                self.indicator_canvas.coords(bar, 10 + self.bars.index(bar) * 20, 25, 25 + self.bars.index(bar) * 20, 25)
            self.log_message("Mikrofon ausgeschaltet.")

    def _prompt_server_config(self) -> bool:
        config_window = tk.Toplevel(self.root)
        config_window.title("Root-Server-Daten eingeben")
        config_window.geometry("400x300")
        
        tk.Label(config_window, text="Host (IP/Hostname):").pack(pady=5)
        host_entry = tk.Entry(config_window)
        host_entry.pack(pady=5)
        
        tk.Label(config_window, text="Benutzername:").pack(pady=5)
        username_entry = tk.Entry(config_window)
        username_entry.pack(pady=5)
        
        tk.Label(config_window, text="Passwort (optional, wenn Schlüssel verwendet):").pack(pady=5)
        password_entry = tk.Entry(config_window, show="*")
        password_entry.pack(pady=5)
        
        tk.Label(config_window, text="Schlüsselpfad (.ppk, optional):").pack(pady=5)
        key_path_entry = tk.Entry(config_window)
        key_path_entry.pack(pady=5)
        
        def submit():
            host = host_entry.get().strip()
            username = username_entry.get().strip()
            password = password_entry.get().strip()
            key_path = key_path_entry.get().strip()
            
            if not host or not username:
                messagebox.showerror("Fehler", "Host und Benutzername sind erforderlich!")
                return
            
            self.server_config = {
                "host": host,
                "username": username,
                "password": password,
                "key_path": key_path
            }
            self.log_message("Root-Server-Daten gespeichert.")
            config_window.destroy()
        
        tk.Button(config_window, text="Bestätigen", command=submit).pack(pady=10)
        config_window.transient(self.root)
        config_window.grab_set()
        self.root.wait_window(config_window)
        
        return self.server_config is not None

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
            return "chrome"
        except Exception:
            return "chrome"
    
    def _initialize_browser(self) -> None:
        self.browser_name = self._get_default_browser()
        self.log_message(f"Standardbrowser: {self.browser_name.capitalize()}.")
        
        try:
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
            self.log_message(f"Fehler beim Browser-Start: {e}. Verwende Chrome.")
            service = ChromeService(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service)
            self.driver.maximize_window()
            self.browser_name = "chrome"
            self.context.last_application = "chrome"
    
    def log_message(self, message: str) -> None:
        self.chat_area.configure(state='normal')
        self.chat_area.insert(tk.END, f"{message}\n")
        self.chat_area.configure(state='disabled')
        self.chat_area.see(tk.END)
        self.root.update()
        self.logger.info(message)
        self._speak(message)
    
    def _focus_application(self, app_name: str) -> bool:
        self.log_message(f"Fokussieren von '{app_name}' nicht möglich, Bildschirmerkennung deaktiviert. Starte Anwendung direkt.")
        return False
    
    def _parse_intent(self, command: str) -> Tuple[Optional[str], Dict[str, str]]:
        doc = self.nlp(command.lower())
        intent = None
        params = {}
        
        intent_keywords = {
            "open": ["öffne", "mach auf", "starte", "open"],
            "play": ["spiele", "spiel", "play", "musik"],
            "search": ["suche", "google", "find", "such"],
            "close": ["schließe", "close", "beende", "schließ"],
            "maximize": ["maximiere", "maximize", "vergrößere"],
            "write": ["schreibe", "write", "tippe", "eingabe"],
            "save": ["speichere", "save", "sichern"],
            "click": ["klick", "click", "anklicken"],
            "virus": ["virus", "viren", "scan", "prüfe"],
            "upload": ["upload", "hochladen", "lade hoch"],
            "discord": ["discord", "nachricht", "senden"],
            "winscp": ["winscp", "server", "sftp"],
            "putty": ["putty", "ssh", "terminal"],
            "task": ["aufgabe", "task", "mache", "erledige"],
            "help": ["hilfe", "help", "befehle"]
        }
        
        for token in doc:
            for key, keywords in intent_keywords.items():
                if token.text in keywords:
                    intent = key
                    break
            if intent:
                break
        
        for ent in doc.ents:
            if ent.label_ in ["PERSON", "ORG", "PRODUCT"]:
                params["target"] = ent.text
            elif ent.label_ == "GPE":
                params["location"] = ent.text
        
        for token in doc:
            if token.pos_ == "NOUN" and "target" not in params:
                params["target"] = token.text
            elif token.pos_ == "VERB" and not intent:
                intent = token.text
        
        if not intent:
            reference_words = ["es", "das", "dieses"]
            for token in doc:
                if token.text in reference_words and self.context.last_application:
                    params["target"] = self.context.last_application
                    intent = "open" if "open" in command else "close"
                    break
        
        if intent == "play" and "target" in params:
            self.context.user_preferences["music"] += 1
        elif intent in ["open", "search"] and params.get("target") in ["edge", "chrome", "firefox"]:
            self.context.user_preferences["browser"] += 1
        elif intent in ["write", "save"] and params.get("target") in ["word", "excel"]:
            self.context.user_preferences["document"] += 1
        
        if not intent:
            max_sim = 0
            best_intent = None
            command_doc = self.nlp(command)
            for key, keywords in intent_keywords.items():
                for kw in keywords:
                    kw_doc = self.nlp(kw)
                    sim = command_doc.similarity(kw_doc)
                    if sim > max_sim and sim > 0.7:
                        max_sim = sim
                        best_intent = key
            intent = best_intent
        
        return intent, params

    def _execute_task(self, task: str) -> None:
        self.log_message(f"Mache: '{task}'...")
        self.context.last_action = "task"
        task_lower = task.lower()
        steps = task_lower.split(" und ")
        
        for step in steps:
            step = step.strip()
            try:
                intent, params = self._parse_intent(step)
                
                if intent == "open" and "tab" in step:
                    self.log_message("Öffnen neuer Tabs nicht möglich, Bildschirmerkennung deaktiviert. Öffne Browser stattdessen.")
                    browser = params.get("target", self.browser_name)
                    if browser not in ["edge", "chrome", "firefox"]:
                        browser = self.browser_name
                    
                    if browser == "edge":
                        subprocess.Popen("msedge")
                    elif browser == "firefox":
                        subprocess.Popen("firefox")
                    else:
                        subprocess.Popen("chrome")
                    self.log_message(f"Browser {browser.capitalize()} gestartet!")
                
                elif intent == "search":
                    browser = params.get("target", self.browser_name)
                    if browser not in ["edge", "chrome", "firefox"]:
                        browser = self.browser_name
                    search_term = step.replace("suche", "").replace("google", "").replace("nach", "").strip()
                    self.log_message(f"Suche nach '{search_term}' in {browser.capitalize()}...")
                    
                    self.driver.get(f"https://www.google.com/search?q={quote(search_term)}")
                    self.log_message(f"Suche nach '{search_term}' in {browser.capitalize()} durchgeführt!")
                
                elif intent == "close":
                    app = params.get("target", self.context.last_application)
                    if not app:
                        self.log_message("Keine Anwendung angegeben. Was soll ich schließen?")
                        continue
                    self.log_message(f"Schließen von '{app}' nicht möglich, Bildschirmerkennung deaktiviert.")
                
                elif intent == "maximize":
                    self.log_message("Maximieren nicht möglich, Bildschirmerkennung deaktiviert.")
                
                elif intent == "open":
                    program = params.get("target", step.replace("öffne", "").strip())
                    if not program:
                        self.log_message("Was soll ich öffnen?")
                        continue
                    self._open_file_or_program(program)
                
                elif intent == "write" and "word" in step:
                    self.log_message("Schreiben in Word nicht möglich, Bildschirmerkennung deaktiviert.")
                
                elif intent == "save":
                    self.log_message("Speichern nicht möglich, Bildschirmerkennung deaktiviert.")
                
                elif intent == "write":
                    self.log_message("Tippen nicht möglich, Bildschirmerkennung deaktiviert.")
                
                else:
                    self.log_message(f"Schritt '{step}' nicht verstanden. Sag z. B. ‚Öffne Edge‘, ‚Suche nach xAI‘ oder ‚Spiele Musik‘.")
            
            except Exception as e:
                self.log_message(f"Fehler bei Schritt '{step}': {e}")

    def process_command(self, event: Optional[tk.Event] = None, command: Optional[str] = None) -> None:
        cmd = command if command else self.input_field.get().strip()
        if not command:
            self.input_field.delete(0, tk.END)
        
        if not cmd:
            return
        
        self.log_message(f"Du: {cmd}")
        
        cmd = re.sub(r'^facebot[,]?[\s]*(hey\s)?', '', cmd, flags=re.IGNORECASE).strip().lower()
        
        intent, params = self._parse_intent(cmd)
        
        if cmd in ["beenden", "exit"] or intent == "exit":
            self.log_message("Okay, ich mach Schluss. Tschüss!")
            self.listening = False
            self.driver.quit()
            self.root.quit()
        elif cmd in ["hilfe", "help"] or intent == "help":
            self.log_message("Ich kann folgendes:\n- Viren prüfen: ‚Prüf auf Viren‘\n- Server: ‚Starte WinSCP‘, ‚Starte PuTTY‘\n- Dateien hochladen: ‚Lade Datei hoch‘\n- Discord: ‚Sende Nachricht an @user‘ (manuell abschicken)\n- Musik: ‚Spiele Shape of You‘\n- Programme öffnen: ‚Öffne Edge‘\n- Aufgaben: ‚Suche nach xAI‘\n- Beenden: ‚Beenden‘\nSag mir einfach, was du willst! Hinweis: Bildschirmerkennung ist deaktiviert.")
        elif intent == "virus":
            self._check_for_viruses()
        elif intent == "click":
            self.log_message("Klicken nicht möglich, Bildschirmerkennung deaktiviert.")
        elif intent == "winscp":
            if not self.server_config and not self._prompt_server_config():
                self.log_message("Keine Server-Daten. Abbruch.")
                return
            self._start_winscp()
        elif intent == "putty":
            if not self.server_config and not self._prompt_server_config():
                self.log_message("Keine Server-Daten. Abbruch.")
                return
            self._start_putty()
        elif intent == "upload":
            file_path = params.get("target", cmd.replace("upload", "").strip())
            if not file_path:
                self.log_message("Welche Datei soll ich hochladen?")
                return
            if not self.server_config and not self._prompt_server_config():
                self.log_message("Keine Server-Daten. Abbruch.")
                return
            self._upload_file(file_path)
        elif intent == "discord":
            parts = cmd.replace("discord", "").strip().split(maxsplit=1)
            if len(parts) < 2:
                self.log_message("Sag mir, an wen und was ich senden soll, z. B. ‚Sende an @user Hallo‘.")
            else:
                target, message = parts
                self._send_discord_message(target, message)
        elif intent == "play":
            song_name = params.get("target", cmd.replace("spiele", "").replace("play", "").strip())
            if not song_name:
                self.log_message("Welches Lied soll ich spielen?")
            else:
                self._play_spotify_song(song_name)
        elif intent == "open":
            target = params.get("target", cmd.replace("öffne", "").strip())
            if not target:
                self.log_message("Was soll ich öffnen?")
            else:
                self._open_file_or_program(target)
        elif intent == "task":
            task = params.get("target", cmd.replace("aufgabe", "").strip())
            if not task:
                self.log_message("Was soll ich machen? Sag z. B. ‚Suche nach xAI‘.")
            else:
                self._execute_task(task)
        else:
            self.log_message(f"Ich hab '{cmd}' nicht ganz verstanden. Meinst du etwas wie ‚Öffne Edge‘ oder ‚Spiele Musik‘?")
            if self.context.user_preferences["music"] > self.context.user_preferences["browser"]:
                self.log_message("Da du oft Musik hörst, soll ich dir ein Lied vorschlagen? Sag z. B. ‚Spiele Shape of You‘.")
            elif self.context.user_preferences["browser"] > 0:
                self.log_message("Willst du im Browser was machen? Sag z. B. ‚Suche nach xAI‘.")

    def _check_for_viruses(self) -> None:
        self.log_message("Prüfe auf Viren...")
        self.context.last_action = "virus"
        suspicious_processes = ['malware.exe', 'virus.exe']
        found = False
        for proc in psutil.process_iter(['name']):
            if proc.info['name'].lower() in suspicious_processes:
                self.log_message(f"Achtung: Verdächtiger Prozess: {proc.info['name']}")
                found = True
        if not found:
            self.log_message("Alles sauber, keine Viren gefunden.")
    
    def _perform_click(self) -> None:
        self.log_message("Klicken nicht möglich, Bildschirmerkennung deaktiviert.")
    
    def _start_winscp(self) -> None:
        if not os.path.exists(self.config.winscp_path):
            self.log_message("WinSCP nicht gefunden! Bitte installiere es.")
            return
        
        self.log_message("Starte WinSCP und verbinde mit dem Server...")
        self.context.last_action = "winscp"
        try:
            if self.server_config["key_path"]:
                cmd = f'"{self.config.winscp_path}" sftp://{self.server_config["username"]}@{self.server_config["host"]} /privatekey="{self.server_config["key_path"]}"'
            else:
                cmd = f'"{self.config.winscp_path}" sftp://{self.server_config["username"]}@{self.server_config["host"]}'
            subprocess.Popen(cmd, shell=True)
            time.sleep(3)
            if self.server_config["password"] and not self.server_config["key_path"]:
                self.log_message("Passwort-Eingabe nicht möglich, Bildschirmerkennung deaktiviert. Gib das Passwort manuell ein.")
            self.log_message("WinSCP gestartet! Du kannst jetzt Dateien übertragen.")
        except Exception as e:
            self.log_message(f"Fehler beim Starten von WinSCP: {e}")
    
    def _start_putty(self) -> None:
        if not os.path.exists(self.config.putty_path):
            self.log_message("PuTTY nicht gefunden! Bitte installiere es.")
            return
        
        self.log_message("Starte PuTTY und verbinde mit dem Server...")
        self.context.last_action = "putty"
        try:
            if self.server_config["key_path"]:
                cmd = f'"{self.config.putty_path}" -ssh {self.server_config["username"]}@{self.server_config["host"]} -i "{self.server_config["key_path"]}"'
            else:
                cmd = f'"{self.config.putty_path}" -ssh {self.server_config["username"]}@{self.server_config["host"]} -pw {self.server_config["password"]}'
            subprocess.Popen(cmd, shell=True)
            self.log_message("PuTTY verbunden! Du kannst jetzt Shell-Befehle ausführen.")
        except Exception as e:
            self.log_message(f"Fehler beim Starten von PuTTY: {e}")
    
    def _upload_file(self, file_path: str) -> None:
        if not file_path or not os.path.exists(file_path):
            self.log_message("Datei nicht gefunden! Gib einen gültigen Pfad an.")
            return
        
        if not os.path.exists(self.config.winscp_path):
            self.log_message("WinSCP nicht gefunden! Bitte installiere es.")
            return
        
        self.log_message(f"Lade '{file_path}' auf den Server...")
        self.context.last_action = "upload"
        try:
            if self.server_config["key_path"]:
                script_content = (
                    f'open sftp://{self.server_config["username"]}@{self.server_config["host"]} -privatekey="{self.server_config["key_path"]}"\n'
                    f'put "{file_path}" /root/\n'
                    f'exit'
                )
            else:
                script_content = (
                    f'open sftp://{self.server_config["username"]}:{self.server_config["password"]}@{self.server_config["host"]}\n'
                    f'put "{file_path}" /root/\n'
                    f'exit'
                )
            script_path = "upload_script.txt"
            with open(script_path, "w") as f:
                f.write(script_content)
            
            cmd = f'"{self.config.winscp_path}" /script="{script_path}"'
            subprocess.run(cmd, shell=True)
            os.remove(script_path)
            self.log_message("Datei erfolgreich hochgeladen!")
        except Exception as e:
            self.log_message(f"Fehler beim Hochladen: {e}")
    
    def _play_spotify_song(self, song_name: str) -> None:
        self.log_message(f"Spiele '{song_name}' auf Spotify...")
        self.context.last_action = "play"
        self.context.user_preferences["music"] += 1
        
        try:
            encoded_song = quote(song_name)
            search_url = self.config.spotify_search_url.format(encoded_song)
            self.driver.get(search_url)
            time.sleep(3)
            
            try:
                first_result = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, self.config.tracklist_xpath))
                )
                first_result.click()
                self.log_message(f"'{song_name}' wird abgespielt! Falls ein CAPTCHA erscheint, löse es manuell.")
            except Exception as e:
                self.log_message(f"Fehler beim Abspielen auf Spotify: {e}. Bist du eingeloggt? Ist das Lied verfügbar? Falls ein CAPTCHA erscheint, löse es manuell.")
        
        except Exception as e:
            self.log_message(f"Fehler beim Öffnen von Spotify: {e}")
    
    def _open_file_or_program(self, target: str) -> None:
        self.log_message(f"Öffne '{target}'...")
        self.context.last_action = "open"
        self.context.last_application = target
        
        try:
            if os.path.isabs(target) and os.path.exists(target):
                os.startfile(target)
                self.log_message(f"'{target}' geöffnet!")
                return
            
            program_path = shutil.which(target)
            if program_path:
                subprocess.Popen(program_path, shell=True)
                self.log_message(f"Programm '{target}' gestartet!")
                return
            
            for root, _, files in os.walk(self.config.base_search_dir):
                if target in files:
                    file_path = os.path.join(root, target)
                    os.startfile(file_path)
                    self.log_message(f"Datei '{file_path}' geöffnet!")
                    return
            
            self.log_message(f"'{target}' nicht gefunden. Gib einen gültigen Pfad oder Programmnamen an.")
        
        except Exception as e:
            self.log_message(f"Fehler beim Öffnen von '{target}': {e}")
    
    def _handle_cloudflare_captcha(self) -> bool:
        self.log_message("CAPTCHA-Handling nicht möglich, Bildschirmerkennung deaktiviert. Löse CAPTCHAs manuell im Browser.")
        return False
    
    def _send_discord_message(self, target: str, message: str) -> None:
        self.log_message(f"Sende Nachricht an '{target}' auf Discord...")
        self.context.last_action = "discord"
        
        try:
            self.driver.get(self.config.discord_login_url)
            time.sleep(3)
            
            if self.config.discord_email and self.config.discord_password:
                try:
                    email_field = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, self.config.discord_email_xpath))
                    )
                    password_field = self.driver.find_element(By.XPATH, self.config.discord_password_xpath)
                    
                    email_field.send_keys(self.config.discord_email)
                    password_field.send_keys(self.config.discord_password)
                    self.driver.find_element(By.XPATH, self.config.discord_submit_xpath).click()
                    self.log_message("Logge in Discord ein... Falls ein CAPTCHA erscheint, löse es manuell.")
                    time.sleep(5)
                except Exception as e:
                    self.log_message(f"Fehler beim Discord-Login: {e}. Bitte manuell einloggen.")
            
            self.log_message(f"Automatisches Senden von Nachrichten nicht möglich, Bildschirmerkennung deaktiviert. Gib die Nachricht '{message}' an '{target}' manuell im Discord-Fenster ein.")
        except Exception as e:
            self.log_message(f"Fehler beim Senden der Nachricht: {e}")
    
    def run(self) -> None:
        self.root.mainloop()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        bot = FaceBot(root)
        bot.run()
    except Exception as e:
        print(f"Fehler: {e}")