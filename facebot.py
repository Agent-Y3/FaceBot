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
import pyautogui
import time
import re
import winreg
import subprocess
import shutil
from PIL import Image

class FaceBot:
    def __init__(self, root):
        self.root = root
        self.root.title("FaceBot")
        self.driver = None
        self.browser_name = None
        self.server_config = None  # Wird durch Benutzereingabe gesetzt
        
        # GUI-Setup
        self.chat_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=25, state='disabled')
        self.chat_area.pack(padx=10, pady=10)
        
        self.input_frame = tk.Frame(root)
        self.input_frame.pack(padx=10, pady=5, fill=tk.X)
        
        self.input_field = tk.Entry(self.input_frame)
        self.input_field.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.input_field.bind("<Return>", self.process_command)
        
        self.send_button = tk.Button(self.input_frame, text="Senden", command=self.process_command)
        self.send_button.pack(side=tk.RIGHT)
        
        # Browser initialisieren
        self.initialize_browser()
        
        # PyAutoGUI Konfiguration
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        # Pfade zu WinSCP und PuTTY (anpassen!)
        self.winscp_path = r"C:\Program Files (x86)\WinSCP\WinSCP.exe"
        self.putty_path = r"C:\Program Files\PuTTY\putty.exe"
        
        # Spotify-Konfiguration
        self.spotify_mode = "browser"  # "browser" oder "desktop"
        self.spotify_app_path = r"%userprofile%\AppData\Roaming\Spotify\Spotify.exe"  # Anpassen für Desktop-App
        self.spotify_search_image = "spotify_search.png"  # Screenshot des Suchfelds (Desktop-App)
        
        # Dateisuche-Konfiguration
        self.base_search_dir = r"%userprofile%"  # Basisverzeichnis für Dateisuche
        
        # Discord-Zugangsdaten (anpassen oder leer lassen für manuelle Eingabe)
        self.discord_config = {
            "email": "",  # E-Mail oder Benutzername
            "password": ""  # Passwort
        }
        
        # Bildpfade für Bilderkennung (anpassen!)
        self.captcha_image = "captcha_checkbox.png"
        self.discord_input_image = "discord_input.png"
        
        self.append_message(f"FaceBot: Bereit! Verwende Browser: {self.browser_name.capitalize()}. Sprich mich mit 'FaceBot' an. Befehle: virus, klick, winscp, putty, upload, discord, spiele, öffne, hilfe, beenden.")

    def prompt_server_config(self):
        """Fragt den Benutzer nach Root-Server-Daten in einem Tkinter-Fenster."""
        config_window = tk.Toplevel(self.root)
        config_window.title("Root-Server-Daten eingeben")
        config_window.geometry("400x300")
        
        # Labels und Eingabefelder
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
            self.append_message("FaceBot: Root-Server-Daten gespeichert.")
            config_window.destroy()
        
        tk.Button(config_window, text="Bestätigen", command=submit).pack(pady=10)
        config_window.transient(self.root)
        config_window.grab_set()
        self.root.wait_window(config_window)
        
        return self.server_config is not None

    def get_default_browser(self):
        """Ermittelt den Standardbrowser des Systems."""
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
    
    def initialize_browser(self):
        """Initialisiert den Selenium WebDriver für den Standardbrowser."""
        self.browser_name = self.get_default_browser()
        self.append_message(f"FaceBot: Erkenne Standardbrowser: {self.browser_name.capitalize()}.")
        
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
        except Exception as e:
            self.append_message(f"FaceBot: Fehler beim Browser-Start: {e}. Verwende Chrome.")
            service = ChromeService(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service)
            self.driver.maximize_window()
            self.browser_name = "chrome"
    
    def append_message(self, message):
        """Fügt eine Nachricht in die Chat-Area ein."""
        self.chat_area.configure(state='normal')
        self.chat_area.insert(tk.END, message + "\n")
        self.chat_area.configure(state='disabled')
        self.chat_area.see(tk.END)
        self.root.update()
    
    def process_command(self, event=None):
        """Verarbeitet Benutzereingaben."""
        command = self.input_field.get().strip()
        self.input_field.delete(0, tk.END)
        
        if not command:
            return
        
        self.append_message(f"Du: {command}")
        
        if not command.lower().startswith("facebot"):
            self.append_message("FaceBot: Bitte sprich mich mit 'FaceBot' an!")
            return
        
        cmd = re.sub(r'^facebot[,]?[\s]*', '', command, flags=re.IGNORECASE).strip().lower()
        
        if cmd in ["beenden", "exit"]:
            self.append_message("FaceBot: Beende... Tschüss!")
            self.driver.quit()
            self.root.quit()
        elif cmd in ["hilfe", "help"]:
            self.append_message("FaceBot: Befehle:\n- virus: Überprüft auf Viren\n- klick: Klickt auf aktuelle Position\n- winscp: Startet WinSCP und loggt in Root-Server\n- putty: Startet PuTTY und loggt in Root-Server\n- upload [datei]: Lädt Datei auf Root-Server\n- discord [benutzer/kanal] [nachricht]: Sendet Nachricht auf Discord\n- spiele [liedname]: Spielt ein Lied auf Spotify\n- öffne [name/pfad]: Öffnet eine Datei oder ein Programm\n- beenden: Beendet mich")
        elif cmd.startswith("virus"):
            self.check_for_viruses()
        elif cmd.startswith("klick"):
            self.perform_click()
        elif cmd.startswith("winscp"):
            if not self.server_config and not self.prompt_server_config():
                self.append_message("FaceBot: Keine Server-Daten angegeben. Abbruch.")
                return
            self.start_winscp()
        elif cmd.startswith("putty"):
            if not self.server_config and not self.prompt_server_config():
                self.append_message("FaceBot: Keine Server-Daten angegeben. Abbruch.")
                return
            self.start_putty()
        elif cmd.startswith("upload"):
            if not self.server_config and not self.prompt_server_config():
                self.append_message("FaceBot: Keine Server-Daten angegeben. Abbruch.")
                return
            file_path = cmd.replace("upload", "").strip()
            self.upload_file(file_path)
        elif cmd.startswith("discord"):
            parts = cmd.replace("discord", "").strip().split(maxsplit=1)
            if len(parts) < 2:
                self.append_message("FaceBot: Bitte gib Benutzer/Kanal und Nachricht an (z. B. 'FaceBot, discord @user Hallo').")
            else:
                target, message = parts
                self.send_discord_message(target, message)
        elif cmd.startswith("spiele"):
            song_name = cmd.replace("spiele", "").strip()
            if not song_name:
                self.append_message("FaceBot: Bitte gib einen Liednamen an (z. B. 'FaceBot, spiele Bohemian Rhapsody').")
            else:
                self.play_spotify_song(song_name)
        elif cmd.startswith("öffne"):
            target = cmd.replace("öffne", "").strip()
            if not target:
                self.append_message("FaceBot: Bitte gib einen Dateinamen, Pfad oder Programmnamen an (z. B. 'FaceBot, öffne notepad').")
            else:
                self.open_file_or_program(target)
        else:
            self.query_grok(cmd)
    
    def check_for_viruses(self):
        """Simulierte Virenprüfung."""
        self.append_message("FaceBot: Überprüfe auf Viren...")
        suspicious_processes = ['malware.exe', 'virus.exe']
        found = False
        for proc in psutil.process_iter(['name']):
            if proc.info['name'].lower() in suspicious_processes:
                self.append_message(f"FaceBot: WARNUNG: Verdächtiger Prozess: {proc.info['name']}")
                found = True
        if not found:
            self.append_message("FaceBot: Keine verdächtigen Prozesse gefunden.")
    
    def perform_click(self):
        """Simuliert einen Mausklick."""
        self.append_message("FaceBot: Klicke auf den Bildschirm!")
        x, y = pyautogui.position()
        pyautogui.moveTo(x, y, duration=0.5)
        pyautogui.click(x=x, y=y)
        self.append_message(f"FaceBot: Geklickt bei ({x}, {y})!")
    
    def start_winscp(self):
        """Startet WinSCP und loggt in Root-Server ein."""
        if not os.path.exists(self.winscp_path):
            self.append_message("FaceBot: WinSCP nicht gefunden! Bitte installiere es.")
            return
        
        self.append_message("FaceBot: Starte WinSCP und verbinde mit Root-Server...")
        try:
            if self.server_config["key_path"]:
                cmd = f'"{self.winscp_path}" sftp://{self.server_config["username"]}@{self.server_config["host"]} /privatekey="{self.server_config["key_path"]}"'
            else:
                cmd = f'"{self.winscp_path}" sftp://{self.server_config["username"]}@{self.server_config["host"]}'
            subprocess.Popen(cmd, shell=True)
            time.sleep(3)
            if self.server_config["password"] and not self.server_config["key_path"]:
                pyautogui.write(self.server_config["password"])
                pyautogui.press("enter")
            self.append_message("FaceBot: WinSCP verbunden! Du kannst jetzt Dateien übertragen.")
        except Exception as e:
            self.append_message(f"FaceBot: Fehler beim Starten von WinSCP: {e}")
    
    def start_putty(self):
        """Startet PuTTY und loggt in Root-Server ein."""
        if not os.path.exists(self.putty_path):
            self.append_message("FaceBot: PuTTY nicht gefunden! Bitte installiere es.")
            return
        
        self.append_message("FaceBot: Starte PuTTY und verbinde mit Root-Server...")
        try:
            if self.server_config["key_path"]:
                cmd = f'"{self.putty_path}" -ssh {self.server_config["username"]}@{self.server_config["host"]} -i "{self.server_config["key_path"]}"'
            else:
                cmd = f'"{self.putty_path}" -ssh {self.server_config["username"]}@{self.server_config["host"]} -pw {self.server_config["password"]}'
            subprocess.Popen(cmd, shell=True)
            self.append_message("FaceBot: PuTTY verbunden! Du kannst jetzt Shell-Befehle ausführen.")
        except Exception as e:
            self.append_message(f"FaceBot: Fehler beim Starten von PuTTY: {e}")
    
    def upload_file(self, file_path):
        """Lädt eine Datei auf den Root-Server mit WinSCP."""
        if not file_path or not os.path.exists(file_path):
            self.append_message("FaceBot: Datei nicht gefunden! Gib einen gültigen Pfad an.")
            return
        
        if not os.path.exists(self.winscp_path):
            self.append_message("FaceBot: WinSCP nicht gefunden! Bitte installiere es.")
            return
        
        self.append_message(f"FaceBot: Lade Datei '{file_path}' auf Root-Server...")
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
            
            cmd = f'"{self.winscp_path}" /script="{script_path}"'
            subprocess.run(cmd, shell=True)
            os.remove(script_path)
            self.append_message("FaceBot: Datei erfolgreich hochgeladen!")
        except Exception as e:
            self.append_message(f"FaceBot: Fehler beim Hochladen: {e}")
    
    def locate_element_on_screen(self, image_path):
        """Findet ein Element auf dem Bildschirm mit Bilderkennung."""
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=0.8)
            if location:
                x, y = pyautogui.center(location)
                self.append_message(f"FaceBot: Element gefunden bei ({x}, {y})!")
                return x, y
            self.append_message(f"FaceBot: Element '{image_path}' nicht gefunden.")
            return None
        except Exception as e:
            self.append_message(f"FaceBot: Fehler bei Bilderkennung: {e}")
            return None
    
    def play_spotify_song(self, song_name):
        """Spielt ein Lied auf Spotify ab."""
        self.append_message(f"FaceBot: Spiele '{song_name}' auf Spotify...")
        
        try:
            if self.spotify_mode == "desktop":
                if not os.path.exists(self.spotify_app_path):
                    self.append_message("FaceBot: Spotify-App nicht gefunden! Bitte überprüfe den Pfad.")
                    return
                
                # Starte Spotify-App
                subprocess.Popen(self.spotify_app_path, shell=True)
                time.sleep(5)  # Warte, bis die App geladen ist
                
                # Suche Suchfeld
                search_pos = self.locate_element_on_screen(self.spotify_search_image)
                if search_pos:
                    x, y = search_pos
                    pyautogui.moveTo(x, y, duration=0.5)
                    pyautogui.click()
                    pyautogui.write(song_name)
                    pyautogui.press("enter")
                    time.sleep(2)
                    # Klicke auf das erste Ergebnis (angenommen, es ist ca. 100 Pixel unterhalb)
                    pyautogui.moveTo(x, y + 100, duration=0.5)
                    pyautogui.click()
                    self.append_message(f"FaceBot: '{song_name}' wird abgespielt!")
                else:
                    self.append_message("FaceBot: Spotify-Suchfeld nicht gefunden. Bitte öffne Spotify manuell.")
            
            else:  # Browser-Modus
                self.driver.get("https://open.spotify.com")
                time.sleep(3)
                
                # Handle Cloudflare-Captcha, falls vorhanden
                self.handle_cloudflare_captcha()
                
                # Suche Suchfeld
                try:
                    search_button = WebDriverWait(self.driver, 10).until(
                        EC.element_to_be_clickable((By.XPATH, "//button[@data-testid='search-icon']"))
                    )
                    search_button.click()
                    
                    search_field = WebDriverWait(self.driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//input[@data-testid='search-input']"))
                    )
                    search_field.send_keys(song_name)
                    search_field.submit()
                    time.sleep(2)
                    
                    # Klicke auf das erste Ergebnis
                    first_result = WebDriverWait(self.driver, 5).until(
                        EC.element_to_be_clickable((By.XPATH, "//div[@data-testid='tracklist-row']"))
                    )
                    first_result.click()
                    self.append_message(f"FaceBot: '{song_name}' wird abgespielt!")
                except Exception as e:
                    self.append_message(f"FaceBot: Fehler beim Abspielen auf Spotify: {e}. Bitte manuell einloggen oder überprüfe Spotify.")
        
        except Exception as e:
            self.append_message(f"FaceBot: Fehler beim Abspielen auf Spotify: {e}")
    
    def open_file_or_program(self, target):
        """Öffnet eine Datei oder ein Programm."""
        self.append_message(f"FaceBot: Öffne '{target}'...")
        
        try:
            # 1. Prüfe, ob es ein absoluter Pfad ist
            if os.path.isabs(target) and os.path.exists(target):
                os.startfile(target)
                self.append_message(f"FaceBot: '{target}' geöffnet!")
                return
            
            # 2. Prüfe, ob es ein Programm in PATH oder im Startmenü ist
            program_path = shutil.which(target)
            if program_path:
                subprocess.Popen(program_path, shell=True)
                self.append_message(f"FaceBot: Programm '{target}' gestartet!")
                return
            
            # 3. Suche nach Datei in base_search_dir
            for root, _, files in os.walk(self.base_search_dir):
                if target in files:
                    file_path = os.path.join(root, target)
                    os.startfile(file_path)
                    self.append_message(f"FaceBot: Datei '{file_path}' geöffnet!")
                    return
            
            self.append_message(f"FaceBot: '{target}' nicht gefunden. Gib einen gültigen Pfad, Dateinamen oder Programmnamen an.")
        
        except Exception as e:
            self.append_message(f"FaceBot: Fehler beim Öffnen von '{target}': {e}")
    
    def handle_cloudflare_captcha(self):
        """Erkennt und klickt auf das Cloudflare-Captcha."""
        self.append_message("FaceBot: Suche nach Cloudflare-Captcha...")
        
        # Versuche Bilderkennung zuerst
        captcha_pos = self.locate_element_on_screen(self.captcha_image)
        if captcha_pos:
            x, y = captcha_pos
            self.append_message(f"FaceBot: Bewege Maus zu Captcha bei ({x}, {y})...")
            pyautogui.moveTo(x, y, duration=0.5)
            pyautogui.click()
            self.append_message("FaceBot: Auf Cloudflare-Captcha geklickt!")
            time.sleep(2)
            return True
        
        # Fallback: Selenium-basierte Erkennung
        try:
            iframe = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//iframe[contains(@src, 'cloudflare')]"))
            )
            self.append_message("FaceBot: Cloudflare-Captcha erkannt. Ermittle Position...")
            self.driver.switch_to.frame(iframe)
            
            checkbox = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, "//input[@type='checkbox']"))
            )
            
            location = checkbox.location
            size = checkbox.size
            browser_window = self.driver.get_window_position()
            x = browser_window['x'] + location['x'] + size['width'] // 2
            y = browser_window['y'] + location['y'] + size['height'] // 2 + 100
            
            self.append_message(f"FaceBot: Bewege Maus zu Captcha bei ({x}, {y})...")
            pyautogui.moveTo(x, y, duration=0.5)
            pyautogui.click()
            self.append_message("FaceBot: Auf Cloudflare-Captcha geklickt!")
            
            self.driver.switch_to.default_content()
            time.sleep(2)
            return True
        except Exception as e:
            self.append_message(f"FaceBot: Kein Cloudflare-Captcha gefunden oder Fehler: {e}")
            self.driver.switch_to.default_content()
            return False
    
    def send_discord_message(self, target, message):
        """Loggt sich in Discord ein und sendet eine Nachricht."""
        self.append_message(f"FaceBot: Öffne Discord und sende Nachricht an '{target}'...")
        
        try:
            self.driver.get("https://discord.com/login")
            time.sleep(3)
            
            # Handle Cloudflare-Captcha, falls vorhanden
            self.handle_cloudflare_captcha()
            
            # Logge dich ein (manuelle Eingabe oder gespeicherte Zugangsdaten)
            if self.discord_config["email"] and self.discord_config["password"]:
                try:
                    email_field = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.NAME, "email"))
                    )
                    password_field = self.driver.find_element(By.NAME, "password")
                    
                    email_field.send_keys(self.discord_config["email"])
                    password_field.send_keys(self.discord_config["password"])
                    self.driver.find_element(By.XPATH, "//button[@type='submit']").click()
                    self.append_message("FaceBot: Logge in Discord ein...")
                    time.sleep(5)  # Warte auf Login und mögliche 2FA
                except Exception as e:
                    self.append_message(f"FaceBot: Fehler beim Discord-Login: {e}. Bitte manuell einloggen.")
            
            # Navigiere zu Kanal/Benutzer (angenommen, du bist eingeloggt)
            self.append_message(f"FaceBot: Suche Eingabefeld für Nachricht an '{target}'...")
            input_pos = self.locate_element_on_screen(self.discord_input_image)
            if input_pos:
                x, y = input_pos
                pyautogui.moveTo(x, y, duration=0.5)
                pyautogui.click()
                pyautogui.write(f"{target}: {message}")
                pyautogui.press("enter")
                self.append_message(f"FaceBot: Nachricht '{message}' an '{target}' gesendet!")
            else:
                self.append_message("FaceBot: Discord-Eingabefeld nicht gefunden. Bitte öffne Discord manuell.")
        except Exception as e:
            self.append_message(f"FaceBot: Fehler beim Senden der Discord-Nachricht: {e}")
    
    def open_grok(self):
        """Öffnet grok.com im Browser und behandelt Cloudflare-Captcha."""
        self.append_message(f"FaceBot: Öffne grok.com in {self.browser_name.capitalize()}...")
        self.driver.get("https://grok.com")
        time.sleep(2)
        
        if self.handle_cloudflare_captcha():
            self.append_message("FaceBot: Captcha behandelt. Seite geladen.")
        else:
            self.append_message("FaceBot: Kein Captcha erkannt. Klicke auf die Seite.")
            pyautogui.moveTo(500, 500, duration=0.5)
            pyautogui.click()
    
    def query_grok(self, command):
        """Fragt Grok nach dem passenden Befehl."""
        self.append_message(f"FaceBot: Unbekannter Befehl '{command}'. Frage Grok...")
        try:
            self.driver.get("https://grok.com")
            time.sleep(2)
            
            if self.handle_cloudflare_captcha():
                self.append_message("FaceBot: Captcha behandelt. Suche Grok-Eingabefeld...")
            
            try:
                search_box = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.ID, "search"))
                )
                search_box.send_keys(f"Welcher Befehl passt zu '{command}'? Verfügbare Befehle: virus, klick, winscp, putty, upload, discord, spiele, öffne, hilfe, beenden")
                search_box.submit()
                time.sleep(2)
                self.append_message("FaceBot: Grok abgefragt. Überprüfe die Seite für Vorschläge.")
            except:
                self.append_message("FaceBot: Grok-Suchfeld nicht gefunden. Versuche Captcha erneut...")
                if self.handle_cloudflare_captcha():
                    self.append_message("FaceBot: Captcha behandelt. Seite geladen.")
                else:
                    self.append_message("FaceBot: Kein Captcha oder Suchfeld gefunden. Klicke auf die Seite.")
                    pyautogui.moveTo(500, 500, duration=0.5)
                    pyautogui.click()
        except Exception as e:
            self.append_message(f"FaceBot: Fehler bei Grok-Abfrage: {e}")
            if self.handle_cloudflare_captcha():
                self.append_message("FaceBot: Captcha behandelt als Fallback.")
            else:
                pyautogui.moveTo(500, 500, duration=0.5)
                pyautogui.click()
    
    def run(self):
        """Startet die GUI."""
        self.root.mainloop()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        bot = FaceBot(root)
        bot.run()
    except Exception as e:
        print(f"Fehler: {e}")