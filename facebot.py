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
from urllib.parse import quote
import logging
from typing import Optional, Dict, Tuple
from dataclasses import dataclass

@dataclass
class Config:
    winscp_path: str = r"C:\Program Files (x86)\WinSCP\WinSCP.exe"
    putty_path: str = r"C:\Program Files\PuTTY\putty.exe"
    base_search_dir: str = r"C:\Users\xByYu"
    discord_email: str = ""
    discord_password: str = ""
    captcha_image: str = "captcha_checkbox.png"
    discord_input_image: str = "discord_input.png"
    spotify_search_url: str = "https://open.spotify.com/search/{}"
    discord_login_url: str = "https://discord.com/login"
    grok_url: str = "https://grok.com"
    tracklist_xpath: str = "//div[@data-testid='tracklist-row']"
    captcha_iframe_xpath: str = "//iframe[contains(@src, 'cloudflare')]"
    captcha_checkbox_xpath: str = "//input[@type='checkbox']"
    discord_email_xpath: str = "//input[@name='email']"
    discord_password_xpath: str = "//input[@name='password']"
    discord_submit_xpath: str = "//button[@type='submit']"
    grok_search_xpath: str = "//input[@id='search']"

class FaceBot:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("FaceBot")
        self.driver = None
        self.browser_name = None
        self.server_config = None
        self.config = Config()
        self.logger = self._setup_logger()
        
        self.chat_area = scrolledtext.ScrolledText(root, wrap=tk.WORD, width=60, height=25, state='disabled')
        self.chat_area.pack(padx=10, pady=10)
        
        self.input_frame = tk.Frame(root)
        self.input_frame.pack(padx=10, pady=5, fill=tk.X)
        
        self.input_field = tk.Entry(self.input_frame)
        self.input_field.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.input_field.bind("<Return>", self.process_command)
        
        self.send_button = tk.Button(self.input_frame, text="Senden", command=self.process_command)
        self.send_button.pack(side=tk.RIGHT)
        
        self._initialize_browser()
        
        pyautogui.FAILSAFE = True
        pyautogui.PAUSE = 0.5
        
        self.log_message(f"Bereit! Verwende Browser: {self.browser_name.capitalize()}. Sprich mich mit 'FaceBot' an. Befehle: virus, klick, winscp, putty, upload, discord, spiele, öffne, hilfe, beenden.")

    def _setup_logger(self) -> logging.Logger:
        logger = logging.getLogger("FaceBot")
        logger.setLevel(logging.INFO)
        handler = logging.StreamHandler()
        handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
        logger.addHandler(handler)
        return logger

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
        self.log_message(f"Erkenne Standardbrowser: {self.browser_name.capitalize()}.")
        
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
            self.log_message(f"Fehler beim Browser-Start: {e}. Verwende Chrome.")
            service = ChromeService(ChromeDriverManager().install())
            self.driver = webdriver.Chrome(service=service)
            self.driver.maximize_window()
            self.browser_name = "chrome"
    
    def log_message(self, message: str) -> None:
        self.chat_area.configure(state='normal')
        self.chat_area.insert(tk.END, f"{message}\n")
        self.chat_area.configure(state='disabled')
        self.chat_area.see(tk.END)
        self.root.update()
        self.logger.info(message)
    
    def process_command(self, event: Optional[tk.Event] = None) -> None:
        command = self.input_field.get().strip()
        self.input_field.delete(0, tk.END)
        
        if not command:
            return
        
        self.log_message(f"Du: {command}")
        
        if not command.lower().startswith("facebot"):
            self.log_message("Bitte sprich mich mit 'FaceBot' an!")
            return
        
        cmd = re.sub(r'^facebot[,]?[\s]*', '', command, flags=re.IGNORECASE).strip().lower()
        
        if cmd in ["beenden", "exit"]:
            self.log_message("Beende... Tschüss!")
            self.driver.quit()
            self.root.quit()
        elif cmd in ["hilfe", "help"]:
            self.log_message("Befehle:\n- virus: Überprüft auf Viren\n- klick: Klickt auf aktuelle Position\n- winscp: Startet WinSCP und loggt in Root-Server\n- putty: Startet PuTTY und loggt in Root-Server\n- upload [datei]: Lädt Datei auf Root-Server\n- discord [benutzer/kanal] [nachricht]: Sendet Nachricht auf Discord\n- spiele [liedname]: Spielt ein Lied auf Spotify\n- öffne [name/pfad]: Öffnet eine Datei oder ein Programm\n- beenden: Beendet mich")
        elif cmd.startswith("virus"):
            self._check_for_viruses()
        elif cmd.startswith("klick"):
            self._perform_click()
        elif cmd.startswith("winscp"):
            if not self.server_config and not self._prompt_server_config():
                self.log_message("Keine Server-Daten angegeben. Abbruch.")
                return
            self._start_winscp()
        elif cmd.startswith("putty"):
            if not self.server_config and not self._prompt_server_config():
                self.log_message("Keine Server-Daten angegeben. Abbruch.")
                return
            self._start_putty()
        elif cmd.startswith("upload"):
            if not self.server_config and not self._prompt_server_config():
                self.log_message("Keine Server-Daten angegeben. Abbruch.")
                return
            file_path = cmd.replace("upload", "").strip()
            self._upload_file(file_path)
        elif cmd.startswith("discord"):
            parts = cmd.replace("discord", "").strip().split(maxsplit=1)
            if len(parts) < 2:
                self.log_message("Bitte gib Benutzer/Kanal und Nachricht an (z. B. 'FaceBot, discord @user Hallo').")
            else:
                target, message = parts
                self._send_discord_message(target, message)
        elif cmd.startswith("spiele"):
            song_name = cmd.replace("spiele", "").strip()
            if not song_name:
                self.log_message("Bitte gib einen Liednamen an (z. B. 'FaceBot, spiele Bohemian Rhapsody').")
            else:
                self._play_spotify_song(song_name)
        elif cmd.startswith("öffne"):
            target = cmd.replace("öffne", "").strip()
            if not target:
                self.log_message("Bitte gib einen Dateinamen, Pfad oder Programmnamen an (z. B. 'FaceBot, öffne notepad').")
            else:
                self._open_file_or_program(target)
        else:
            self._query_grok(cmd)
    
    def _check_for_viruses(self) -> None:
        self.log_message("Überprüfe auf Viren...")
        suspicious_processes = ['malware.exe', 'virus.exe']
        found = False
        for proc in psutil.process_iter(['name']):
            if proc.info['name'].lower() in suspicious_processes:
                self.log_message(f"WARNUNG: Verdächtiger Prozess: {proc.info['name']}")
                found = True
        if not found:
            self.log_message("Keine verdächtigen Prozesse gefunden.")
    
    def _perform_click(self) -> None:
        self.log_message("Klicke auf den Bildschirm!")
        x, y = pyautogui.position()
        pyautogui.moveTo(x, y, duration=0.5)
        pyautogui.click(x=x, y=y)
        self.log_message(f"Geklickt bei ({x}, {y})!")
    
    def _start_winscp(self) -> None:
        if not os.path.exists(self.config.winscp_path):
            self.log_message("WinSCP nicht gefunden! Bitte installiere es.")
            return
        
        self.log_message("Starte WinSCP und verbinde mit Root-Server...")
        try:
            if self.server_config["key_path"]:
                cmd = f'"{self.config.winscp_path}" sftp://{self.server_config["username"]}@{self.server_config["host"]} /privatekey="{self.server_config["key_path"]}"'
            else:
                cmd = f'"{self.config.winscp_path}" sftp://{self.server_config["username"]}@{self.server_config["host"]}'
            subprocess.Popen(cmd, shell=True)
            time.sleep(3)
            if self.server_config["password"] and not self.server_config["key_path"]:
                pyautogui.write(self.server_config["password"])
                pyautogui.press("enter")
            self.log_message("WinSCP verbunden! Du kannst jetzt Dateien übertragen.")
        except Exception as e:
            self.log_message(f"Fehler beim Starten von WinSCP: {e}")
    
    def _start_putty(self) -> None:
        if not os.path.exists(self.config.putty_path):
            self.log_message("PuTTY nicht gefunden! Bitte installiere es.")
            return
        
        self.log_message("Starte PuTTY und verbinde mit Root-Server...")
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
        
        self.log_message(f"Lade Datei '{file_path}' auf Root-Server...")
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
    
    def _locate_element_on_screen(self, image_path: str) -> Optional[Tuple[int, int]]:
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=0.8)
            if location:
                x, y = pyautogui.center(location)
                self.log_message(f"Element gefunden bei ({x}, {y})!")
                return x, y
            self.log_message(f"Element '{image_path}' nicht gefunden.")
            return None
        except Exception as e:
            self.log_message(f"Fehler bei Bilderkennung: {e}")
            return None
    
    def _play_spotify_song(self, song_name: str) -> None:
        self.log_message(f"Spiele '{song_name}' auf Spotify...")
        
        try:
            encoded_song = quote(song_name)
            search_url = self.config.spotify_search_url.format(encoded_song)
            self.driver.get(search_url)
            time.sleep(3)
            
            self._handle_cloudflare_captcha()
            
            try:
                first_result = WebDriverWait(self.driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, self.config.tracklist_xpath))
                )
                first_result.click()
                self.log_message(f"'{song_name}' wird abgespielt!")
            except Exception as e:
                self.log_message(f"Fehler beim Abspielen auf Spotify: {e}. Bitte überprüfe, ob du eingeloggt bist oder ob das Lied verfügbar ist.")
        
        except Exception as e:
            self.log_message(f"Fehler beim Öffnen von Spotify: {e}")
    
    def _open_file_or_program(self, target: str) -> None:
        self.log_message(f"Öffne '{target}'...")
        
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
            
            self.log_message(f"'{target}' nicht gefunden. Gib einen gültigen Pfad, Dateinamen oder Programmnamen an.")
        
        except Exception as e:
            self.log_message(f"Fehler beim Öffnen von '{target}': {e}")
    
    def _handle_cloudflare_captcha(self) -> bool:
        self.log_message("Suche nach Cloudflare-Captcha...")
        
        captcha_pos = self._locate_element_on_screen(self.config.captcha_image)
        if captcha_pos:
            x, y = captcha_pos
            self.log_message(f"Bewege Maus zu Captcha bei ({x}, {y})...")
            pyautogui.moveTo(x, y, duration=0.5)
            pyautogui.click()
            self.log_message("Auf Cloudflare-Captcha geklickt!")
            time.sleep(2)
            return True
        
        try:
            iframe = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.XPATH, self.config.captcha_iframe_xpath))
            )
            self.log_message("Cloudflare-Captcha erkannt. Ermittle Position...")
            self.driver.switch_to.frame(iframe)
            
            checkbox = WebDriverWait(self.driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, self.config.captcha_checkbox_xpath))
            )
            
            location = checkbox.location
            size = checkbox.size
            browser_window = self.driver.get_window_position()
            x = browser_window['x'] + location['x'] + size['width'] // 2
            y = browser_window['y'] + location['y'] + size['height'] // 2 + 100
            
            self.log_message(f"Bewege Maus zu Captcha bei ({x}, {y})...")
            pyautogui.moveTo(x, y, duration=0.5)
            pyautogui.click()
            self.log_message("Auf Cloudflare-Captcha geklickt!")
            
            self.driver.switch_to.default_content()
            time.sleep(2)
            return True
        except Exception as e:
            self.log_message(f"Kein Cloudflare-Captcha gefunden oder Fehler: {e}")
            self.driver.switch_to.default_content()
            return False
    
    def _send_discord_message(self, target: str, message: str) -> None:
        self.log_message(f"Öffne Discord und sende Nachricht an '{target}'...")
        
        try:
            self.driver.get(self.config.discord_login_url)
            time.sleep(3)
            
            self._handle_cloudflare_captcha()
            
            if self.config.discord_email and self.config.discord_password:
                try:
                    email_field = WebDriverWait(self.driver, 10).until(
                        EC.presence_of_element_located((By.XPATH, self.config.discord_email_xpath))
                    )
                    password_field = self.driver.find_element(By.XPATH, self.config.discord_password_xpath)
                    
                    email_field.send_keys(self.config.discord_email)
                    password_field.send_keys(self.config.discord_password)
                    self.driver.find_element(By.XPATH, self.config.discord_submit_xpath).click()
                    self.log_message("Logge in Discord ein...")
                    time.sleep(5)
                except Exception as e:
                    self.log_message(f"Fehler beim Discord-Login: {e}. Bitte manuell einloggen.")
            
            self.log_message(f"Suche Eingabefeld für Nachricht an '{target}'...")
            input_pos = self._locate_element_on_screen(self.config.discord_input_image)
            if input_pos:
                x, y = input_pos
                pyautogui.moveTo(x, y, duration=0.5)
                pyautogui.click()
                pyautogui.write(f"{target}: {message}")
                pyautogui.press("enter")
                self.log_message(f"Nachricht '{message}' an '{target}' gesendet!")
            else:
                self.log_message("Discord-Eingabefeld nicht gefunden. Bitte öffne Discord manuell.")
        except Exception as e:
            self.log_message(f"Fehler beim Senden der Discord-Nachricht: {e}")
    
    def _open_grok(self) -> None:
        self.log_message(f"Öffne grok.com in {self.browser_name.capitalize()}...")
        self.driver.get(self.config.grok_url)
        time.sleep(2)
        
        if self._handle_cloudflare_captcha():
            self.log_message("Captcha behandelt. Seite geladen.")
        else:
            self.log_message("Kein Captcha erkannt. Klicke auf die Seite.")
            pyautogui.moveTo(500, 500, duration=0.5)
            pyautogui.click()
    
    def _query_grok(self, command: str) -> None:
        self.log_message(f"Unbekannter Befehl '{command}'. Frage Grok...")
        try:
            self.driver.get(self.config.grok_url)
            time.sleep(2)
            
            if self._handle_cloudflare_captcha():
                self.log_message("Captcha behandelt. Suche Grok-Eingabefeld...")
            
            try:
                search_box = WebDriverWait(self.driver, 5).until(
                    EC.presence_of_element_located((By.XPATH, self.config.grok_search_xpath))
                )
                search_box.send_keys(f"Welcher Befehl passt zu '{command}'? Verfügbare Befehle: virus, klick, winscp, putty, upload, discord, spiele, öffne, hilfe, beenden")
                search_box.submit()
                time.sleep(2)
                self.log_message("Grok abgefragt. Überprüfe die Seite für Vorschläge.")
            except:
                self.log_message("Grok-Suchfeld nicht gefunden. Versuche Captcha erneut...")
                if self._handle_cloudflare_captcha():
                    self.log_message("Captcha behandelt. Seite geladen.")
                else:
                    self.log_message("Kein Captcha oder Suchfeld gefunden. Klicke auf die Seite.")
                    pyautogui.moveTo(500, 500, duration=0.5)
                    pyautogui.click()
        except Exception as e:
            self.log_message(f"Fehler bei Grok-Abfrage: {e}")
            if self._handle_cloudflare_captcha():
                self.log_message("Captcha behandelt als Fallback.")
            else:
                pyautogui.moveTo(500, 500, duration=0.5)
                pyautogui.click()
    
    def run(self) -> None:
        self.root.mainloop()

if __name__ == "__main__":
    try:
        root = tk.Tk()
        bot = FaceBot(root)
        bot.run()
    except Exception as e:
        print(f"Fehler: {e}")