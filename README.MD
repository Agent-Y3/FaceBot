# FaceBot

FaceBot ist ein vielseitiger, automatisierter Assistent mit einer Tkinter-basierten GUI, der Aufgaben wie das Abspielen von Musik auf Spotify, das Öffnen von Dateien/Programmen, die Interaktion mit Root-Servern (über WinSCP/PuTTY), das Senden von Discord-Nachrichten und die Handhabung von Cloudflare-Captchas ermöglicht. Der Bot nutzt Selenium für Browser-Automatisierung und PyAutoGUI für Bildschirmsteuerung.

## Funktionen

- **Spotify-Wiedergabe**: Spielt Lieder über `https://open.spotify.com/search/[Liedname]` im Browser ab.
- **Datei-/Programm-Öffnen**: Öffnet Dateien oder Programme via absoluten Pfad, PATH oder Suche im Basisverzeichnis.
- **Root-Server-Management**: Verbindet sich über WinSCP/PuTTY und lädt Dateien hoch (unterstützt Passwort- und Schlüssel-Authentifizierung).
- **Discord-Nachrichten**: Sendet Nachrichten an Benutzer/Kanäle im Browser.
- **Cloudflare-Captcha-Handhabung**: Erkennt und klickt automatisch auf Checkbox-Captchas.
- **Virenprüfung**: Simulierte Überprüfung auf verdächtige Prozesse.
- **Maussteuerung**: Simuliert Mausklicks an der aktuellen Position.
- **Grok-Abfragen**: Leitet unbekannte Befehle an grok.com weiter.

## Voraussetzungen

- **Betriebssystem**: Windows (für `os.startfile`, WinSCP, PuTTY).
- **Python**: Version 3.8 oder höher.
- **Browser**: Chrome, Firefox oder Edge.
- **Software**: WinSCP und PuTTY installiert.
- **Bildschirmauflösung**: Kompatibel mit Screenshots für Bilderkennung.

## Installation

1. **Python-Bibliotheken installieren**:
   ```bash
   pip install tkinter selenium psutil pyautogui webdriver-manager requests pillow
   ```

2. **WinSCP und PuTTY installieren**:
   - Lade und installiere [WinSCP](https://winscp.net/) und [PuTTY](https://www.putty.org/).
   - Überprüfe die Pfade in der `Config`-Klasse (Standard: `C:\Program Files (x86)\WinSCP\WinSCP.exe` und `C:\Program Files\PuTTY\putty.exe`).

3. **Screenshots für Bilderkennung vorbereiten**:
   - Erstelle Screenshots von:
     - Cloudflare-Captcha-Checkbox (`captcha_checkbox.png`).
     - Discord-Nachrichten-Eingabefeld (`discord_input.png`).
   - Speichere sie im gleichen Verzeichnis wie `facebot.py` (z. B. `C:\Users\xByYu\Downloads`).
   - Passe die Pfade in `Config` an, falls nötig.

4. **Code herunterladen**:
   - Speichere `facebot.py` in einem Verzeichnis deiner Wahl (z. B. `C:\Users\xByYu\Downloads`).

## Konfiguration

Bearbeite die `Config`-Klasse in `facebot.py`, um Pfade und Einstellungen anzupassen:

```python
@dataclass
class Config:
    winscp_path: str = r"C:\Program Files (x86)\WinSCP\WinSCP.exe"
    putty_path: str = r"C:\Program Files\PuTTY\putty.exe"
    base_search_dir: str = r"C:\Users\xByYu"  # Basisverzeichnis für Dateisuche
    discord_email: str = ""  # Discord E-Mail oder Benutzername
    discord_password: str = ""  # Discord Passwort
    captcha_image: str = "captcha_checkbox.png"
    discord_input_image: str = "discord_input.png"
```

- **Spotify**: Stelle sicher, dass du in Spotify im Browser eingeloggt bist (`https://open.spotify.com`).
- **Discord**: Gib `discord_email` und `discord_password` ein oder lasse leer für manuelles Login.
- **Root-Server**: Daten (Host, Benutzername, Passwort/Schlüssel) werden bei Bedarf abgefragt.

## Nutzung

1. **Skript starten**:
   ```bash
   python facebot.py
   ```

2. **GUI-Bedienung**:
   - Gib Befehle in das Eingabefeld ein, beginnend mit `FaceBot`.
   - Beispiel: `FaceBot, spiele Shape of You`.

3. **Verfügbare Befehle**:
   - `virus`: Simulierte Virenprüfung.
   - `klick`: Klickt an der aktuellen Mausposition.
   - `winscp`: Startet WinSCP und verbindet sich mit dem Root-Server.
   - `putty`: Startet PuTTY und verbindet sich mit dem Root-Server.
   - `upload [Pfad]`: Lädt eine Datei auf den Root-Server (z. B. `upload C:\file.txt`).
   - `discord [Ziel] [Nachricht]`: Sendet eine Nachricht auf Discord (z. B. `discord @user Hallo`).
   - `spiele [Liedname]`: Spielt ein Lied auf Spotify (z. B. `spiele Bohemian Rhapsody`).
   - `öffne [Name/Pfad]`: Öffnet eine Datei oder ein Programm (z. B. `öffne notepad`, `öffne C:\file.txt`).
   - `hilfe`: Zeigt alle Befehle.
   - `beenden`: Schließt den Bot.

## Beispiele

- **Spotify**:
  ```text
  FaceBot, spiele Shape of You
  ```
  Öffnet `https://open.spotify.com/search/Shape%20of%20You` und spielt das erste Ergebnis.

- **Datei öffnen**:
  ```text
  FaceBot, öffne C:\Users\xByYu\Documents\report.pdf
  ```
  Öffnet `report.pdf` im Standardprogramm.

- **Programm starten**:
  ```text
  FaceBot, öffne firefox
  ```
  Startet Firefox, wenn installiert.

- **Discord-Nachricht**:
  ```text
  FaceBot, discord @friend Hallo
  ```
  Sendet "Hallo" an `@friend`.

## Fehlerbehebung

- **Spotify-Wiedergabe schlägt fehl**:
  - Stelle sicher, dass du in Spotify im Browser eingeloggt bist.
  - Überprüfe, ob das Lied verfügbar ist.
  - Prüfe Konsolenausgaben für Fehlerdetails.


- **Discord-Login-Probleme**:
  - Gib `discord_email` und `discord_password` in `Config` ein oder logge dich manuell ein.
  - Bei 2FA: Gib den Code manuell ein.

- **Allgemeine Fehler**:
  - Starte das Skript als Administrator für `pyautogui` oder Programme.
  - Überprüfe Konsolenausgaben (Logging) für detaillierte Fehlermeldungen.
  - Bewege die Maus in die obere linke Ecke, um `pyautogui`-Aktionen zu stoppen.

## Einschränkungen

- **Plattform**: Optimiert für Windows. Für andere Betriebssysteme sind Anpassungen nötig.
- **Captcha**: Unterstützt nur Checkbox-Captchas. Bildbasierte Captchas erfordern zusätzliche Logik.
- **Dateisuche**: Kann bei großen Verzeichnissen langsam sein. Nutze absolute Pfade für schnellere Ergebnisse.
- **Sicherheit**: Zugangsdaten werden im Speicher gehalten. Verwende bei sensiblen Daten Vorsicht.

## Erweiterungen

Mögliche zukünftige Verbesserungen:
- Unterstützung für Spotify-Playlists/Alben.
- OpenCV für robustere Bilderkennung.
- Persistentes Speichern von Root-Server-Daten.
- Unterstützung für Discord-Desktop-App.
- Handhabung von Bild-Captchas.


## Kontakt

Bei Fragen oder Problemen kontaktiere mich über Discord: yu_dino oder öffne ein Issue im Repository.

## Proof that ist not a Virus

http://46.232.250.103/VirenTest.jpg