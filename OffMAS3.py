"""
Office Source Downloader & Installer
- Télécharge les sources Office via YAOCTRU Generator
- Configure l'installation via YAOCTRIR Configurator
- Active Office via Ohook (optionnel)
- Désinstalle Office proprement
- Interface sombre avec cartes et sélecteurs en cascade
- Support aria2c pour téléchargement multi-connexions
"""

import subprocess
import threading
import re
import os
import sys
import shutil
import time
import datetime
import configparser
import traceback
import glob
import ctypes
from ctypes import wintypes
import customtkinter as ctk
from tkinter import filedialog
import tkinter.messagebox as mb
import webbrowser

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# ──────────────────────────────────────────────
# Debug
# ──────────────────────────────────────────────

DEBUG = True


def dbg(msg: str, level: str = "INFO"):
    if not DEBUG:
        return
    ts = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3]
    tag = {"INFO": "ℹ", "WARN": "⚠", "ERR": "✗", "OK": "✓", "STEP": "►"}.get(level, "·")
    try:
        print(f"[{ts}] {tag} {msg}", flush=True)
    except Exception:
        pass


# ──────────────────────────────────────────────
# Configuration par défaut
# ──────────────────────────────────────────────

DEFAULT_CONFIG = {
    "General": {
        "auto_activate": "True",
        "auto_install_after_download": "False",
    },
    "Office": {
        "default_channel": "3: Current / Monthly",
        "default_build": "1: Windows 11/10 (Latest)",
        "default_bitness": "2: 64-bit (x64)",
        "default_language": "11: fr-FR (Français)",
        "default_suite": "O365HomePremRetail",
        "default_apps": "Word,Excel,PowerPoint",
    },
}

CONFIG_FILENAME = "office_tool_config.ini"

# ──────────────────────────────────────────────
# Constantes YAOCTRU (Generator)
# ──────────────────────────────────────────────

CHANNELS_YAOCTRU = {
    "1: Beta / Insider Fast": 1,
    "2: Current / Monthly Preview": 2,
    "3: Current / Monthly": 3,
    "4: Monthly Enterprise": 4,
    "6: Semi-Annual Enterprise": 6,
    "7: DevMain Channel": 7,
    "8: Microsoft Elite": 8,
    "9: Perpetual2019 VL": 9,
    "10: Microsoft2019 VL": 10,
    "11: Perpetual2021 VL": 11,
    "12: Microsoft2021 VL": 12,
    "13: Perpetual2024 VL": 13,
    "14: Microsoft2024 VL": 14,
}

BUILDS = {
    "1: Windows 11/10 (Latest)": 1,
    "2: Windows 8.1": 2,
    "3: Windows 7": 3,
}

BITNESS = {
    "1: 32-bit (x86)": 1,
    "2: 64-bit (x64)": 2,
    "3: Dual (x64 and x86)": 3,
    "4: ARM64 (x86 Emulation)": 4,
    "5: ARM64 (x64 Emulation)": 5,
}

LANGUAGES = {
    "11: fr-FR (Français)": 11,
    "01: en-US (Anglais US)": 1,
    "06: de-DE (Allemand)": 6,
    "08: es-ES (Espagnol)": 8,
}

DL_TYPES = {
    "1: Full Office Source": 1,
    "2: Language Pack": 2,
    "3: Proofing Tools": 3,
}

DL_OUTPUTS = {"1: Aria2 script": 1}
OUTPUT_SUFFIXES = {1: "_aria2.bat"}

# ──────────────────────────────────────────────
# Constantes YAOCTRIR (Configurator / Suites)
# ──────────────────────────────────────────────

SUITES = {
    "Microsoft 365": {
        "Famille (Home Premium)": "O365HomePremRetail",
        "Entreprise (ProPlus)": "O365ProPlusRetail",
        "Business (Business)": "O365BusinessRetail",
        "Small Business": "O365SmallBusPremRetail",
        "Education": "O365EduCloudRetail",
    },
    "Office 2024": {
        "Pro Plus 2024": "ProPlus2024Retail",
        "Professional 2024": "Professional2024Retail",
        "Standard 2024": "Standard2024Retail",
        "Home & Business 2024": "HomeBusiness2024Retail",
        "Home & Student 2024": "Home2024Retail",
    },
    "Office 2021": {
        "Pro Plus 2021": "ProPlus2021Retail",
        "Professional 2021": "Professional2021Retail",
        "Standard 2021": "Standard2021Retail",
        "Home & Business 2021": "HomeBusiness2021Retail",
        "Home & Student 2021": "HomeStudent2021Retail",
    },
    "Office 2019": {
        "Pro Plus 2019": "ProPlus2019Retail",
        "Standard 2019": "Standard2019Retail",
    },
}

APPS_MAP = {
    "Word": "Word", "Excel": "Excel", "PowerPoint": "PowerPoint",
    "Outlook": "Outlook", "OneNote": "OneNote", "Access": "Access",
    "Publisher": "Publisher", "Skype/Lync": "Lync", "OneDrive": "OneDrive",
    "Teams": "Teams", "Groove": "Groove",
}

CHANNELS_YAOCTRIR = [
    "Monthly", "MonthlyPreview", "Broad", "Targeted", "Beta", "Dogfood",
    "PerpetualVL2019", "PerpetualVL2021", "PerpetualVL2024",
]

# ──────────────────────────────────────────────
# Palette de couleurs
# ──────────────────────────────────────────────

COL = {
    "bg_app": "#0a0a0a",
    "bg_card": "#141414",
    "bg_card_alt": "#1a1a1a",
    "bg_input": "#1e1e1e",
    "bg_result": "#111111",
    "border": "#2a2a2a",
    "border_light": "#333333",
    "text_primary": "#e0e0e0",
    "text_secondary": "#999999",
    "text_muted": "#666666",
    "text_dim": "#555555",
    "accent_blue": "#4a90d9",
    "accent_blue_h": "#3a7bc8",
    "accent_green": "#5cb85c",
    "accent_green_h": "#4a9a4a",
    "accent_amber": "#d4a843",
    "accent_amber_h": "#b8922e",
    "accent_red": "#d9534f",
    "accent_red_h": "#c9433f",
    "accent_purple": "#7c6fbf",
    "accent_purple_h": "#6a5daa",
    "accent_orange": "#e67e22",
    "accent_orange_h": "#d35400",
    "accent_teal": "#5bc0be",
    "progress_bg": "#1e1e1e",
    "progress_fill": "#5cb85c",
    "btn_neutral": "#2a2a2a",
    "btn_neutral_h": "#383838",
    "status_ok": "#5cb85c",
    "status_warn": "#d4a843",
    "status_err": "#d9534f",
    "status_info": "#4a90d9",
}

ARIA2_CONNECTIONS = 16
ARIA2_SPLIT = 16


# ──────────────────────────────────────────────
# Utilitaires
# ──────────────────────────────────────────────

def get_config_path() -> str:
    p = os.path.join(os.path.dirname(os.path.abspath(__file__)), CONFIG_FILENAME)
    dbg(f"get_config_path() -> {p}")
    return p


def load_config() -> configparser.ConfigParser:
    dbg("load_config() appelé")
    cfg = configparser.ConfigParser()
    p = get_config_path()
    for sec, vals in DEFAULT_CONFIG.items():
        if not cfg.has_section(sec):
            cfg.add_section(sec)
        for k, v in vals.items():
            cfg.set(sec, k, str(v))
    if os.path.isfile(p):
        dbg(f"  Fichier config trouvé : {p}", "OK")
        cfg.read(p, encoding="utf-8")
        modified = False
        for sec, vals in DEFAULT_CONFIG.items():
            if not cfg.has_section(sec):
                cfg.add_section(sec)
                modified = True
            for k, v in vals.items():
                if not cfg.has_option(sec, k):
                    cfg.set(sec, k, str(v))
                    modified = True
                    dbg(f"  Clé manquante ajoutée : [{sec}] {k} = {v}", "WARN")
        if modified:
            save_config(cfg)
    else:
        dbg(f"  Fichier config absent, création : {p}", "WARN")
        save_config(cfg)
    for sec in cfg.sections():
        for k, v in cfg.items(sec):
            dbg(f"  Config [{sec}] {k} = {v}")
    return cfg


def save_config(cfg: configparser.ConfigParser) -> None:
    p = get_config_path()
    dbg(f"save_config() -> {p}")
    try:
        with open(p, "w", encoding="utf-8") as f:
            cfg.write(f)
        dbg("  Config sauvegardée", "OK")
    except Exception as e:
        dbg(f"  Erreur sauvegarde config : {e}", "ERR")


def format_size(size_bytes) -> str:
    if size_bytes is None or size_bytes < 0:
        return "? B"
    b = float(size_bytes)
    if b < 1024:
        return f"{int(b)} B"
    elif b < 1024 ** 2:
        return f"{b / 1024:.1f} KB"
    elif b < 1024 ** 3:
        return f"{b / 1024 ** 2:.1f} MB"
    else:
        return f"{b / 1024 ** 3:.2f} GB"

def find_script(name: str) -> str | None:
    dbg(f"find_script('{name}') recherche…")
    sd = os.path.dirname(os.path.abspath(__file__))
    for p in [
        os.path.join(sd, "Downloads", name),
        os.path.join(sd, name),
        os.path.join(os.getcwd(), "Downloads", name),
        os.path.join(os.getcwd(), name),
    ]:
        if os.path.exists(p):
            dbg(f"  Trouvé : {p}", "OK")
            return p
    dbg(f"  '{name}' NON TROUVÉ dans les chemins candidats", "WARN")
    return None


def _get_work_dir(script_path: str | None) -> str:
    if script_path:
        d = os.path.dirname(script_path)
        c2r = os.path.join(d, "C2R_Monthly")
        if os.path.isdir(c2r):
            dbg(f"_get_work_dir() -> {d} (C2R_Monthly trouvé à côté du script)")
            return d
        sd = os.path.dirname(os.path.abspath(__file__))
        dbg(f"_get_work_dir() -> {sd} (fallback script dir)")
        return sd
    sd = os.path.dirname(os.path.abspath(__file__))
    dbg(f"_get_work_dir(None) -> {sd}")
    return sd

def get_installed_office_info() -> dict | None:
    dbg("get_installed_office_info() interrogation registre…")
    key_paths = [
        r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
        r"HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun\Configuration",
    ]
    for kp in key_paths:
        try:
            cmd = f'reg query "{kp}" /v VersionToReport'
            dbg(f"  reg query : {cmd}")
            p = subprocess.run(cmd, capture_output=True, text=True, shell=True,
                               creationflags=0x08000000)
            if p.returncode != 0:
                dbg(f"  reg query retcode={p.returncode} pour {kp}", "WARN")
                continue
            m = re.search(r'REG_SZ\s+(.+)', p.stdout)
            if not m:
                dbg(f"  Pas de REG_SZ dans la sortie pour {kp}", "WARN")
                continue
            info = {"version": m.group(1).strip(), "arch": "x64", "lang": ""}
            dbg(f"  Version trouvée : {info['version']}", "OK")

            # Architecture
            cmd_a = f'reg query "{kp}" /v Platform'
            pa = subprocess.run(cmd_a, capture_output=True, text=True, shell=True,
                                creationflags=0x08000000)
            if pa.returncode == 0:
                ma = re.search(r'REG_SZ\s+(.+)', pa.stdout)
                if ma:
                    info["arch"] = ma.group(1).strip()
                    dbg(f"  Architecture : {info['arch']}")

            # Langue — essayer plusieurs clés
            lang_keys = ["ClientCulture", "ClientFolder"]
            for lang_key in lang_keys:
                cmd_l = f'reg query "{kp}" /v {lang_key}'
                pl = subprocess.run(cmd_l, capture_output=True, text=True, shell=True,
                                    creationflags=0x08000000)
                if pl.returncode == 0:
                    ml = re.search(r'REG_SZ\s+(.+)', pl.stdout)
                    if ml:
                        val = ml.group(1).strip()
                        if lang_key == "ClientCulture":
                            info["lang"] = val
                            dbg(f"  Langue (ClientCulture) : {info['lang']}")
                            break
                        elif lang_key == "ClientFolder":
                            # ClientFolder contient parfois la langue
                            if "-" in val and len(val) <= 6:
                                info["lang"] = val
                                dbg(f"  Langue (ClientFolder) : {info['lang']}")
                                break

            # Si langue toujours vide, chercher dans ProductReleaseIds
            if not info["lang"]:
                cmd_pr = f'reg query "{kp}" /v ProductReleaseIds'
                pr = subprocess.run(cmd_pr, capture_output=True, text=True, shell=True,
                                    creationflags=0x08000000)
                if pr.returncode == 0:
                    dbg(f"  ProductReleaseIds stdout : {pr.stdout.strip()}")

                # Essayer aussi le Culture de l'interface
                for culture_key in ["x-none", "InstallationPath"]:
                    pass  # Déjà couvert ci-dessus

                # Dernière tentative : chercher dans les sous-clés de langue
                cmd_lang_sub = f'reg query "{kp.rsplit(chr(92), 1)[0]}\\ClickToRun\\ProductReleaseIDs" /s'
                try:
                    pl2 = subprocess.run(cmd_lang_sub, capture_output=True, text=True,
                                         shell=True, creationflags=0x08000000, timeout=5)
                    if pl2.returncode == 0:
                        # Chercher un pattern de culture comme "fr-fr", "en-us"
                        lang_match = re.search(r'\b([a-z]{2}-[a-z]{2})\b',
                                               pl2.stdout, re.IGNORECASE)
                        if lang_match:
                            info["lang"] = lang_match.group(1)
                            dbg(f"  Langue (sous-clé) : {info['lang']}")
                except Exception:
                    pass

            # Ultime fallback : lire ClientCulture depuis une autre clé
            if not info["lang"]:
                alt_paths = [
                    r"HKCU\SOFTWARE\Microsoft\Office\16.0\Common\LanguageResources",
                ]
                for alt_kp in alt_paths:
                    try:
                        cmd_alt = f'reg query "{alt_kp}" /v UILanguage'
                        pa2 = subprocess.run(cmd_alt, capture_output=True, text=True,
                                             shell=True, creationflags=0x08000000, timeout=5)
                        if pa2.returncode == 0:
                            ma2 = re.search(r'REG_DWORD\s+0x([0-9a-fA-F]+)', pa2.stdout)
                            if ma2:
                                lcid = int(ma2.group(1), 16)
                                lcid_to_lang = {
                                    1036: "fr-FR", 1033: "en-US", 1031: "de-DE",
                                    3082: "es-ES", 1040: "it-IT", 1046: "pt-BR",
                                    2070: "pt-PT", 1049: "ru-RU", 2052: "zh-CN",
                                    1041: "ja-JP", 1042: "ko-KR",
                                }
                                info["lang"] = lcid_to_lang.get(lcid, f"LCID:{lcid}")
                                dbg(f"  Langue (LCID {lcid}) : {info['lang']}")
                    except Exception:
                        pass

            dbg(f"  Résultat final : {info}", "OK")
            return info
        except Exception as e:
            dbg(f"  Exception reg query {kp} : {e}", "ERR")
            continue
    dbg("  Aucune installation Office détectée", "WARN")
    return None

def check_office_activation_status() -> dict:
    dbg("check_office_activation_status() démarrage")
    info = {"installed": False, "activated": False, "version": ""}
    inst = get_installed_office_info()
    if inst:
        info["installed"] = True
        info["version"] = inst["version"]
        dbg(f"  Office installé : version={inst['version']}", "OK")
    else:
        dbg("  Office non installé")
        return info

    # ── Méthode 1 : Détection Ohook (prioritaire) ──
    dbg("  Vérification Ohook…")
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")

    # Ohook installe des DLL spécifiques
    ohook_files = [
        os.path.join(pf, "Microsoft Office", "root", "vfs",
                     "System", "sppcs.dll"),
        os.path.join(pf, "Microsoft Office", "root", "Office16",
                     "sppc64.dll"),
        os.path.join(pf, "Microsoft Office", "root", "Office16",
                     "sppc32.dll"),
        os.path.join(pf86, "Microsoft Office", "root", "vfs",
                     "System", "sppcs.dll"),
        os.path.join(pf86, "Microsoft Office", "root", "Office16",
                     "sppc64.dll"),
        os.path.join(pf86, "Microsoft Office", "root", "Office16",
                     "sppc32.dll"),
    ]

    # Aussi chercher via InstallationPath du registre
    try:
        for reg_key in [
            r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
            r"HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun\Configuration",
        ]:
            cmd = f'reg query "{reg_key}" /v InstallationPath'
            result = subprocess.run(cmd, capture_output=True, text=True,
                                    shell=True, creationflags=0x08000000, timeout=5)
            if result.returncode == 0:
                m = re.search(r'REG_SZ\s+(.+)', result.stdout)
                if m:
                    install_path = m.group(1).strip()
                    ohook_files.extend([
                        os.path.join(install_path, "root", "vfs",
                                     "System", "sppcs.dll"),
                        os.path.join(install_path, "root", "Office16",
                                     "sppc64.dll"),
                        os.path.join(install_path, "root", "Office16",
                                     "sppc32.dll"),
                    ])
                    break
    except Exception:
        pass

    for dll_path in ohook_files:
        if os.path.isfile(dll_path):
            dbg(f"  Ohook DLL trouvée : {dll_path}", "OK")
            info["activated"] = True

            # Vérifier que c'est bien un hook Ohook (pas le fichier système original)
            try:
                size = os.path.getsize(dll_path)
                dbg(f"  Taille DLL : {size} octets")
                # Les DLL Ohook font typiquement < 500 KB
                # Les originales Microsoft sont plus grosses
                if size < 600000:
                    dbg("  Taille compatible Ohook -> ACTIVÉ", "OK")
                else:
                    dbg("  Taille trop grande, peut-être DLL originale", "WARN")
                    info["activated"] = False
                    continue
            except Exception:
                pass

            dbg("  Activation Ohook détectée", "OK")
            return info

    # ── Méthode 2 : Registre Ohook (clés spécifiques) ──
    dbg("  Vérification registre Ohook…")
    ohook_reg_keys = [
        r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
    ]
    for reg_key in ohook_reg_keys:
        try:
            # Chercher les clés que Ohook écrit
            cmd = f'reg query "{reg_key}" /v ProductReleaseIds'
            result = subprocess.run(cmd, capture_output=True, text=True,
                                    shell=True, creationflags=0x08000000, timeout=5)
            if result.returncode == 0:
                m = re.search(r'REG_SZ\s+(.+)', result.stdout)
                if m:
                    dbg(f"  ProductReleaseIds : {m.group(1).strip()}")
        except Exception:
            pass

    # ── Méthode 3 : ospp.vbs (fallback, ne détecte pas Ohook) ──
    dbg("  Vérification ospp.vbs (fallback)…")

    ospp = None
    paths_to_try = [
        os.path.join(pf, "Microsoft Office", "root", "Office16", "ospp.vbs"),
        os.path.join(pf86, "Microsoft Office", "root", "Office16", "ospp.vbs"),
        os.path.join(pf, "Microsoft Office", "Office16", "ospp.vbs"),
        os.path.join(pf86, "Microsoft Office", "Office16", "ospp.vbs"),
    ]

    # Ajouter chemin via registre
    try:
        reg_path = r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        cmd = f'reg query "{reg_path}" /v InstallationPath'
        result = subprocess.run(cmd, capture_output=True, text=True, shell=True,
                                creationflags=0x08000000, timeout=5)
        if result.returncode == 0:
            m = re.search(r'REG_SZ\s+(.+)', result.stdout)
            if m:
                install_path = m.group(1).strip()
                paths_to_try.insert(0, os.path.join(
                    install_path, "root", "Office16", "ospp.vbs"))
    except Exception:
        pass

    for pp in paths_to_try:
        dbg(f"  Recherche ospp.vbs : {pp}")
        if os.path.isfile(pp):
            ospp = pp
            dbg(f"  ospp.vbs trouvé : {pp}", "OK")
            break

    if ospp:
        try:
            cmd = f'cscript //nologo "{ospp}" /dstatus'
            dbg(f"  Exécution : {cmd}")
            p = subprocess.run(cmd, capture_output=True, text=True, timeout=15,
                               shell=True, creationflags=0x08000000)
            out = p.stdout
            dbg(f"  ospp.vbs retcode={p.returncode}, sortie ({len(out)} car.):")
            for line in out.splitlines()[:10]:
                stripped = line.strip()
                if stripped:
                    dbg(f"    | {stripped}")
            if "licensed" in out.lower() and "notification" not in out.lower():
                info["activated"] = True
                dbg("  Statut : LICENSED (activé)", "OK")
            elif "notification" in out.lower() or "grace" in out.lower():
                # Peut être Ohook — vérifier les DLL une dernière fois
                dbg("  Statut : Notification/Grace — vérif Ohook implicite", "WARN")
            else:
                dbg("  Statut : NON licensed", "WARN")
        except subprocess.TimeoutExpired:
            dbg("  ospp.vbs timeout (>15s)", "ERR")
        except Exception as e:
            dbg(f"  Erreur ospp.vbs : {e}", "ERR")
    else:
        dbg("  ospp.vbs NON TROUVÉ", "WARN")

    dbg(f"  Résultat final activation : {info}")
    return info

def _count_urls_in_bat(file_path: str) -> int:
    dbg(f"_count_urls_in_bat('{file_path}')")
    count = 0
    try:
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                if "https://" in line and not line.strip().startswith("::"):
                    count += 1
    except Exception as e:
        dbg(f"  Erreur lecture : {e}", "ERR")
    result = max(count, 1)
    dbg(f"  URLs trouvées : {count} -> retourne {result}")
    return result


ANSI_RE = re.compile(r'\x1b\[[0-9;]*[A-Za-z]')


def _strip_ansi(text: str) -> str:
    return ANSI_RE.sub('', text)


# ──────────────────────────────────────────────
# Interface
# ──────────────────────────────────────────────

class App(ctk.CTk):
    def __init__(self):
        dbg("=" * 60)
        dbg("App.__init__() démarrage application")
        dbg(f"  Python : {sys.version}")
        dbg(f"  Script : {os.path.abspath(__file__)}")
        dbg(f"  CWD    : {os.getcwd()}")
        dbg(f"  PID    : {os.getpid()}")
        dbg("=" * 60)

        super().__init__()
        self.title("Office Source Downloader & Installer")
        self.geometry("1080x920")
        self.minsize(960, 840)
        self.configure(fg_color=COL["bg_app"])

        self.config_data = load_config()

        # ── Attributs d'instance ──
        self.scripts = {}
        self.source_valid = False
        self._downloading = False
        self._installing = False
        self._c2r_path = ""
        self._download_cancel = threading.Event()
        self._download_error_detected = False

        # Préférences
        self._pref_channel = self.config_data.get(
            "Office", "default_channel", fallback="3: Current / Monthly").strip()
        self._pref_build = self.config_data.get(
            "Office", "default_build", fallback="1: Windows 11/10 (Latest)").strip()
        self._pref_bitness = self.config_data.get(
            "Office", "default_bitness", fallback="2: 64-bit (x64)").strip()
        self._pref_lang = self.config_data.get(
            "Office", "default_language", fallback="11: fr-FR (Français)").strip()
        self._pref_suite = self.config_data.get(
            "Office", "default_suite", fallback="O365HomePremRetail").strip()
        self._pref_apps = set(self.config_data.get(
            "Office", "default_apps", fallback="Word,Excel,PowerPoint").split(","))

        dbg(f"  Préférences chargées : channel={self._pref_channel}, "
            f"build={self._pref_build}, bitness={self._pref_bitness}, "
            f"lang={self._pref_lang}, suite={self._pref_suite}")
        dbg(f"  Apps par défaut : {self._pref_apps}")

        self._check_required_scripts()
        self._build_ui()
        dbg("UI construite", "OK")

        self.after(300, self._on_check_activation)
        dbg("  Vérification activation planifiée (after 300ms)")
        self.after(400, self._scan_source)
        dbg("  Scan source planifié (after 400ms)")

    # ──────────────────────────────────────────
    # Vérification scripts requis
    # ──────────────────────────────────────────

    def _check_required_scripts(self):
        dbg("--- Vérification des scripts requis ---", "STEP")
        required = [
            "YAOCTRU_Generator.cmd",
            "Ohook_Activation_AIO.cmd",
            "aria2c.exe",
        ]
        for name in required:
            path = find_script(name)
            self.scripts[name] = path
            if path:
                dbg(f"  ✓ {name} -> {path}", "OK")
            else:
                dbg(f"  ✗ {name} -> MANQUANT", "ERR")

    # ──────────────────────────────────────────
    # Élévation administrateur
    # ──────────────────────────────────────────

    @staticmethod
    def _is_admin() -> bool:
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except Exception:
            return False

    @staticmethod
    def _run_as_admin(cmd: str, work_dir: str, wait: bool = True) -> int:
        dbg(f"_run_as_admin() cmd={cmd}", "STEP")
        dbg(f"  work_dir={work_dir}")

        class SHELLEXECUTEINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", ctypes.c_ulong),
                ("fMask", ctypes.c_ulong),
                ("hwnd", ctypes.c_void_p),
                ("lpVerb", ctypes.c_wchar_p),
                ("lpFile", ctypes.c_wchar_p),
                ("lpParameters", ctypes.c_wchar_p),
                ("lpDirectory", ctypes.c_wchar_p),
                ("nShow", ctypes.c_int),
                ("hInstApp", ctypes.c_void_p),
                ("lpIDList", ctypes.c_void_p),
                ("lpClass", ctypes.c_wchar_p),
                ("hkeyClass", ctypes.c_void_p),
                ("dwHotKey", ctypes.c_ulong),
                ("hIconOrMonitor", ctypes.c_void_p),
                ("hProcess", ctypes.c_void_p),
            ]

        SEE_MASK_NOCLOSEPROCESS = 0x00000040
        SEE_MASK_NOASYNC = 0x00000100
        SW_HIDE = 0

        sei = SHELLEXECUTEINFO()
        sei.cbSize = ctypes.sizeof(SHELLEXECUTEINFO)
        sei.fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_NOASYNC
        sei.hwnd = None
        sei.lpVerb = "runas"
        sei.lpFile = "cmd.exe"
        sei.lpParameters = f'/c "{cmd}"'
        sei.lpDirectory = work_dir
        sei.nShow = SW_HIDE
        sei.hProcess = None

        dbg(f'  ShellExecuteEx : runas cmd.exe /c "{cmd}"')

        success = ctypes.windll.shell32.ShellExecuteExW(ctypes.byref(sei))
        if not success:
            err = ctypes.GetLastError()
            dbg(f"  ShellExecuteEx échoué : erreur {err}", "ERR")
            if err == 1223:
                dbg("  UAC refusé par l'utilisateur", "WARN")
            return -1

        if not wait or not sei.hProcess:
            dbg("  Processus lancé (pas d'attente)", "OK")
            return 0

        handle = sei.hProcess
        dbg(f"  Attente du processus (handle={handle})…")

        INFINITE = 0xFFFFFFFF
        ctypes.windll.kernel32.WaitForSingleObject(
            ctypes.c_void_p(handle), INFINITE)

        exit_code = ctypes.c_ulong(0)
        ctypes.windll.kernel32.GetExitCodeProcess(
            ctypes.c_void_p(handle), ctypes.byref(exit_code))
        ctypes.windll.kernel32.CloseHandle(ctypes.c_void_p(handle))

        dbg(f"  Processus terminé : exit_code={exit_code.value}", "OK")
        return exit_code.value

    # ──────────────────────────────────────────
    # Construction UI
    # ──────────────────────────────────────────

    def _make_sep(self, parent):
        sep = ctk.CTkFrame(parent, height=1, fg_color=COL["border"])
        sep.pack(fill="x", padx=16, pady=(4, 4))
        return sep

    def _build_ui(self):
        dbg("_build_ui() construction de l'interface", "STEP")

        # ═══ TITRE ═══
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
        title_frame.pack(fill="x", padx=30, pady=(18, 4))
        ctk.CTkLabel(title_frame, text="Office 365",
                     font=ctk.CTkFont(family="Segoe UI", size=28, weight="bold"),
                     text_color=COL["text_primary"]).pack(side="left")
        ctk.CTkLabel(title_frame, text="  Téléchargement & Installation",
                     font=ctk.CTkFont(family="Segoe UI", size=28),
                     text_color=COL["text_muted"]).pack(side="left")
        ctk.CTkFrame(self, height=1, fg_color=COL["border"]).pack(
            fill="x", padx=30, pady=(8, 12))

        # ═══ GRILLE UNIQUE : Zone haute + 4 cartes ═══
        grid_frame = ctk.CTkFrame(self, fg_color="transparent")
        grid_frame.pack(fill="both", expand=True, padx=30, pady=(0, 8))
        for i in range(4):
            grid_frame.columnconfigure(i, weight=1, uniform="grid")
        grid_frame.rowconfigure(0, weight=0)  # ligne haute : taille auto
        grid_frame.rowconfigure(1, weight=1)  # ligne cartes : extensible

        # ── LIGNE 0, GAUCHE : État d'activation (colonnes 0-1) ──
        af = ctk.CTkFrame(grid_frame, corner_radius=10, fg_color=COL["bg_card"],
                          border_width=1, border_color=COL["border"])
        af.grid(row=0, column=0, columnspan=2, sticky="nsew",
                padx=(0, 4), pady=(0, 8))

        ah = ctk.CTkFrame(af, fg_color="transparent")
        ah.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(ah, text="ÉTAT OFFICE",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=COL["text_muted"]).pack(side="left")

        ah_btns = ctk.CTkFrame(ah, fg_color="transparent")
        ah_btns.pack(side="right")

        self.btn_uninstall = ctk.CTkButton(
            ah_btns, text="Désinstaller", command=self._on_uninstall,
            fg_color=COL["accent_red"], hover_color=COL["accent_red_h"],
            text_color="#ffffff", corner_radius=6, height=26, width=120,
            font=ctk.CTkFont(size=11), state="disabled")
        self.btn_uninstall.pack(side="left", padx=(0, 4))

        self.btn_check = ctk.CTkButton(
            ah_btns, text="Revérifier", command=self._on_check_activation,
            fg_color=COL["btn_neutral"], hover_color=COL["btn_neutral_h"],
            text_color=COL["text_primary"], corner_radius=6, height=26, width=100,
            font=ctk.CTkFont(size=11))
        self.btn_check.pack(side="left")

        self.act_status = ctk.CTkLabel(
            af, text="Vérification en cours…",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=COL["status_warn"], anchor="w")
        self.act_status.pack(padx=16, pady=(4, 2), anchor="w")

        self.act_details = ctk.CTkLabel(
            af, text="",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COL["text_secondary"], anchor="w", justify="left")
        self.act_details.pack(padx=16, pady=(0, 12), anchor="w")

        # ── LIGNE 0, DROITE : Source & Actions (colonnes 2-3) ──
        sf = ctk.CTkFrame(grid_frame, corner_radius=10, fg_color=COL["bg_card"],
                          border_width=1, border_color=COL["border"])
        sf.grid(row=0, column=2, columnspan=2, sticky="nsew",
                padx=(4, 0), pady=(0, 8))

        sh = ctk.CTkFrame(sf, fg_color="transparent")
        sh.pack(fill="x", padx=16, pady=(12, 4))
        ctk.CTkLabel(sh, text="SOURCE OFFICE",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=COL["text_muted"]).pack(side="left")
        self.source_badge = ctk.CTkLabel(
            sh, text="ABSENT", font=ctk.CTkFont(size=10, weight="bold"),
            text_color=COL["status_err"])
        self.source_badge.pack(side="right")

        self.source_info = ctk.CTkLabel(
            sf, text="Aucune source détectée — lancez un téléchargement",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=COL["text_dim"], anchor="w", justify="left")
        self.source_info.pack(padx=16, pady=(4, 2), anchor="w")

        self.installed_info_label = ctk.CTkLabel(
            sf, text="",
            font=ctk.CTkFont(family="Segoe UI", size=10),
            text_color=COL["text_dim"], anchor="w")
        self.installed_info_label.pack(padx=16, pady=(0, 4), anchor="w")

        btn_row = ctk.CTkFrame(sf, fg_color="transparent")
        btn_row.pack(padx=16, pady=(4, 12))

        self.btn_browse = ctk.CTkButton(
            btn_row, text="Parcourir…", command=self._on_browse,
            fg_color=COL["btn_neutral"], hover_color=COL["btn_neutral_h"],
            text_color=COL["text_primary"], corner_radius=6, height=28, width=100,
            font=ctk.CTkFont(size=11))
        self.btn_browse.pack(side="left", padx=(0, 6))

        self.btn_delete_source = ctk.CTkButton(
            btn_row, text="Supprimer source", command=self._on_delete_source,
            fg_color=COL["accent_orange"], hover_color=COL["accent_orange_h"],
            text_color="#ffffff", corner_radius=6, height=28, width=130,
            font=ctk.CTkFont(size=11), state="disabled")
        self.btn_delete_source.pack(side="left")

        # ── LIGNE 1 : 4 cartes ──
        dbg("  Construction carte 1 : Téléchargement")
        self._build_card_download(grid_frame, 0)
        dbg("  Construction carte 2 : Édition")
        self._build_card_edition(grid_frame, 1)
        dbg("  Construction carte 3 : Applications")
        self._build_card_apps(grid_frame, 2)
        dbg("  Construction carte 4 : Paramètres")
        self._build_card_settings(grid_frame, 3)

        # ═══ ZONE BASSE : Progression + Install ═══
        self._make_sep(self)

        prog_frame = ctk.CTkFrame(self, corner_radius=10, fg_color=COL["bg_card"],
                                  border_width=1, border_color=COL["border"])
        prog_frame.pack(fill="x", padx=30, pady=(0, 4))

        prog_top = ctk.CTkFrame(prog_frame, fg_color="transparent")
        prog_top.pack(fill="x", padx=16, pady=(12, 4))

        self.dl_status_label = ctk.CTkLabel(
            prog_top, text="En attente",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=COL["text_secondary"], anchor="w")
        self.dl_status_label.pack(side="left")

        self.btn_cancel_dl = ctk.CTkButton(
            prog_top, text="Annuler", command=self._on_cancel,
            fg_color=COL["accent_red"], hover_color=COL["accent_red_h"],
            text_color="#ffffff", corner_radius=6, height=26, width=80,
            font=ctk.CTkFont(size=11), state="disabled")
        self.btn_cancel_dl.pack(side="right")

        self.dl_progress = ctk.CTkProgressBar(
            prog_frame, mode="determinate", height=8,
            progress_color=COL["progress_fill"], fg_color=COL["progress_bg"],
            corner_radius=4)
        self.dl_progress.pack(fill="x", padx=16, pady=(2, 4))
        self.dl_progress.set(0)

        self.dl_percent = ctk.CTkLabel(
            prog_frame, text="0 % (0/0)",
            font=ctk.CTkFont(family="Segoe UI", size=13, weight="bold"),
            text_color=COL["text_primary"])
        self.dl_percent.pack(padx=16, pady=(0, 12))

        # Bouton principal
        act_frame = ctk.CTkFrame(self, fg_color="transparent")
        act_frame.pack(pady=(4, 16))

        self.btn_install = ctk.CTkButton(
            act_frame, text="Télécharger et installer Office",
            command=self._on_download_and_install,
            fg_color=COL["accent_green"], hover_color=COL["accent_green_h"],
            text_color="#1a1a1a", corner_radius=6, height=40, width=280,
            font=ctk.CTkFont(size=14, weight="bold"))
        self.btn_install.pack()

    # ── Carte 1 : Téléchargement ──
    def _build_card_download(self, parent, col):
        card = ctk.CTkFrame(parent, corner_radius=10, fg_color=COL["bg_card"],
                            border_width=1, border_color=COL["border"])
        card.grid(row=1, column=col, sticky="nsew",
                  padx=(0 if col == 0 else 4, 4 if col < 3 else 0), pady=0)
                  
        self._card_header(card, "TÉLÉCHARGEMENT", COL["accent_purple"])
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        def _add_sel(label_text, values_dict, pref):
            row = ctk.CTkFrame(content, fg_color=COL["bg_card_alt"],
                               corner_radius=6, border_width=1, border_color=COL["border"])
            row.pack(fill="x", padx=4, pady=3)
            ctk.CTkLabel(row, text=label_text,
                         font=ctk.CTkFont(size=11, weight="bold"),
                         text_color=COL["text_secondary"], anchor="w").pack(
                padx=(10, 4), pady=(6, 2), anchor="w")
            keys = list(values_dict.keys())
            cb = ctk.CTkComboBox(
                row, values=keys, state="readonly",
                font=ctk.CTkFont(size=11), dropdown_font=ctk.CTkFont(size=11),
                corner_radius=4, fg_color=COL["bg_input"],
                border_color=COL["border_light"], button_color=COL["btn_neutral"],
                button_hover_color=COL["btn_neutral_h"],
                dropdown_fg_color=COL["bg_card_alt"],
                dropdown_hover_color=COL["btn_neutral"],
                text_color=COL["text_primary"])
            selected = self._pick_best(keys, pref)
            cb.set(selected)
            dbg(f"    Sélecteur '{label_text}' -> '{selected}'")
            cb.pack(fill="x", padx=6, pady=(0, 6))
            return cb

        self.combo_channel = _add_sel("Canal :", CHANNELS_YAOCTRU, self._pref_channel)
        self.combo_build = _add_sel("Build OS :", BUILDS, self._pref_build)
        self.combo_bitness = _add_sel("Architecture :", BITNESS, self._pref_bitness)
        self.combo_lang = _add_sel("Langue :", LANGUAGES, self._pref_lang)
        self.combo_dl_type = _add_sel("Type DL :", DL_TYPES, "1: Full Office Source")

        # Bouton télécharger uniquement
        self.btn_download_only = ctk.CTkButton(
            content, text="Télécharger uniquement",
            command=self._on_download_only,
            fg_color=COL["accent_purple"], hover_color=COL["accent_purple_h"],
            text_color="#ffffff", corner_radius=6, height=32,
            font=ctk.CTkFont(size=12, weight="bold"))
        self.btn_download_only.pack(fill="x", padx=4, pady=(8, 0))

    def _on_download_only(self):
        """Télécharge la source Office sans installer"""
        dbg("_on_download_only() appelé", "STEP")
        if self._downloading or self._installing:
            dbg("  Déjà en cours, ignoré", "WARN")
            return

        if not self.scripts.get("aria2c.exe"):
            dbg("  aria2c.exe manquant -> blocage", "ERR")
            mb.showerror("Dépendance manquante",
                         "aria2c.exe est introuvable !\n\n"
                         "Placez aria2c.exe dans le même dossier que le programme\n"
                         "et relancez l'application.")
            self._dl_set_status("aria2c.exe manquant — téléchargement impossible",
                                COL["status_err"])
            return

        if not self.scripts.get("YAOCTRU_Generator.cmd"):
            dbg("  YAOCTRU_Generator.cmd manquant -> blocage", "ERR")
            mb.showerror("Script manquant", "YAOCTRU_Generator.cmd est introuvable !")
            return

        self._download_cancel.clear()
        self._downloading = True
        self._download_error_detected = False
        self.btn_install.configure(state="disabled")
        self.btn_download_only.configure(state="disabled")
        self.btn_uninstall.configure(state="disabled")
        self.btn_cancel_dl.configure(state="normal")
        self.dl_progress.set(0)
        self.dl_percent.configure(text="0 % (0/0)")

        dbg("  Paramètres de téléchargement :")
        dbg(f"    Canal     : {self.combo_channel.get()}")
        dbg(f"    Build     : {self.combo_build.get()}")
        dbg(f"    Arch      : {self.combo_bitness.get()}")
        dbg(f"    Langue    : {self.combo_lang.get()}")
        dbg(f"    Type DL   : {self.combo_dl_type.get()}")

        threading.Thread(target=self._t_download_only, daemon=True,
                         name="DownloadOnly").start()

    def _t_download_only(self):
        """Thread : téléchargement seul sans installation"""
        dbg("[Thread] _t_download_only() démarré", "STEP")
        start_time = time.time()

        try:
            # ── Phase 1 : Génération ──
            dbg("[Thread] === PHASE 1 : Génération ===", "STEP")
            self.after(0, lambda: self._dl_set_status(
                "Génération du script de téléchargement…", COL["status_warn"]))

            gen_ok = self._run_generator()
            dbg(f"[Thread] Génération résultat : {gen_ok}")

            if self._download_cancel.is_set():
                self.after(0, lambda: self._dl_set_status(
                    "Annulé par l'utilisateur", COL["status_warn"]))
                return
            if not gen_ok:
                self.after(0, lambda: self._dl_set_status(
                    "Échec de la génération du script", COL["status_err"]))
                return

            # ── Phase 2 : Téléchargement ──
            dbg("[Thread] === PHASE 2 : Téléchargement ===", "STEP")
            dl_ok = self._run_download()
            dbg(f"[Thread] Téléchargement résultat : {dl_ok}")

            if self._download_cancel.is_set():
                self.after(0, lambda: self._dl_set_status(
                    "Annulé par l'utilisateur", COL["status_warn"]))
                return

            elapsed = time.time() - start_time

            if self._download_error_detected:
                self.after(0, lambda: self._dl_set_status(
                    "Téléchargement échoué — erreur détectée", COL["status_err"]))
                return
            if not dl_ok:
                self.after(0, lambda: self._dl_set_status(
                    "Téléchargement échoué — vérifiez les logs", COL["status_err"]))
                return

            dbg(f"[Thread] Téléchargement seul terminé en {elapsed:.1f}s", "OK")
            self.after(0, lambda e=elapsed: self._dl_set_status(
                f"✓ Téléchargement terminé ({e:.0f}s) — prêt à installer",
                COL["status_ok"]))

        except Exception as e:
            dbg(f"[Thread] EXCEPTION : {e}", "ERR")
            dbg(traceback.format_exc(), "ERR")
            self.after(0, lambda e=e: self._dl_set_status(
                f"Erreur inattendue : {e}", COL["status_err"]))
        finally:
            dbg("[Thread] _t_download_only() terminé")
            self.after(0, self._on_worker_done)

    # ── Carte 2 : Édition ──

    def _build_card_edition(self, parent, col):
        card = ctk.CTkFrame(parent, corner_radius=10, fg_color=COL["bg_card"],
                            border_width=1, border_color=COL["border"])
        card.grid(row=1, column=col, sticky="nsew", padx=4, pady=0)
        self._card_header(card, "ÉDITION", COL["accent_blue"])
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        cats = list(SUITES.keys())
        self.combo_suite_cat = ctk.CTkComboBox(
            content, values=cats, state="readonly",
            font=ctk.CTkFont(size=12), dropdown_font=ctk.CTkFont(size=11),
            corner_radius=4, fg_color=COL["bg_input"],
            border_color=COL["border_light"], button_color=COL["btn_neutral"],
            button_hover_color=COL["btn_neutral_h"],
            dropdown_fg_color=COL["bg_card_alt"],
            dropdown_hover_color=COL["btn_neutral"],
            text_color=COL["text_primary"],
            command=self._on_suite_cat_changed)
        self.combo_suite_cat.set(cats[0])
        self.combo_suite_cat.pack(fill="x", padx=4, pady=(4, 8))

        self.suite_radio_frame = ctk.CTkFrame(content, fg_color="transparent")
        self.suite_radio_frame.pack(fill="both", expand=True)

        self.selected_suite_id = ctk.StringVar(value="O365HomePremRetail")
        self._populate_suite_radios(cats[0])

    def _populate_suite_radios(self, cat):
        dbg(f"_populate_suite_radios('{cat}')")
        for w in self.suite_radio_frame.winfo_children():
            w.destroy()
        first = True
        for name, sku in SUITES[cat].items():
            rb = ctk.CTkRadioButton(
                self.suite_radio_frame, text=name,
                variable=self.selected_suite_id, value=sku,
                font=ctk.CTkFont(size=11), text_color=COL["text_primary"],
                fg_color=COL["accent_blue"], hover_color=COL["accent_blue_h"])
            rb.pack(anchor="w", padx=8, pady=3)
            if first:
                self.selected_suite_id.set(sku)
                dbg(f"  Suite par défaut : {name} -> {sku}")
                first = False

    def _on_suite_cat_changed(self, _=None):
        cat = self.combo_suite_cat.get()
        dbg(f"_on_suite_cat_changed() -> '{cat}'")
        self._populate_suite_radios(cat)

    # ── Carte 3 : Applications ──

    def _build_card_apps(self, parent, col):
        card = ctk.CTkFrame(parent, corner_radius=10, fg_color=COL["bg_card"],
                            border_width=1, border_color=COL["border"])
        card.grid(row=1, column=col, sticky="nsew", padx=4, pady=0)
        self._card_header(card, "APPLICATIONS", COL["accent_teal"])
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=8, pady=(0, 8))

        default_on = {"Word", "Excel", "PowerPoint"}
        self.apps_vars = {}
        for name in APPS_MAP:
            var = ctk.BooleanVar(value=(name in default_on))
            self.apps_vars[name] = var
            sw = ctk.CTkSwitch(
                content, text=name, variable=var,
                font=ctk.CTkFont(size=12), text_color=COL["text_primary"],
                fg_color=COL["bg_input"], progress_color=COL["accent_teal"],
                button_color=COL["text_muted"],
                button_hover_color=COL["text_secondary"])
            sw.pack(anchor="w", padx=10, pady=4)
        dbg(f"  Apps activées par défaut : {default_on}")

    # ── Carte 4 : Paramètres ──

    def _build_card_settings(self, parent, col):
        card = ctk.CTkFrame(parent, corner_radius=10, fg_color=COL["bg_card"],
                            border_width=1, border_color=COL["border"])
        card.grid(row=1, column=col, sticky="nsew", padx=(4, 0), pady=0)
        self._card_header(card, "PARAMÈTRES", COL["accent_orange"])
        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        row_ch = ctk.CTkFrame(content, fg_color=COL["bg_card_alt"],
                              corner_radius=6, border_width=1, border_color=COL["border"])
        row_ch.pack(fill="x", padx=4, pady=(4, 8))
        ctk.CTkLabel(row_ch, text="Canal de mise à jour :",
                     font=ctk.CTkFont(size=11, weight="bold"),
                     text_color=COL["text_secondary"], anchor="w").pack(
            padx=(10, 4), pady=(6, 2), anchor="w")
        self.combo_update_channel = ctk.CTkComboBox(
            row_ch, values=CHANNELS_YAOCTRIR, state="readonly",
            font=ctk.CTkFont(size=11), dropdown_font=ctk.CTkFont(size=11),
            corner_radius=4, fg_color=COL["bg_input"],
            border_color=COL["border_light"], button_color=COL["btn_neutral"],
            button_hover_color=COL["btn_neutral_h"],
            dropdown_fg_color=COL["bg_card_alt"],
            dropdown_hover_color=COL["btn_neutral"],
            text_color=COL["text_primary"])
        self.combo_update_channel.set("Monthly")
        self.combo_update_channel.pack(fill="x", padx=6, pady=(0, 6))

        self.opt_updates = ctk.BooleanVar(value=True)
        self.opt_eula = ctk.BooleanVar(value=True)
        self.opt_shutdown = ctk.BooleanVar(value=True)
        self.opt_activate = ctk.BooleanVar(value=True)
        self.opt_telemetry = ctk.BooleanVar(value=True)
        self.opt_no_bing = ctk.BooleanVar(value=True)
        self.opt_display = ctk.BooleanVar(value=True)
        self.opt_pin_taskbar = ctk.BooleanVar(value=True)

        checks = [
            ("Mises à jour auto", self.opt_updates),
            ("Accepter EULA", self.opt_eula),
            ("Fermer applis ouvertes", self.opt_shutdown),
            ("Activation automatique (Ohook)", self.opt_activate),
            ("Désactiver télémétrie", self.opt_telemetry),
            ("Désactiver Bing", self.opt_no_bing),
            ("Affichage complet", self.opt_display),
            ("Icônes Barre des tâches", self.opt_pin_taskbar),
        ]
        for txt, var in checks:
            cb = ctk.CTkCheckBox(
                content, text=txt, variable=var,
                font=ctk.CTkFont(size=11), text_color=COL["text_primary"],
                fg_color=COL["bg_input"], hover_color=COL["btn_neutral_h"],
                checkmark_color=COL["accent_orange"],
                border_color=COL["border_light"])
            cb.pack(anchor="w", padx=8, pady=3)

    def _card_header(self, card, title, accent):
        hdr = ctk.CTkFrame(card, fg_color="transparent")
        hdr.pack(fill="x", padx=16, pady=(14, 8))
        ctk.CTkFrame(hdr, width=4, height=20, fg_color=accent,
                     corner_radius=2).pack(side="left", padx=(0, 8))
        ctk.CTkLabel(hdr, text=title,
                     font=ctk.CTkFont(size=12, weight="bold"),
                     text_color=COL["text_primary"]).pack(side="left")

    def _pick_best(self, values: list, preference: str) -> str:
        if not values:
            return ""
        pref = preference.lower().strip()
        if pref:
            for v in values:
                if v.lower().strip() == pref:
                    return v
            for v in values:
                if v.lower().strip().startswith(pref):
                    return v
        return values[0]

    # ──────────────────────────────────────────
    # Détection source existante
    # ──────────────────────────────────────────

    def _scan_source(self):
        dbg("_scan_source() recherche de source Office…", "STEP")
        sd = os.path.dirname(os.path.abspath(__file__))
        candidates = [
            os.path.join(sd, "C2R_Monthly"),
            os.path.join(sd, "Downloads", "C2R_Monthly"),
            os.path.join(os.getcwd(), "C2R_Monthly"),
            os.path.join(os.getcwd(), "Downloads", "C2R_Monthly"),
        ]

        source = None
        for c in candidates:
            data_dir = os.path.join(c, "Office", "Data")
            exists = os.path.isdir(data_dir)
            dbg(f"  Candidat : {c} -> Office/Data "
                f"{'EXISTE' if exists else 'absent'}")
            if exists and source is None:
                source = c

        inst = get_installed_office_info()
        if inst:
            txt = (f"Installé : v{inst['version']} | "
                   f"{inst['arch']} | {inst['lang']}")
            self.installed_info_label.configure(text=txt)
            dbg(f"  Office installé détecté : {txt}", "OK")
        else:
            self.installed_info_label.configure(text="")
            dbg("  Aucun Office installé détecté")

        if source:
            self._c2r_path = source
            self.source_valid = True
            ver, arch, lang = "??", "??", "??"
            try:
                data_p = os.path.join(source, "Office", "Data")
                cabs = glob.glob(os.path.join(data_p, "v*_*.cab"))
                dbg(f"  Fichiers .cab trouvés : {len(cabs)}")
                if cabs:
                    first_cab = os.path.basename(cabs[0])
                    dbg(f"  Premier cab : {first_cab}")
                    parts = first_cab.split("_")
                    if len(parts) >= 2:
                        arch = "x64" if "64" in parts[0] else "x86"
                        ver = parts[1].replace(".cab", "")
                    dbg(f"  Parsé : ver={ver}, arch={arch}")
                    vf = os.path.join(data_p, ver)
                    if os.path.isdir(vf):
                        for s in os.listdir(vf):
                            if s.startswith("stream.") and s.endswith(".dat"):
                                pp = s.split(".")
                                if len(pp) >= 3 and pp[2] not in ("x-none", "dat"):
                                    lang = pp[2]
                                    dbg(f"  Langue détectée : {lang}")
                                    break
            except Exception as e:
                dbg(f"  Erreur analyse source : {e}", "ERR")

            self.source_info.configure(
                text=f"Source : v{ver} | {arch} | {lang}",
                text_color=COL["status_ok"])
            self.source_badge.configure(text="PRÉSENT", text_color=COL["status_ok"])
            self.btn_delete_source.configure(state="normal")
            self.btn_install.configure(
                text="INSTALLER OFFICE", command=self._on_install_only, state="normal")
            dbg(f"  Source VALIDE : {source} (v{ver} {arch} {lang})", "OK")
            return source
        else:
            self._c2r_path = ""
            self.source_valid = False
            self.source_info.configure(
                text="Aucune source détectée — lancez un téléchargement",
                text_color=COL["text_dim"])
            self.source_badge.configure(text="ABSENT", text_color=COL["status_err"])
            self.btn_delete_source.configure(state="disabled")
            self.btn_install.configure(
                text="Télécharger et installer Office",
                command=self._on_download_and_install, state="normal")
            dbg("  Aucune source trouvée", "WARN")
            return None

    def _on_browse(self):
        dbg("_on_browse() ouverture boîte de dialogue")
        p = filedialog.askdirectory(
            title="Sélectionner le dossier source Office (C2R_Monthly)")
        if p:
            dbg(f"  Dossier sélectionné : {p}")
            self._c2r_path = p
            self._scan_source()
        else:
            dbg("  Annulé par l'utilisateur")

    def _on_delete_source(self):
        dbg("_on_delete_source() demande de suppression")
        if not self._c2r_path or not os.path.isdir(self._c2r_path):
            dbg("  Pas de source à supprimer", "WARN")
            return
        if not mb.askyesno("Confirmation",
                           f"Supprimer la source ?\n\n{self._c2r_path}\n\nAction irréversible."):
            dbg("  Suppression annulée par l'utilisateur")
            return
        path_to_delete = self._c2r_path
        dbg(f"  Suppression de : {path_to_delete}", "STEP")

        def _do_delete():
            try:
                shutil.rmtree(path_to_delete)
                dbg(f"  Source supprimée : {path_to_delete}", "OK")
                self.after(0, self._scan_source)
            except Exception as e:
                dbg(f"  Erreur suppression : {e}", "ERR")
                self.after(0, lambda: mb.showerror(
                    "Erreur", f"Impossible de supprimer :\n{e}"))
        threading.Thread(target=_do_delete, daemon=True).start()

    # ──────────────────────────────────────────
    # Vérification activation
    # ──────────────────────────────────────────

    def _on_check_activation(self):
        dbg("_on_check_activation() lancement vérification", "STEP")
        self.btn_check.configure(state="disabled")
        self.act_status.configure(text="Vérification en cours…",
                                  text_color=COL["status_warn"])
        self.act_details.configure(text="")
        threading.Thread(target=self._t_check_act, daemon=True).start()

    def _t_check_act(self):
        dbg("[Thread] _t_check_act() démarré")
        info = check_office_activation_status()
        dbg(f"[Thread] _t_check_act() résultat : {info}")
        self.after(0, self._u_act, info)

    def _u_act(self, info: dict):
        dbg(f"_u_act() mise à jour UI : {info}")
        self.btn_check.configure(state="normal")
        if info["installed"] and info["activated"]:
            self.act_status.configure(text="Office est activé",
                                      text_color=COL["status_ok"])
            self.act_details.configure(text=f"Version : {info['version']}")
            self.btn_uninstall.configure(state="normal")
            dbg("  UI -> Office activé", "OK")
        elif info["installed"]:
            self.act_status.configure(text="Office installé mais NON activé",
                                      text_color=COL["status_err"])
            self.act_details.configure(text=f"Version : {info['version']}")
            self.btn_uninstall.configure(state="normal")
            dbg("  UI -> Office installé non activé", "WARN")
        else:
            self.act_status.configure(text="Office n'est pas installé",
                                      text_color=COL["text_dim"])
            self.act_details.configure(text="")
            self.btn_uninstall.configure(state="disabled")
            dbg("  UI -> Office non installé")

    # ──────────────────────────────────────────
    # Actions utilisateur
    # ──────────────────────────────────────────

    def _on_download_and_install(self):
        dbg("_on_download_and_install() appelé", "STEP")
        if self._downloading or self._installing:
            dbg("  Déjà en cours, ignoré", "WARN")
            return

        if not self.scripts.get("aria2c.exe"):
            dbg("  aria2c.exe manquant -> blocage", "ERR")
            mb.showerror("Dépendance manquante",
                         "aria2c.exe est introuvable !\n\n"
                         "Placez aria2c.exe dans le même dossier que le programme\n"
                         "et relancez l'application.")
            self._dl_set_status("aria2c.exe manquant — téléchargement impossible",
                                COL["status_err"])
            return

        if not self.scripts.get("YAOCTRU_Generator.cmd"):
            dbg("  YAOCTRU_Generator.cmd manquant -> blocage", "ERR")
            mb.showerror("Script manquant", "YAOCTRU_Generator.cmd est introuvable !")
            return

        self._download_cancel.clear()
        self._downloading = True
        self._download_error_detected = False
        self.btn_install.configure(state="disabled")
        self.btn_uninstall.configure(state="disabled")
        self.btn_download_only.configure(state="disabled")   # ← AJOUTER
        self.btn_cancel_dl.configure(state="normal")
        self.dl_progress.set(0)
        self.dl_percent.configure(text="0 % (0/0)")

        dbg("  Paramètres de téléchargement :")
        dbg(f"    Canal     : {self.combo_channel.get()}")
        dbg(f"    Build     : {self.combo_build.get()}")
        dbg(f"    Arch      : {self.combo_bitness.get()}")
        dbg(f"    Langue    : {self.combo_lang.get()}")
        dbg(f"    Type DL   : {self.combo_dl_type.get()}")
        dbg(f"    Suite     : {self.selected_suite_id.get()}")
        dbg(f"    Canal MAJ : {self.combo_update_channel.get()}")
        apps_on = [n for n, v in self.apps_vars.items() if v.get()]
        apps_off = [n for n, v in self.apps_vars.items() if not v.get()]
        dbg(f"    Apps ON   : {apps_on}")
        dbg(f"    Apps OFF  : {apps_off}")
        dbg(f"    Activate  : {self.opt_activate.get()}")

        threading.Thread(target=self._t_download_and_install, daemon=True).start()

    def _on_install_only(self):
        dbg("_on_install_only() appelé", "STEP")
        if self._installing:
            dbg("  Déjà en cours, ignoré", "WARN")
            return
        self._installing = True
        self.btn_install.configure(state="disabled")
        self.btn_uninstall.configure(state="disabled")
        threading.Thread(target=self._t_install_only, daemon=True).start()

    def _on_cancel(self):
        dbg("_on_cancel() annulation demandée", "WARN")
        self._download_cancel.set()
        self.btn_cancel_dl.configure(state="disabled")
        self._dl_set_status("Annulation…", COL["status_warn"])

    # ──────────────────────────────────────────
    # Désinstallation Office
    # ──────────────────────────────────────────

    def _on_uninstall(self):
        dbg("_on_uninstall() appelé", "STEP")
        if self._downloading or self._installing:
            dbg("  Opération en cours, ignoré", "WARN")
            return

        info = get_installed_office_info()
        if not info:
            dbg("  Office non détecté", "WARN")
            mb.showinfo("Désinstallation", "Aucune installation Office détectée.")
            return

        version = info.get("version", "inconnue")
        arch = info.get("arch", "?")
        lang = info.get("lang", "?")

        confirm = mb.askyesno(
            "Désinstallation complète d'Office",
            f"Voulez-vous désinstaller complètement Office ?\n\n"
            f"  Version : {version}\n"
            f"  Architecture : {arch}\n"
            f"  Langue : {lang}\n\n"
            f"Cette opération va :\n"
            f"  • Fermer toutes les applications Office\n"
            f"  • Supprimer tous les composants Office\n"
            f"  • Nettoyer le registre\n"
            f"  • Supprimer les fichiers résiduels\n\n"
            f"Les documents personnels ne seront PAS supprimés.\n\n"
            f"Continuer ?",
            icon="warning")

        if not confirm:
            dbg("  Désinstallation annulée par l'utilisateur")
            return

        self._installing = True
        self.btn_install.configure(state="disabled")
        self.btn_uninstall.configure(state="disabled")
        self.btn_check.configure(state="disabled")
        self.btn_cancel_dl.configure(state="normal")
        self._download_cancel.clear()

        self.after(0, lambda: self._dl_set_status(
            "Désinstallation d'Office en cours…", COL["status_warn"]))
        self.after(0, lambda: self.dl_progress.set(0))

        threading.Thread(target=self._t_uninstall, daemon=True, name="Uninstall").start()

    def _t_uninstall(self):
        dbg("[Thread] _t_uninstall() démarré", "STEP")
        start_time = time.time()

        try:
            # ── Étape 1/4 : Fermeture des applications ──
            dbg("  Étape 1/4 : Fermeture des applications Office", "STEP")
            self.after(0, lambda: self._dl_set_status(
                "Étape 1/4 — Fermeture des applications Office…",
                COL["status_warn"]))
            self.after(0, lambda: self.dl_progress.set(0.05))

            office_processes = [
                "WINWORD.EXE", "EXCEL.EXE", "POWERPNT.EXE",
                "OUTLOOK.EXE", "ONENOTE.EXE", "MSACCESS.EXE",
                "MSPUB.EXE", "lync.exe", "Teams.exe",
                "OfficeClickToRun.exe", "OfficeC2RClient.exe",
                "AppVShNotify.exe",
            ]
            killed = 0
            for proc_name in office_processes:
                try:
                    result = subprocess.run(
                        ["taskkill", "/F", "/IM", proc_name],
                        capture_output=True, text=True,
                        creationflags=0x08000000, timeout=5)
                    if result.returncode == 0:
                        killed += 1
                        dbg(f"    Fermé : {proc_name}", "OK")
                except Exception:
                    pass
            if killed > 0:
                dbg(f"  {killed} processus fermé(s)", "OK")
                time.sleep(2)

            if self._download_cancel.is_set():
                self.after(0, lambda: self._dl_set_status(
                    "Désinstallation annulée", COL["status_warn"]))
                return

            # ── Étape 2/4 : Désinstallation via ODT/C2R ──
            dbg("  Étape 2/4 : Désinstallation via ODT/C2R", "STEP")
            self.after(0, lambda: self._dl_set_status(
                "Étape 2/4 — Désinstallation des composants Office…",
                COL["status_warn"]))
            self.after(0, lambda: self.dl_progress.set(0.15))

            uninstalled = self._uninstall_via_c2r_setup()

            if self._download_cancel.is_set():
                self.after(0, lambda: self._dl_set_status(
                    "Désinstallation annulée", COL["status_warn"]))
                return

            # ── Étape 3/4 : Fallback via OfficeClickToRun ──
            if not uninstalled:
                dbg("  Étape 3/4 : Fallback via OfficeClickToRun", "STEP")
                self.after(0, lambda: self._dl_set_status(
                    "Étape 3/4 — Méthode alternative…",
                    COL["status_warn"]))
                self.after(0, lambda: self.dl_progress.set(0.35))
                self._uninstall_via_click_to_run()
            else:
                self.after(0, lambda: self.dl_progress.set(0.5))

            if self._download_cancel.is_set():
                self.after(0, lambda: self._dl_set_status(
                    "Désinstallation annulée", COL["status_warn"]))
                return

            # ── Étape 4/4 : Nettoyage complet ──
            dbg("  Étape 4/4 : Nettoyage complet", "STEP")
            self.after(0, lambda: self._dl_set_status(
                "Étape 4/4 — Nettoyage complet (registre, fichiers, services)…",
                COL["status_warn"]))
            self.after(0, lambda: self.dl_progress.set(0.6))
            self._full_cleanup()

            # ── Vérification finale ──
            self.after(0, lambda: self.dl_progress.set(0.95))
            dbg("  Vérification post-désinstallation…", "STEP")
            time.sleep(3)

            post_info = get_installed_office_info()
            really_there = self._is_office_really_installed()
            elapsed = time.time() - start_time

            if post_info and really_there:
                dbg(f"  Office encore installé : v{post_info['version']}",
                    "WARN")
                self.after(0, lambda: self.dl_progress.set(1))
                self.after(0, lambda e=elapsed: self._dl_set_status(
                    f"Désinstallation partielle ({e:.0f}s) — "
                    f"redémarrage recommandé", COL["status_warn"]))
            elif post_info and not really_there:
                # Résidus registre restants, refaire un nettoyage
                dbg("  Résidus registre restants, 2ème nettoyage…", "WARN")
                self._full_cleanup()
                time.sleep(2)
                post_info2 = get_installed_office_info()
                if post_info2:
                    dbg("  Résidus persistants après 2ème nettoyage", "WARN")
                    self.after(0, lambda: self.dl_progress.set(1))
                    self.after(0, lambda e=elapsed: self._dl_set_status(
                        f"Désinstallation terminée ({e:.0f}s) — "
                        f"résidus registre persistants, redémarrez",
                        COL["status_warn"]))
                else:
                    dbg("  2ème nettoyage a tout supprimé", "OK")
                    self.after(0, lambda: self.dl_progress.set(1))
                    self.after(0, lambda e=elapsed: self._dl_set_status(
                        f"✓ Office désinstallé avec succès ({e:.0f}s)",
                        COL["status_ok"]))
            else:
                dbg(f"  Office complètement supprimé en {elapsed:.1f}s", "OK")
                self.after(0, lambda: self.dl_progress.set(1))
                self.after(0, lambda e=elapsed: self._dl_set_status(
                    f"✓ Office désinstallé avec succès ({e:.0f}s)",
                    COL["status_ok"]))

        except Exception as e:
            dbg(f"[Thread] EXCEPTION désinstallation : {e}", "ERR")
            dbg(traceback.format_exc(), "ERR")
            self.after(0, lambda e=e: self._dl_set_status(
                f"Erreur désinstallation : {e}", COL["status_err"]))
        finally:
            dbg("[Thread] _t_uninstall() terminé")
            self.after(0, self._on_worker_done)
            self.after(500, self._on_check_activation)

    def _get_installed_product_ids(self) -> list:
        dbg("_get_installed_product_ids() recherche…")
        product_ids = []
        reg_path = r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
        try:
            cmd = f'reg query "{reg_path}" /v ProductReleaseIds'
            result = subprocess.run(cmd, capture_output=True, text=True, shell=True,
                                    creationflags=0x08000000, timeout=5)
            if result.returncode == 0:
                m = re.search(r'REG_SZ\s+(.+)', result.stdout)
                if m:
                    ids_str = m.group(1).strip()
                    product_ids = [pid.strip() for pid in ids_str.split(",") if pid.strip()]
                    dbg(f"  ProductReleaseIds : {product_ids}", "OK")
        except Exception as e:
            dbg(f"  Erreur lecture ProductReleaseIds : {e}", "ERR")
        return product_ids

    def _uninstall_via_c2r_setup(self) -> bool:
        """Désinstalle via setup.exe /configure avec XML Remove All"""
        dbg("_uninstall_via_c2r_setup() démarré", "STEP")

        script_dir = os.path.dirname(os.path.abspath(__file__))

        # Chercher setup.exe : à côté du script D'ABORD, puis dans ClickToRun
        setup_paths = [
            os.path.join(script_dir, "setup.exe"),
            os.path.join(script_dir, "Downloads", "setup.exe"),
            os.path.join(os.environ.get("CommonProgramFiles", ""),
                         "Microsoft Shared", "ClickToRun", "setup.exe"),
            os.path.join(os.environ.get("ProgramFiles", ""),
                         "Common Files", "Microsoft Shared", "ClickToRun", "setup.exe"),
            os.path.join(os.environ.get("ProgramFiles(x86)", ""),
                         "Common Files", "Microsoft Shared", "ClickToRun", "setup.exe"),
        ]

        setup_exe = None
        for p in setup_paths:
            if p and os.path.isfile(p):
                setup_exe = p
                dbg(f"  setup.exe trouvé : {p}", "OK")
                break
            else:
                dbg(f"  setup.exe absent : {p}")

        if not setup_exe:
            dbg("  setup.exe NON TROUVÉ", "WARN")
            return False

        # Créer XML de désinstallation
        xml_path = os.path.join(script_dir, "_uninstall_config.xml")
        xml_content = '''<Configuration>
  <Remove All="TRUE" />
  <Display Level="None" AcceptEULA="TRUE" />
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE" />
</Configuration>'''

        try:
            with open(xml_path, "w", encoding="utf-8") as f:
                f.write(xml_content)
            dbg(f"  XML désinstallation créé : {xml_path}", "OK")
        except Exception as e:
            dbg(f"  Erreur écriture XML : {e}", "ERR")
            return False

        # Créer wrapper
        wrapper_path = os.path.join(script_dir, "_auto_uninstall.bat")
        uninstall_log = os.path.join(script_dir, "_uninstall_output.log")
        try:
            with open(wrapper_path, "w", encoding="cp850") as f:
                f.write("@echo off\r\n")
                f.write(f'cd /d "{os.path.dirname(setup_exe)}"\r\n')
                f.write(f'echo [%date% %time%] Lancement desinstallation > "{uninstall_log}"\r\n')
                f.write(f'echo Setup: {setup_exe} >> "{uninstall_log}"\r\n')
                f.write(f'echo XML: {xml_path} >> "{uninstall_log}"\r\n')
                f.write(f'echo. >> "{uninstall_log}"\r\n')
                f.write(f'"{setup_exe}" /configure "{xml_path}" >> "{uninstall_log}" 2>&1\r\n')
                f.write("set EC=%errorlevel%\r\n")
                f.write(f'echo. >> "{uninstall_log}"\r\n')
                f.write(f'echo Exit code: %EC% >> "{uninstall_log}"\r\n')
                f.write("exit /b %EC%\r\n")
            dbg(f"  Wrapper créé : {wrapper_path}", "OK")
        except Exception as e:
            dbg(f"  Erreur création wrapper : {e}", "ERR")
            self._cleanup_file(xml_path)
            return False

        try:
            exit_code = self._run_as_admin(wrapper_path, script_dir, wait=True)
            dbg(f"  Désinstallation ODT terminée : exit_code={exit_code}")

            # Lire la sortie
            if os.path.isfile(uninstall_log):
                try:
                    with open(uninstall_log, "r", encoding="cp850",
                              errors="replace") as f:
                        content = f.read()
                    dbg("  ╔══ Sortie désinstallation ══", "STEP")
                    for line in content.splitlines():
                        s = line.rstrip()
                        if s:
                            dbg(f"  ║ {s}")
                    dbg("  ╚══ Fin sortie ══", "STEP")
                except Exception:
                    pass
                self._cleanup_file(uninstall_log)

            return exit_code == 0
        except Exception as e:
            dbg(f"  Erreur exécution désinstallation : {e}", "ERR")
            return False
        finally:
            self._cleanup_file(wrapper_path)
            self._cleanup_file(xml_path)

    def _uninstall_via_click_to_run(self) -> bool:
        """Fallback : désinstalle via OfficeClickToRun.exe productstoremove"""
        dbg("_uninstall_via_click_to_run() démarré", "STEP")

        c2r_paths = [
            os.path.join(os.environ.get("CommonProgramFiles", ""),
                         "Microsoft Shared", "ClickToRun", "OfficeClickToRun.exe"),
            os.path.join(os.environ.get("ProgramFiles", ""),
                         "Common Files", "Microsoft Shared", "ClickToRun",
                         "OfficeClickToRun.exe"),
        ]

        c2r_exe = None
        for p in c2r_paths:
            if p and os.path.isfile(p):
                c2r_exe = p
                dbg(f"  OfficeClickToRun.exe trouvé : {p}", "OK")
                break

        if not c2r_exe:
            dbg("  OfficeClickToRun.exe non trouvé", "WARN")
            return False

        product_ids = self._get_installed_product_ids()
        if not product_ids:
            product_ids = ["O365HomePremRetail", "O365ProPlusRetail"]
            dbg(f"  Fallback Product IDs : {product_ids}", "WARN")

        script_dir = os.path.dirname(os.path.abspath(__file__))
        wrapper_path = os.path.join(script_dir, "_auto_uninstall_c2r.bat")

        try:
            with open(wrapper_path, "w", encoding="cp850") as f:
                f.write("@echo off\r\n")
                for pid in product_ids:
                    f.write(
                        f'"{c2r_exe}" scenario=install '
                        f'scenariosubtype=ARP sourcetype=None '
                        f'productstoremove={pid}.16_en-us '
                        f'DisplayLevel=False\r\n')
                    f.write("timeout /t 5 /nobreak >nul\r\n")
                f.write("exit /b 0\r\n")
            dbg(f"  Wrapper C2R créé : {wrapper_path}", "OK")
        except Exception as e:
            dbg(f"  Erreur création wrapper : {e}", "ERR")
            return False

        try:
            exit_code = self._run_as_admin(wrapper_path, script_dir, wait=True)
            dbg(f"  Désinstallation C2R terminée : exit_code={exit_code}")
            return exit_code == 0
        except Exception as e:
            dbg(f"  Erreur : {e}", "ERR")
            return False
        finally:
            self._cleanup_file(wrapper_path)

    # ──────────────────────────────────────────
    # Thread : Download + Install
    # ──────────────────────────────────────────

    def _t_download_and_install(self):
        dbg("[Thread] _t_download_and_install() démarré", "STEP")
        start_time = time.time()

        try:
            if not self.scripts.get("aria2c.exe"):
                dbg("  [Thread] aria2c.exe manquant", "ERR")
                self.after(0, lambda: self._dl_set_status(
                    "aria2c.exe manquant — impossible de télécharger", COL["status_err"]))
                return

            # ── Phase 1 : Génération ──
            dbg("[Thread] === PHASE 1 : Génération ===", "STEP")
            self.after(0, lambda: self._dl_set_status(
                "Génération du script de téléchargement…", COL["status_warn"]))

            gen_ok = self._run_generator()
            dbg(f"[Thread] Génération résultat : {gen_ok}")

            if self._download_cancel.is_set():
                self.after(0, lambda: self._dl_set_status(
                    "Annulé par l'utilisateur", COL["status_warn"]))
                return
            if not gen_ok:
                self.after(0, lambda: self._dl_set_status(
                    "Échec de la génération du script", COL["status_err"]))
                return

            # ── Phase 2 : Téléchargement ──
            dbg("[Thread] === PHASE 2 : Téléchargement ===", "STEP")
            dl_ok = self._run_download()
            dbg(f"[Thread] Téléchargement résultat : {dl_ok}")

            if self._download_cancel.is_set():
                self.after(0, lambda: self._dl_set_status(
                    "Annulé par l'utilisateur", COL["status_warn"]))
                return

            elapsed = time.time() - start_time
            dbg(f"[Thread] Téléchargement terminé en {elapsed:.1f}s", "OK")

            if self._download_error_detected:
                self.after(0, lambda: self._dl_set_status(
                    "Téléchargement échoué — erreur détectée", COL["status_err"]))
                return
            if not dl_ok:
                self.after(0, lambda: self._dl_set_status(
                    "Téléchargement échoué — vérifiez les logs", COL["status_err"]))
                return

            # ── Vérification post-téléchargement ──
            dbg("[Thread] Vérification post-téléchargement…", "STEP")
            source_found = threading.Event()
            source_result = [None]

            def _check_source():
                source_result[0] = self._scan_source()
                source_found.set()
            self.after(0, _check_source)
            source_found.wait(timeout=10)

            if source_result[0] is None:
                self.after(0, lambda: self._dl_set_status(
                    "Téléchargement terminé mais aucune source trouvée", COL["status_err"]))
                return

            dbg(f"[Thread] Source trouvée : {source_result[0]}", "OK")
            self.after(0, lambda: self._dl_set_status(
                "Téléchargement terminé — lancement installation…", COL["status_ok"]))

            # ── Phase 3 : Installation ──
            time.sleep(1.5)
            dbg("[Thread] === PHASE 3 : Installation ===", "STEP")
            self._run_install()

            elapsed = time.time() - start_time
            dbg(f"[Thread] Séquence complète terminée en {elapsed:.1f}s", "OK")

        except Exception as e:
            dbg(f"[Thread] EXCEPTION : {e}", "ERR")
            dbg(traceback.format_exc(), "ERR")
            self.after(0, lambda e=e: self._dl_set_status(
                f"Erreur inattendue : {e}", COL["status_err"]))
        finally:
            dbg("[Thread] _t_download_and_install() terminé, nettoyage")
            self.after(0, self._on_worker_done)

    def _t_install_only(self):
        dbg("[Thread] _t_install_only() démarré", "STEP")
        try:
            self.after(0, lambda: self._dl_set_status(
                "Installation en cours…", COL["status_warn"]))
            self._run_install()
        except Exception as e:
            dbg(f"[Thread] EXCEPTION dans _t_install_only : {e}", "ERR")
            dbg(traceback.format_exc(), "ERR")
            self.after(0, lambda e=e: self._dl_set_status(
                f"Erreur : {e}", COL["status_err"]))
        finally:
            dbg("[Thread] _t_install_only() terminé")
            self.after(0, self._on_worker_done)

    # ──────────────────────────────────────────
    # Génération script aria2
    # ──────────────────────────────────────────

    def _run_generator(self) -> bool:
        dbg("_run_generator() démarré", "STEP")
        script_path = self.scripts.get("YAOCTRU_Generator.cmd")
        if not script_path:
            script_path = find_script("YAOCTRU_Generator.cmd")
        if not script_path:
            dbg("  YAOCTRU_Generator.cmd introuvable !", "ERR")
            self.after(0, lambda: self._dl_set_status(
                "YAOCTRU_Generator.cmd introuvable", COL["status_err"]))
            return False

        work_dir = _get_work_dir(script_path)
        dbg(f"  Script : {script_path}")
        dbg(f"  Work dir : {work_dir}")
        os.makedirs(work_dir, exist_ok=True)

        inputs = [
            CHANNELS_YAOCTRU[self.combo_channel.get()],
            BUILDS[self.combo_build.get()],
            BITNESS[self.combo_bitness.get()],
            LANGUAGES[self.combo_lang.get()],
            DL_TYPES[self.combo_dl_type.get()],
            1,
        ]
        triggers = [
            "Enter Channel option", "Enter Build option",
            "Enter Bitness option", "Enter Language option",
            "Enter Download option", "Enter Output option",
        ]

        dbg(f"  Inputs à envoyer : {inputs}")
        dbg(f"  Triggers attendus : {triggers}")

        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        try:
            dbg(f"  Lancement subprocess : {script_path}")
            p = subprocess.Popen(
                [script_path], cwd=work_dir,
                stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, shell=True, text=True,
                encoding='cp850', bufsize=0, startupinfo=si)
            dbg(f"  Process PID={p.pid} lancé", "OK")
        except Exception as e:
            dbg(f"  Erreur lancement : {e}", "ERR")
            self.after(0, lambda e=e: self._dl_set_status(
                f"Impossible de lancer le générateur : {e}", COL["status_err"]))
            return False

        step = 0
        buf = ""
        chars_read = 0
        while True:
            if self._download_cancel.is_set():
                dbg("  Annulation détectée pendant génération", "WARN")
                try:
                    p.terminate()
                except Exception:
                    pass
                return False

            char = p.stdout.read(1)
            if not char:
                if p.poll() is not None:
                    break
                continue
            buf += char
            chars_read += 1

            if step < len(inputs) and triggers[step] in buf:
                val = str(inputs[step])
                dbg(f"  Trigger détecté [{step}] : '{triggers[step]}' -> envoi '{val}'")
                try:
                    p.stdin.write(val + "\n")
                    p.stdin.flush()
                except Exception as e:
                    dbg(f"  Erreur écriture stdin : {e}", "ERR")
                step += 1
                buf = ""

            low = buf.lower()
            if "press any key" in low or "appuyez" in low:
                dbg("  'Press any key' détecté -> envoi Enter")
                try:
                    p.stdin.write("\n")
                    p.stdin.flush()
                except Exception:
                    pass
                buf = ""

            if char == '\n':
                stripped = buf.strip()
                if stripped and len(stripped) > 3:
                    clean = _strip_ansi(stripped)
                    if clean and not clean.isspace():
                        dbg(f"  [GEN] {clean}")
                buf = ""

        retcode = p.wait()
        dbg(f"  Générateur terminé : retcode={retcode}, chars lus={chars_read}")
        return retcode == 0

    # ──────────────────────────────────────────
    # Téléchargement
    # ──────────────────────────────────────────

    def _run_download(self) -> bool:
        dbg("_run_download() démarré", "STEP")
        script_path = self.scripts.get("YAOCTRU_Generator.cmd")
        if not script_path:
            script_path = find_script("YAOCTRU_Generator.cmd")
        work_dir = _get_work_dir(script_path)

        suffix = "_aria2.bat"
        try:
            all_files = os.listdir(work_dir)
            bat_files = [os.path.join(work_dir, f) for f in all_files if f.endswith(suffix)]
            dbg(f"  Fichiers *{suffix} dans {work_dir} : "
                f"{[os.path.basename(f) for f in bat_files]}")
        except Exception as e:
            dbg(f"  Erreur listdir : {e}", "ERR")
            bat_files = []

        if not bat_files:
            dbg("  Aucun script aria2.bat trouvé !", "ERR")
            self.after(0, lambda: self._dl_set_status(
                "Script de téléchargement non généré", COL["status_err"]))
            return False

        dl_script = max(bat_files, key=os.path.getmtime)
        total_files = _count_urls_in_bat(dl_script)

        dbg(f"  Script sélectionné : {dl_script}")
        dbg(f"  Nombre de fichiers à télécharger : {total_files}")

        self.after(0, lambda: self._dl_set_status(
            f"Téléchargement de {total_files} fichier(s)…", COL["status_warn"]))
        self.after(0, lambda: self._dl_set_percent(0, total_files))

        si = subprocess.STARTUPINFO()
        si.dwFlags |= subprocess.STARTF_USESHOWWINDOW

        try:
            cmd = ["cmd", "/c", dl_script]
            dbg(f"  Lancement : {cmd}")
            p_dl = subprocess.Popen(
                cmd, cwd=work_dir,
                stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT, text=True,
                encoding='cp850', bufsize=0, startupinfo=si)
            dbg(f"  Process DL PID={p_dl.pid} lancé", "OK")
        except Exception as e:
            dbg(f"  Erreur lancement DL : {e}", "ERR")
            self.after(0, lambda e=e: self._dl_set_status(
                f"Impossible de lancer le téléchargement : {e}", COL["status_err"]))
            return False

        current_file = 1
        line_buf = ""
        last_pct_logged = -5

        while True:
            if self._download_cancel.is_set():
                dbg("  Annulation détectée pendant téléchargement", "WARN")
                try:
                    p_dl.terminate()
                except Exception:
                    pass
                return False

            char = p_dl.stdout.read(1)
            if not char:
                if p_dl.poll() is not None:
                    break
                continue

            line_buf += char

            # Détecter erreurs
            if char == '\n' or len(line_buf) > 500:
                lower_line = line_buf.lower()
                if ("error" in lower_line or "is not detected" in lower_line or
                        "not found" in lower_line or "not recognized" in lower_line):
                    clean = _strip_ansi(line_buf.strip())
                    dbg(f"  [DL] ERREUR DÉTECTÉE : {clean}", "ERR")
                    self._download_error_detected = True

            if "download complete" in line_buf.lower():
                dbg(f"  Fichier {current_file}/{total_files} terminé", "OK")
                c, t = current_file, total_files
                self.after(0, lambda c=c, t=t: self.dl_percent.configure(
                    text=f"100 % ({c}/{t})"))
                if current_file < total_files:
                    current_file += 1
                line_buf = ""

            m = re.search(r'\(\s*(\d+(?:\.\d+)?)\s*%\)', line_buf)
            if m:
                val = float(m.group(1))
                c, t = current_file, total_files
                self.after(0, lambda v=val: self.dl_progress.set(v / 100))
                self.after(0, lambda v=val, c=c, t=t: self.dl_percent.configure(
                    text=f"{v:.1f} % ({c - 1}/{t})"))
                if val - last_pct_logged >= 5 or val >= 99.9:
                    dbg(f"  Progression : {val:.1f}% (fichier {c}/{t})")
                    last_pct_logged = val

            tail = line_buf[-80:].lower() if len(line_buf) > 80 else line_buf.lower()
            if "press any key" in tail or "appuyez" in tail:
                if self._download_error_detected:
                    dbg("  Erreur détectée avant 'press any key'", "WARN")
                dbg("  'Press any key' détecté -> envoi Enter")
                try:
                    p_dl.stdin.write("\n")
                    p_dl.stdin.flush()
                except Exception:
                    pass
                line_buf = ""

            elif char == '\n':
                stripped = line_buf.strip()
                if stripped and ('[#' in stripped or 'error' in stripped.lower()
                                or 'fail' in stripped.lower()):
                    clean = _strip_ansi(stripped)
                    if len(clean) > 5:
                        dbg(f"  [DL] {clean[:120]}")
                line_buf = ""

        retcode = p_dl.wait()
        dbg(f"  Processus DL terminé : retcode={retcode}")

        if self._download_error_detected:
            dbg("  Téléchargement terminé AVEC erreurs", "ERR")
            try:
                os.remove(dl_script)
                dbg(f"  Script nettoyé : {dl_script}", "OK")
            except Exception:
                pass
            return False

        if retcode == 0:
            self.after(0, lambda: self.dl_progress.set(1))
            self.after(0, lambda t=total_files: self.dl_percent.configure(
                text=f"100 % ({t}/{t})"))
            try:
                os.remove(dl_script)
                dbg(f"  Script nettoyé : {dl_script}", "OK")
            except Exception as e:
                dbg(f"  Erreur nettoyage script : {e}", "WARN")
            return True
        else:
            dbg(f"  Téléchargement échoué (retcode={retcode})", "ERR")
            try:
                os.remove(dl_script)
            except Exception:
                pass
            return False

    # ──────────────────────────────────────────
    # Installation
    # ──────────────────────────────────────────

    def _run_install(self) -> bool:
        dbg("_run_install() démarré", "STEP")

        if not self.source_valid or not self._c2r_path:
            dbg("  Pas de source valide !", "ERR")
            self.after(0, lambda: self._dl_set_status(
                "Aucune source valide pour l'installation", COL["status_err"]))
            return False

        dbg(f"  Source : {self._c2r_path}")
        dbg(f"  Suite  : {self.selected_suite_id.get()}")

        # ── Snapshot AVANT installation ──
        pre_info = get_installed_office_info()
        pre_version = pre_info["version"] if pre_info else ""
        dbg(f"  Office AVANT installation : "
            f"{'v' + pre_version if pre_version else 'NON installé'}")

        # ── Nettoyage préalable si résidus détectés sans Office réel ──
        if pre_version and not self._is_office_really_installed():
            dbg("  Résidus registre détectés sans Office réel -> nettoyage", "WARN")
            self.after(0, lambda: self._dl_set_status(
                "Nettoyage des résidus d'une ancienne installation…",
                COL["status_warn"]))
            self._full_cleanup()
            time.sleep(2)
            pre_info = get_installed_office_info()
            pre_version = pre_info["version"] if pre_info else ""

        # ── 1. Trouver setup.exe ──
        dbg("  Étape 1/4 : Recherche setup.exe", "STEP")
        setup_exe = self._find_setup_exe()
        if not setup_exe:
            dbg("  setup.exe introuvable dans la source !", "ERR")
            self.after(0, lambda: self._dl_set_status(
                "setup.exe introuvable dans la source Office",
                COL["status_err"]))
            return False
        dbg(f"  setup.exe trouvé : {setup_exe}", "OK")

        # ── 2. Générer XML ODT ──
        dbg("  Étape 2/4 : Génération XML ODT", "STEP")
        xml_path = self._generate_odt_xml()
        if not xml_path:
            dbg("  Échec génération XML", "ERR")
            self.after(0, lambda: self._dl_set_status(
                "Échec de la génération du fichier XML", COL["status_err"]))
            return False
        dbg(f"  XML généré : {xml_path}", "OK")

        # ── Afficher le XML ──
        try:
            with open(xml_path, "r", encoding="utf-8") as f:
                dbg("  Contenu XML :")
                for line in f:
                    dbg(f"    {line.rstrip()}")
        except Exception:
            pass

        # ── 3. Lancer setup.exe /configure ──
        dbg("  Étape 3/4 : Lancement setup.exe /configure", "STEP")
        self.after(0, lambda: self._dl_set_status(
            "Installation d'Office en cours (élévation admin)…",
            COL["status_warn"]))

        script_dir = os.path.dirname(os.path.abspath(__file__))
        wrapper_path = os.path.join(script_dir, "_auto_install.bat")
        install_log = os.path.join(script_dir, "_install_output.log")

        try:
            with open(wrapper_path, "w", encoding="cp850") as f:
                f.write("@echo off\r\n")
                f.write(f'cd /d "{os.path.dirname(setup_exe)}"\r\n')
                f.write(f'echo [%date% %time%] Lancement setup.exe /configure > "{install_log}"\r\n')
                f.write(f'echo Setup: {setup_exe} >> "{install_log}"\r\n')
                f.write(f'echo XML: {xml_path} >> "{install_log}"\r\n')
                f.write(f'echo. >> "{install_log}"\r\n')
                f.write(f'"{setup_exe}" /configure "{xml_path}" >> "{install_log}" 2>&1\r\n')
                f.write("set EC=%errorlevel%\r\n")
                f.write(f'echo. >> "{install_log}"\r\n')
                f.write(f'echo Exit code: %EC% >> "{install_log}"\r\n')
                f.write("exit /b %EC%\r\n")
            dbg(f"  Wrapper créé : {wrapper_path}", "OK")

            with open(wrapper_path, "r", encoding="cp850") as f:
                dbg(f"  Contenu wrapper:\n{f.read()}")

        except Exception as e:
            dbg(f"  Erreur création wrapper : {e}", "ERR")
            self._cleanup_file(xml_path)
            return False

        t_start = time.time()
        install_ok = False

        try:
            exit_code = self._run_as_admin(wrapper_path, script_dir, wait=True)
            elapsed_setup = time.time() - t_start
            dbg(f"  setup.exe terminé : exit_code={exit_code} en {elapsed_setup:.1f}s")

            if exit_code == -1:
                dbg("  Élévation refusée", "ERR")
                self.after(0, lambda: self._dl_set_status(
                    "Installation annulée — élévation admin refusée",
                    COL["status_err"]))
                self._cleanup_file(xml_path)
                self._cleanup_file(wrapper_path)
                return False
        except Exception as e:
            dbg(f"  Erreur installation : {e}", "ERR")
            self._cleanup_file(xml_path)
            self._cleanup_file(wrapper_path)
            return False

        self._cleanup_file(wrapper_path)

        # ── Lire la sortie ──
        if os.path.isfile(install_log):
            try:
                with open(install_log, "r", encoding="cp850", errors="replace") as f:
                    content = f.read()
                dbg("  ╔══ Sortie setup.exe ══", "STEP")
                for line in content.splitlines():
                    s = line.rstrip()
                    if s:
                        dbg(f"  ║ {s}")
                dbg("  ╚══ Fin sortie ══", "STEP")
            except Exception as e:
                dbg(f"  Erreur lecture log : {e}", "ERR")
            self._cleanup_file(install_log)

        # ── Attente du streaming C2R ──
        if exit_code == 0 or elapsed_setup < 30:
            dbg("  Attente de la fin du streaming C2R…", "STEP")
            self.after(0, lambda: self._dl_set_status(
                "Installation en cours — streaming des composants…",
                COL["status_warn"]))

            max_wait = 600
            poll_interval = 5
            waited = 0
            last_log = 0

            while waited < max_wait:
                if self._download_cancel.is_set():
                    break

                if self._is_office_really_installed():
                    dbg(f"  Office RÉELLEMENT installé après {waited}s", "OK")
                    break

                c2r_running = self._is_c2r_process_running()

                if not c2r_running and waited > 30:
                    dbg(f"  Processus C2R terminé après {waited}s", "WARN")
                    break

                time.sleep(poll_interval)
                waited += poll_interval

                if waited - last_log >= 15:
                    status = "C2R actif" if c2r_running else "en attente"
                    dbg(f"  Attente streaming : {waited}s ({status})…")
                    self.after(0, lambda w=waited, s=status:
                               self._dl_set_status(
                                   f"Installation en cours… {w}s ({s})",
                                   COL["status_warn"]))
                    last_log = waited

            time.sleep(3)

        # ── Vérification post-installation ──
        dbg("  Vérification post-installation…", "STEP")
        post_info = get_installed_office_info()
        post_version = post_info["version"] if post_info else ""
        really_installed = self._is_office_really_installed()

        dbg(f"  APRÈS : version='{post_version}', "
            f"réellement installé={really_installed}, exit_code={exit_code}")

        if really_installed and post_version:
            dbg(f"  Installation réussie : v{post_version}", "OK")
            install_ok = True
            self.after(0, lambda v=post_version: self._dl_set_status(
                f"✓ Installation réussie — Office v{v}", COL["status_ok"]))

        elif post_version and not really_installed:
            if self._is_c2r_process_running():
                dbg("  C2R encore actif, attente supplémentaire…", "WARN")
                for _ in range(24):
                    time.sleep(5)
                    if self._is_office_really_installed():
                        install_ok = True
                        dbg(f"  Office finalement installé", "OK")
                        self.after(0, lambda v=post_version: self._dl_set_status(
                            f"✓ Installation réussie — Office v{v}",
                            COL["status_ok"]))
                        break
                    if not self._is_c2r_process_running():
                        break
                if not install_ok:
                    self._full_cleanup()
                    self.after(0, lambda: self._dl_set_status(
                        "Installation incomplète", COL["status_warn"]))
            else:
                self._full_cleanup()
                self.after(0, lambda: self._dl_set_status(
                    "Installation échouée — Office non détecté",
                    COL["status_err"]))
        else:
            self.after(0, lambda ec=exit_code: self._dl_set_status(
                f"Installation échouée (code {ec})",
                COL["status_err"]))

        # ── Nettoyage XML ──
        self._cleanup_file(xml_path)

        # ── 4. Activation Ohook ──
        should_activate = self.opt_activate.get()
        dbg(f"  Étape 4/4 : Activation Ohook = {should_activate} "
            f"(install_ok={install_ok})", "STEP")

        if should_activate and install_ok:
            time.sleep(2)
            self._run_ohook_activation()
        elif not install_ok:
            dbg("  Activation ignorée (installation échouée)")
        else:
            self.after(0, lambda: self._dl_set_status(
                "Installation terminée (activation désactivée)",
                COL["status_ok"]))

        dbg("  Planification revérification activation dans 3s")
        time.sleep(3)
        self.after(0, self._on_check_activation)
        return install_ok

    def _is_office_really_installed(self) -> bool:
        """Vérifie qu'Office est RÉELLEMENT installé (pas juste des résidus registre)"""
        dbg("_is_office_really_installed() vérification…")

        # Vérifier que les exécutables principaux existent
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")

        exe_paths = [
            os.path.join(pf, "Microsoft Office", "root", "Office16", "WINWORD.EXE"),
            os.path.join(pf86, "Microsoft Office", "root", "Office16", "WINWORD.EXE"),
            os.path.join(pf, "Microsoft Office", "root", "Office16", "EXCEL.EXE"),
            os.path.join(pf86, "Microsoft Office", "root", "Office16", "EXCEL.EXE"),
        ]

        # Aussi chercher via le registre InstallationPath
        try:
            for reg_key in [
                r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun\Configuration",
                r"HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun\Configuration",
            ]:
                cmd = f'reg query "{reg_key}" /v InstallationPath'
                result = subprocess.run(cmd, capture_output=True, text=True,
                                        shell=True, creationflags=0x08000000, timeout=5)
                if result.returncode == 0:
                    m = re.search(r'REG_SZ\s+(.+)', result.stdout)
                    if m:
                        install_path = m.group(1).strip()
                        dbg(f"  InstallationPath : {install_path}")
                        exe_paths.insert(0, os.path.join(
                            install_path, "root", "Office16", "WINWORD.EXE"))
                        exe_paths.insert(1, os.path.join(
                            install_path, "root", "Office16", "EXCEL.EXE"))
                        break
        except Exception as e:
            dbg(f"  Erreur lecture InstallationPath : {e}", "ERR")

        for exe in exe_paths:
            if os.path.isfile(exe):
                dbg(f"  Exécutable trouvé : {exe} -> RÉELLEMENT installé", "OK")
                return True

        # Vérifier aussi le service ClickToRun
        try:
            result = subprocess.run(
                'sc query "ClickToRunSvc"',
                capture_output=True, text=True, shell=True,
                creationflags=0x08000000, timeout=5)
            if result.returncode == 0 and "RUNNING" in result.stdout.upper():
                dbg("  Service ClickToRunSvc actif -> RÉELLEMENT installé", "OK")
                return True
        except Exception:
            pass

        dbg("  Aucun exécutable Office trouvé -> RÉSIDUS seulement", "WARN")
        return False

    def _is_c2r_process_running(self) -> bool:
        """Vérifie si un processus d'installation C2R est en cours"""
        c2r_processes = [
            "OfficeClickToRun.exe",
            "OfficeC2RClient.exe",
            "setup.exe",
        ]
        for proc_name in c2r_processes:
            try:
                result = subprocess.run(
                    f'tasklist /FI "IMAGENAME eq {proc_name}" /NH',
                    capture_output=True, text=True, shell=True,
                    creationflags=0x08000000, timeout=5)
                if proc_name.lower() in result.stdout.lower():
                    return True
            except Exception:
                pass
        return False

    def _full_cleanup(self):
        """Nettoyage complet : registre + fichiers + services + tâches planifiées"""
        dbg("_full_cleanup() nettoyage complet démarré", "STEP")

        script_dir = os.path.dirname(os.path.abspath(__file__))
        pf = os.environ.get("ProgramFiles", r"C:\Program Files")
        pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        common = os.environ.get("CommonProgramFiles", "")
        common86 = os.environ.get("CommonProgramFiles(x86)", "")
        prog_data = os.environ.get("ProgramData", r"C:\ProgramData")
        local_app = os.environ.get("LOCALAPPDATA", "")
        appdata = os.environ.get("APPDATA", "")

        wrapper_path = os.path.join(script_dir, "_full_cleanup.bat")

        try:
            with open(wrapper_path, "w", encoding="cp850") as f:
                f.write("@echo off\r\n")
                f.write("echo === Nettoyage complet Office ===\r\n\r\n")

                # 1. Arrêter les services
                f.write("echo [1/6] Arrêt des services...\r\n")
                services = ["ClickToRunSvc", "ose64", "ose",
                            "OfficeSvc", "Microsoft Office Click-to-Run"]
                for svc in services:
                    f.write(f'net stop "{svc}" /y >nul 2>&1\r\n')
                    f.write(f'sc stop "{svc}" >nul 2>&1\r\n')

                # 2. Tuer les processus
                f.write("\r\necho [2/6] Fermeture des processus...\r\n")
                processes = [
                    "OfficeClickToRun.exe", "OfficeC2RClient.exe",
                    "AppVShNotify.exe", "WINWORD.EXE", "EXCEL.EXE",
                    "POWERPNT.EXE", "OUTLOOK.EXE", "ONENOTE.EXE",
                    "MSACCESS.EXE", "MSPUB.EXE", "lync.exe",
                    "Teams.exe", "setup.exe",
                ]
                for proc in processes:
                    f.write(f'taskkill /F /IM "{proc}" >nul 2>&1\r\n')
                f.write("timeout /t 3 /nobreak >nul\r\n")

                # 3. Supprimer les services
                f.write("\r\necho [3/6] Suppression des services...\r\n")
                for svc in services:
                    f.write(f'sc delete "{svc}" >nul 2>&1\r\n')

                # 4. Nettoyage registre complet
                f.write("\r\necho [4/6] Nettoyage registre...\r\n")
                reg_keys = [
                    # Clés principales ClickToRun
                    r"HKLM\SOFTWARE\Microsoft\Office\ClickToRun",
                    r"HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\ClickToRun",
                    # Clés Office 16.0
                    r"HKLM\SOFTWARE\Microsoft\Office\16.0",
                    r"HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\16.0",
                    r"HKCU\SOFTWARE\Microsoft\Office\16.0",
                    # Clés de désinstallation
                    r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365HomePremRetail - fr-fr",
                    r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365HomePremRetail - en-us",
                    r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - fr-fr",
                    r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us",
                    r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365BusinessRetail - fr-fr",
                    r"HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365BusinessRetail - en-us",
                    # Clés AppVISV
                    r"HKLM\SOFTWARE\Microsoft\AppVISV",
                    r"HKLM\SOFTWARE\Wow6432Node\Microsoft\AppVISV",
                    # Clés Office commune
                    r"HKLM\SOFTWARE\Microsoft\Office\Delivery",
                    r"HKLM\SOFTWARE\Microsoft\Office\MS#",
                    # Clés registre utilisateur
                    r"HKCU\SOFTWARE\Microsoft\Office\16.0\Common\Licensing",
                    r"HKCU\SOFTWARE\Microsoft\Office\16.0\Registration",
                ]
                for key in reg_keys:
                    f.write(f'reg delete "{key}" /f >nul 2>&1\r\n')

                # Aussi nettoyer toutes les clés Uninstall Office
                f.write('for /f "tokens=*" %%i in (\'reg query '
                        '"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion'
                        '\\Uninstall" /f "Office" /k 2^>nul ^| findstr /i '
                        '"HKEY_"\') do (\r\n')
                f.write('  reg delete "%%i" /f >nul 2>&1\r\n')
                f.write(')\r\n')

                f.write('for /f "tokens=*" %%i in (\'reg query '
                        '"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion'
                        '\\Uninstall" /f "O365" /k 2^>nul ^| findstr /i '
                        '"HKEY_"\') do (\r\n')
                f.write('  reg delete "%%i" /f >nul 2>&1\r\n')
                f.write(')\r\n')

                f.write('for /f "tokens=*" %%i in (\'reg query '
                        '"HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion'
                        '\\Uninstall" /f "Microsoft 365" /k 2^>nul ^| findstr /i '
                        '"HKEY_"\') do (\r\n')
                f.write('  reg delete "%%i" /f >nul 2>&1\r\n')
                f.write(')\r\n')

                # 5. Suppression des dossiers
                f.write("\r\necho [5/6] Suppression des fichiers...\r\n")
                dirs_to_clean = [
                    os.path.join(pf, "Microsoft Office"),
                    os.path.join(pf86, "Microsoft Office"),
                    os.path.join(common, "Microsoft Shared", "ClickToRun"),
                    os.path.join(common86, "Microsoft Shared", "ClickToRun"),
                    os.path.join(common, "Microsoft Shared", "OfficeSoftwareProtectionPlatform"),
                    os.path.join(prog_data, "Microsoft", "ClickToRun"),
                    os.path.join(prog_data, "Microsoft", "Office"),
                ]
                if local_app:
                    dirs_to_clean.append(
                        os.path.join(local_app, "Microsoft", "Office"))
                if appdata:
                    dirs_to_clean.append(
                        os.path.join(appdata, "Microsoft", "Office"))

                for d in dirs_to_clean:
                    if d:
                        f.write(f'if exist "{d}" (\r\n')
                        f.write(f'  echo   Suppression : {d}\r\n')
                        f.write(f'  rmdir /s /q "{d}" 2>nul\r\n')
                        f.write(f')\r\n')

                # 6. Nettoyage raccourcis et tâches planifiées
                f.write("\r\necho [6/6] Nettoyage raccourcis et taches...\r\n")
                start_menu = os.path.join(
                    prog_data, "Microsoft", "Windows",
                    "Start Menu", "Programs")
                shortcuts = [
                    "Word.lnk", "Excel.lnk", "PowerPoint.lnk",
                    "Outlook.lnk", "OneNote.lnk", "Access.lnk",
                    "Publisher.lnk", "Microsoft Office Tools",
                ]
                for s in shortcuts:
                    sp = os.path.join(start_menu, s)
                    f.write(f'if exist "{sp}" (\r\n')
                    # Gérer fichier ou dossier
                    f.write(f'  if exist "{sp}\\*" (\r\n')
                    f.write(f'    rmdir /s /q "{sp}" 2>nul\r\n')
                    f.write(f'  ) else (\r\n')
                    f.write(f'    del /f /q "{sp}" 2>nul\r\n')
                    f.write(f'  )\r\n')
                    f.write(f')\r\n')

                # Tâches planifiées Office
                f.write('schtasks /delete /tn "\\Microsoft\\Office\\Office Automatic Updates 2.0" /f >nul 2>&1\r\n')
                f.write('schtasks /delete /tn "\\Microsoft\\Office\\Office ClickToRun Service Monitor" /f >nul 2>&1\r\n')
                f.write('schtasks /delete /tn "\\Microsoft\\Office\\Office Feature Updates" /f >nul 2>&1\r\n')
                f.write('schtasks /delete /tn "\\Microsoft\\Office\\Office Feature Updates Logon" /f >nul 2>&1\r\n')
                f.write('schtasks /delete /tn "\\Microsoft\\Office\\OfficeTelemetryAgentFallBack2016" /f >nul 2>&1\r\n')
                f.write('schtasks /delete /tn "\\Microsoft\\Office\\OfficeTelemetryAgentLogOn2016" /f >nul 2>&1\r\n')

                f.write("\r\necho === Nettoyage termine ===\r\n")
                f.write("exit /b 0\r\n")

            dbg(f"  Script nettoyage créé : {wrapper_path}", "OK")

            exit_code = self._run_as_admin(wrapper_path, script_dir, wait=True)
            dbg(f"  Nettoyage complet terminé : exit_code={exit_code}")

        except Exception as e:
            dbg(f"  Erreur nettoyage complet : {e}", "ERR")
            dbg(traceback.format_exc(), "ERR")
        finally:
            self._cleanup_file(wrapper_path)

        # Vérification post-nettoyage
        time.sleep(2)
        post_clean = get_installed_office_info()
        if post_clean:
            dbg(f"  ⚠ Résidus encore présents après nettoyage : "
                f"v{post_clean['version']}", "WARN")
        else:
            dbg("  Nettoyage complet réussi — aucun résidu", "OK")

    def _run_ohook_activation(self):
        self.after(0, lambda: self._dl_set_status(
            "Activation Office (Ohook)…", COL["status_warn"]))

        ohook_path = self.scripts.get("Ohook_Activation_AIO.cmd")
        if not ohook_path:
            ohook_path = find_script("Ohook_Activation_AIO.cmd")

        dbg(f"  Ohook script : {ohook_path}")

        if not ohook_path:
            dbg("  Ohook_Activation_AIO.cmd introuvable", "WARN")
            self.after(0, lambda: self._dl_set_status(
                "Installation OK — Ohook introuvable", COL["status_warn"]))
            return

        ohook_dir = os.path.dirname(ohook_path)
        wrapper_path = os.path.join(ohook_dir, "_auto_activate.bat")
        activate_log = os.path.join(ohook_dir, "_activate_output.log")

        try:
            with open(wrapper_path, "w", encoding="cp850") as f:
                f.write("@echo off\r\n")
                f.write(f'cd /d "{ohook_dir}"\r\n')
                f.write(f'call "{ohook_path}" /Ohook > "{activate_log}" 2>&1\r\n')
                f.write("exit /b %errorlevel%\r\n")
            dbg(f"  Wrapper activation créé : {wrapper_path}", "OK")
        except Exception as e:
            dbg(f"  Erreur création wrapper : {e}", "ERR")
            self.after(0, lambda e=e: self._dl_set_status(
                f"Erreur activation : {e}", COL["status_err"]))
            return

        try:
            ohook_exit = self._run_as_admin(wrapper_path, ohook_dir, wait=True)
            dbg(f"  Ohook terminé : exit_code={ohook_exit}")

            # Lire la sortie
            if os.path.isfile(activate_log):
                try:
                    with open(activate_log, "r", encoding="cp850",
                              errors="replace") as f:
                        content = f.read()
                    dbg("  ╔══ Sortie Ohook ══", "STEP")
                    for line in content.splitlines():
                        s = line.rstrip()
                        if s:
                            dbg(f"  ║ {s}")
                    dbg("  ╚══ Fin sortie Ohook ══", "STEP")
                except Exception as e:
                    dbg(f"  Erreur lecture log Ohook : {e}", "ERR")
                self._cleanup_file(activate_log)

            time.sleep(5)

            if ohook_exit == 0:
                self.after(0, lambda: self._dl_set_status(
                    "✓ Installation et activation terminées",
                    COL["status_ok"]))
            elif ohook_exit == -1:
                self.after(0, lambda: self._dl_set_status(
                    "Installation OK — activation annulée (UAC)",
                    COL["status_warn"]))
            else:
                self.after(0, lambda ec=ohook_exit: self._dl_set_status(
                    f"Installation OK — activation code {ec}",
                    COL["status_warn"]))
        except Exception as e:
            dbg(f"  Erreur activation Ohook : {e}", "ERR")
            self.after(0, lambda e=e: self._dl_set_status(
                f"Erreur activation : {e}", COL["status_err"]))
        finally:
            self._cleanup_file(wrapper_path)

    # ──────────────────────────────────────────
    # Génération INI
    # ──────────────────────────────────────────

    def _find_setup_exe(self) -> str | None:
        """Trouve setup.exe dans la source Office"""
        dbg("_find_setup_exe() recherche…")
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        
        candidates = [
            os.path.join(script_dir, "setup.exe"),
            os.path.join(script_dir, "Downloads", "setup.exe"),
            os.path.join(self._c2r_path, "setup.exe"),
            os.path.join(self._c2r_path, "Office", "setup.exe"),
            os.path.join(self._c2r_path, "Office", "Data", "setup.exe"),
        ]
        for p in candidates:
            dbg(f"  Candidat : {p} -> {'EXISTE' if os.path.isfile(p) else 'absent'}")
            if os.path.isfile(p):
                return p

        # Chercher récursivement
        for root, dirs, files in os.walk(self._c2r_path):
            for f in files:
                if f.lower() == "setup.exe":
                    found = os.path.join(root, f)
                    dbg(f"  Trouvé (recherche récursive) : {found}", "OK")
                    return found

        dbg("  setup.exe NON TROUVÉ", "ERR")
        return None

    def _generate_odt_xml(self) -> str | None:
        """Génère le XML de configuration Office Deployment Tool"""
        dbg("_generate_odt_xml() démarré", "STEP")

        CHANNEL_ODT_MAP = {
            "Monthly": "Current",
            "MonthlyPreview": "CurrentPreview",
            "Broad": "SemiAnnual",
            "Targeted": "SemiAnnualPreview",
            "Beta": "BetaChannel",
            "Dogfood": "BetaChannel",
            "PerpetualVL2019": "PerpetualVL2019",
            "PerpetualVL2021": "PerpetualVL2021",
            "PerpetualVL2024": "PerpetualVL2024",
        }

        try:
            ver, arch, lang = "??", "x64", "fr-FR"
            data_p = os.path.join(self._c2r_path, "Office", "Data")
            cabs = glob.glob(os.path.join(data_p, "v*_*.cab"))
            if cabs:
                parts = os.path.basename(cabs[0]).split("_")
                if len(parts) >= 2:
                    arch = "x64" if "64" in parts[0] else "x86"
                    ver = parts[1].replace(".cab", "")
                vf = os.path.join(data_p, ver)
                if os.path.isdir(vf):
                    for s in os.listdir(vf):
                        if s.startswith("stream.") and s.endswith(".dat"):
                            pp = s.split(".")
                            if len(pp) >= 3 and pp[2] not in ("x-none", "dat"):
                                lang = pp[2]
                                break
            dbg(f"  Source parsée : ver={ver}, arch={arch}, lang={lang}")

            edition = "64" if arch == "x64" else "32"
            suite_id = self.selected_suite_id.get()
            channel_short = self.combo_update_channel.get()
            channel_odt = CHANNEL_ODT_MAP.get(channel_short, "Current")
            src = self._c2r_path.replace("/", "\\")

            dbg(f"  Canal UI : {channel_short} -> ODT : {channel_odt}")

            # Apps exclues
            excluded_apps = []
            for name, var in self.apps_vars.items():
                if not var.get():
                    excluded_apps.append(APPS_MAP[name])
            dbg(f"  Apps exclues : {excluded_apps}")

            exclude_lines = ""
            for app_id in excluded_apps:
                exclude_lines += f'      <ExcludeApp ID="{app_id}" />\n'

            # Pré-calculer les booléens en strings XML
            updates_enabled = "TRUE" if self.opt_updates.get() else "FALSE"
            accept_eula = "TRUE" if self.opt_eula.get() else "FALSE"
            force_shutdown = "TRUE" if self.opt_shutdown.get() else "FALSE"
            pin_taskbar = "TRUE" if self.opt_pin_taskbar.get() else "FALSE"
            display_level = "Full" if self.opt_display.get() else "None"

            xml = (
                '<Configuration>\n'
                f'  <Add SourcePath="{src}" OfficeClientEdition="{edition}"'
                f' Channel="{channel_odt}" Version="{ver}">\n'
                f'    <Product ID="{suite_id}">\n'
                f'      <Language ID="{lang}" />\n'
                f'{exclude_lines}'
                f'    </Product>\n'
                f'  </Add>\n'
                f'  <Updates Enabled="{updates_enabled}"'
                f' Channel="{channel_odt}" />\n'
                f'  <Display Level="{display_level}"'
                f' AcceptEULA="{accept_eula}" />\n'
                f'  <Property Name="FORCEAPPSHUTDOWN"'
                f' Value="{force_shutdown}" />\n'
                f'  <Property Name="PinIconsToTaskbar"'
                f' Value="{pin_taskbar}" />\n'
                f'  <Property Name="AUTOACTIVATE" Value="0" />\n'
                f'</Configuration>'
            )

            dbg(f"  Suite ID  : {suite_id}")
            dbg(f"  Canal     : {channel_odt}")
            dbg(f"  Source    : {src}")
            dbg(f"  Edition   : {edition}-bit")
            dbg(f"  Display   : {display_level}")

            script_dir = os.path.dirname(os.path.abspath(__file__))
            now = datetime.datetime.now().strftime('%Y%m%d-%H%M')
            xml_path = os.path.join(script_dir, f"_install_config_{now}.xml")

            with open(xml_path, "w", encoding="utf-8") as f:
                f.write(xml)
            dbg(f"  XML écrit : {xml_path}", "OK")
            return xml_path

        except Exception as e:
            dbg(f"  Exception génération XML : {e}", "ERR")
            dbg(traceback.format_exc(), "ERR")
            return None

    # ──────────────────────────────────────────
    # Callbacks UI (thread-safe)
    # ──────────────────────────────────────────

    def _dl_set_status(self, text: str, color: str):
        dbg(f"UI status -> '{text}' (color={color})")
        try:
            self.dl_status_label.configure(text=text, text_color=color)
        except Exception as e:
            dbg(f"  Erreur _dl_set_status : {e}", "ERR")

    def _dl_set_percent(self, pct: int, total: int):
        try:
            self.dl_progress.set(pct / 100 if pct > 0 else 0)
            self.dl_percent.configure(text=f"{pct} % (0/{total})")
        except Exception as e:
            dbg(f"  Erreur _dl_set_percent : {e}", "ERR")

    def _cleanup_file(self, path: str):
        try:
            if path and os.path.exists(path):
                os.remove(path)
                dbg(f"  Fichier nettoyé : {path}", "OK")
        except Exception as e:
            dbg(f"  Erreur nettoyage {path} : {e}", "WARN")

    def _on_worker_done(self):
        dbg("_on_worker_done() nettoyage état", "STEP")
        self._downloading = False
        self._installing = False
        self._download_error_detected = False
        self.btn_install.configure(state="normal")
        self.btn_cancel_dl.configure(state="disabled")
        self.btn_check.configure(state="normal")
        self.btn_download_only.configure(state="normal")
        self._scan_source()
        dbg("  Worker terminé, UI restaurée", "OK")

# ──────────────────────────────────────────────
# Point d'entrée
# ──────────────────────────────────────────────

if __name__ == "__main__":
    dbg("=" * 60)
    dbg("Démarrage du programme", "STEP")
    dbg(f"DEBUG = {DEBUG}")
    dbg("=" * 60)

    app = App()
    dbg("Entrée dans mainloop()")
    app.mainloop()
    dbg("Fin du programme", "OK")