import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox
import threading
import webbrowser
import time
import re
import json
import os
import sys
import random
import winsound
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager

# --- GLOBÁLNÍ KONFIGURACE ---
UIS_LOGIN_URL = "https://is.czu.cz/auth/"
OUTLOOK_URL = "https://outlook.office.com/mail/"
MOODLE_LOGIN_URL = "https://moodle.czu.cz/login/index.php"
COFFEE_URL = "https://buymeacoffee.com/colorvant"

def get_config_path():
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(application_path, "smart_sniper_config.json")

CONFIG_FILE = get_config_path()

# --- BARVY (DARK MODE) ---
COLOR_BG = "#1e1e1e"
COLOR_FRAME = "#2b2b2b"
COLOR_TEXT = "#ffffff"
COLOR_ENTRY_BG = "#3c3c3c"
COLOR_BTN_START = "#006400" 
COLOR_BTN_STOP = "#8b0000"  
COLOR_BTN_SCAN = "#005f9e"
COLOR_BTN_DOG = "#A0522D"
COLOR_ACCENT = "#FFD700"    
COLOR_INFO = "#4FC3F7"

# =============================================================================
# POMOCNÁ TŘÍDA PRO CONFIG
# =============================================================================
class ConfigManager:
    def load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f: return json.load(f)
            except: return {}
        return {}
    def save(self, data):
        try:
            # Načíst existující, aby se nepřepsala data druhé aplikace
            existing = self.load()
            existing.update(data)
            with open(CONFIG_FILE, "w", encoding="utf-8") as f: json.dump(existing, f, ensure_ascii=False, indent=4)
        except: pass

# =============================================================================
# TŘÍDA: LAUNCHER (ROZCESTNÍK)
# =============================================================================
class LauncherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Smart Sniper - ČZU")
        self.root.geometry("400x450")
        self.root.configure(bg=COLOR_BG)
        
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TButton", padding=10, font=("Segoe UI", 12, "bold"), background="#444", foreground="white", borderwidth=0)
        style.map("TButton", background=[('active', '#555')])

        tk.Label(root, text="Vyber nástroj", font=("Segoe UI", 20, "bold"), bg=COLOR_BG, fg=COLOR_TEXT).pack(pady=(40, 20))

        btn_uis = ttk.Button(root, text="UIS SNIPER (Zkoušky)", command=self.open_uis_sniper)
        btn_uis.pack(fill=tk.X, padx=50, pady=10)

        btn_tc = ttk.Button(root, text="TC SNIPER (Moodle Testy)", command=self.open_tc_sniper)
        btn_tc.pack(fill=tk.X, padx=50, pady=10)
        
        tk.Label(root, text="v2.0 UIS + TC", font=("Segoe UI", 8), bg=COLOR_BG, fg="gray").pack(side=tk.BOTTOM, pady=5)
        
        btn_coffee = tk.Button(root, text="☕ Podpořit autora", bg=COLOR_ACCENT, fg="black", font=("Segoe UI", 10, "bold"), command=lambda: webbrowser.open(COFFEE_URL))
        btn_coffee.pack(side=tk.BOTTOM, pady=10)

    def open_uis_sniper(self):
        new_window = tk.Toplevel(self.root)
        UISSniperApp(new_window)

    def open_tc_sniper(self):
        new_window = tk.Toplevel(self.root)
        TCSniperApp(new_window)

# =============================================================================
# TŘÍDA: UIS SNIPER (VZHLED ZE STARÉ VERZE, FUNKCE Z NOVÉ)
# =============================================================================
class UISSniperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("UIS Sniper - ČZU Dark Edition")
        self.root.geometry("700x980")
        self.root.resizable(True, True)
        self.root.configure(bg=COLOR_BG)
        
        self.driver = None
        self.is_running = False
        self.thread = None
        
        self.config = ConfigManager()
        self.saved_data = self.config.load()
        
        self.scanned_data = self.saved_data.get("scanned_data", {}) 
        self.all_subjects = self.saved_data.get("all_subjects", [])
        self.outlook_mode = tk.BooleanVar(value=False)

        self.setup_ui()

    def setup_ui(self):
        # --- STYLY ---
        style = ttk.Style()
        style.theme_use('clam') 
        
        style.configure("TFrame", background=COLOR_BG)
        style.configure("TLabelframe", background=COLOR_BG, foreground=COLOR_TEXT)
        style.configure("TLabelframe.Label", background=COLOR_BG, foreground=COLOR_ACCENT)
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Segoe UI", 10))
        style.configure("TButton", padding=6, font=("Segoe UI", 10), background="#444", foreground="white", borderwidth=0)
        style.map("TButton", background=[('active', '#555')])
        style.configure("TCombobox", fieldbackground=COLOR_ENTRY_BG, background="#444", foreground=COLOR_TEXT, arrowcolor="white")
        style.map("TCombobox", fieldbackground=[('readonly', COLOR_ENTRY_BG)], selectbackground=[('readonly', '#555')])
        style.configure("TCheckbutton", background=COLOR_BG, foreground=COLOR_TEXT, font=("Segoe UI", 10))

        # --- HLAVNÍ SCROLLOVACÍ PLÁTNO ---
        main_canvas = tk.Canvas(self.root, bg=COLOR_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=main_canvas.yview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(
                scrollregion=main_canvas.bbox("all")
            )
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar.set)

        main_canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        content_frame = ttk.Frame(scrollable_frame, padding="15")
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 1. PŘIHLAŠOVACÍ ÚDAJE
        lbl_frame_login = ttk.LabelFrame(content_frame, text="1. Přihlašovací údaje (UIS)", padding="10")
        lbl_frame_login.pack(fill=tk.X, pady=5)

        ttk.Label(lbl_frame_login, text="Login:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.entry_user = tk.Entry(lbl_frame_login, width=25, bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, insertbackground='white')
        self.entry_user.insert(0, self.saved_data.get("username", "")) 
        self.entry_user.grid(row=0, column=1, sticky=tk.W, padx=5)

        ttk.Label(lbl_frame_login, text="Heslo:").grid(row=0, column=2, sticky=tk.W, padx=5)
        self.entry_pass = tk.Entry(lbl_frame_login, width=25, show="*", bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, insertbackground='white')
        self.entry_pass.grid(row=0, column=3, sticky=tk.W, padx=5)

        # 2. AUTOMATICKÉ NAČTENÍ
        lbl_frame_scan = ttk.LabelFrame(content_frame, text="2. Automatické načtení (Doporučeno)", padding="10")
        lbl_frame_scan.pack(fill=tk.X, pady=5)
        
        lbl_scan_info = ttk.Label(lbl_frame_scan, text="Klikni pro načtení učitelů a předmětů + detekci tvé fakulty. Data se uloží pro příště.", wraplength=600)
        lbl_scan_info.pack(pady=(0, 5))
        
        self.btn_scan = tk.Button(lbl_frame_scan, text="🔄 Načíst data z UIS", bg=COLOR_BTN_SCAN, fg="white", font=("Segoe UI", 10, "bold"), command=self.start_scan)
        self.btn_scan.pack(fill=tk.X)

        # 3. VÝBĚR PŘEDMĚTU
        lbl_frame_creator = ttk.LabelFrame(content_frame, text="3. Vybrat předmět ke sledování", padding="10")
        lbl_frame_creator.pack(fill=tk.X, pady=5)

        self.frame_detected = tk.Frame(lbl_frame_creator, bg=COLOR_BG)
        self.frame_detected.grid(row=1, column=0, columnspan=3, sticky=tk.W, padx=5, pady=5)
        
        tk.Label(self.frame_detected, text="Fakulta/Obor:", font=("Segoe UI", 9, "bold"), bg=COLOR_BG, fg=COLOR_TEXT).pack(side=tk.LEFT)
        saved_study_info = self.saved_data.get("study_info", "--- (Načte se po přihlášení) ---")
        self.lbl_study_info = tk.Label(self.frame_detected, text=saved_study_info, font=("Segoe UI", 9), bg=COLOR_BG, fg=COLOR_INFO)
        self.lbl_study_info.pack(side=tk.LEFT, padx=5)

        ttk.Label(lbl_frame_creator, text="Učitel:").grid(row=2, column=0, sticky=tk.W, padx=5, pady=2)
        self.cb_teacher = ttk.Combobox(lbl_frame_creator, width=38)
        self.cb_teacher.grid(row=2, column=1, sticky=tk.W, padx=5, pady=2)
        self.cb_teacher.bind("<<ComboboxSelected>>", self.on_teacher_selected) 
        ttk.Label(lbl_frame_creator, text="(např. Jadrná)", font=("Segoe UI", 8), foreground="#888").grid(row=2, column=2, sticky=tk.W)

        ttk.Label(lbl_frame_creator, text="Předmět:").grid(row=3, column=0, sticky=tk.W, padx=5, pady=2)
        self.cb_subject = ttk.Combobox(lbl_frame_creator, width=38) 
        self.cb_subject.grid(row=3, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(lbl_frame_creator, text="(např. Teorie řízení)", font=("Segoe UI", 8), foreground="#888").grid(row=3, column=2, sticky=tk.W)

        if self.scanned_data:
            self.cb_teacher['values'] = sorted(list(self.scanned_data.keys()))
        if self.all_subjects:
            self.cb_subject['values'] = sorted(self.all_subjects)

        ttk.Label(lbl_frame_creator, text="Konkrétní datum:").grid(row=4, column=0, sticky=tk.W, padx=5, pady=2)
        self.entry_date = tk.Entry(lbl_frame_creator, width=15, bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, insertbackground='white')
        self.entry_date.grid(row=4, column=1, sticky=tk.W, padx=5, pady=2)
        ttk.Label(lbl_frame_creator, text="(např. 22.01 nebo prázdné)", font=("Segoe UI", 8), foreground="#888").grid(row=4, column=2, sticky=tk.W)

        btn_add = tk.Button(lbl_frame_creator, text="⬇️ PŘIDAT DO SEZNAMU", bg="#444", fg="white", font=("Segoe UI", 9, "bold"), command=self.add_target)
        btn_add.grid(row=5, column=0, columnspan=3, pady=10, sticky=tk.EW)

        # 4. SEZNAM TERMÍNŮ
        lbl_frame_targets = ttk.LabelFrame(content_frame, text="4. Seznam hlídaných termínů (Priorita shora dolů)", padding="10")
        lbl_frame_targets.pack(fill=tk.BOTH, expand=True, pady=5)
        
        container_list = tk.Frame(lbl_frame_targets, bg=COLOR_BG)
        container_list.pack(fill=tk.BOTH, expand=True)
        
        frame_list = tk.Frame(container_list, bg=COLOR_BG)
        frame_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar_list = tk.Scrollbar(frame_list)
        scrollbar_list.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.list_targets = tk.Listbox(frame_list, height=5, bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, selectbackground=COLOR_ACCENT, selectforeground="black", font=("Consolas", 10), yscrollcommand=scrollbar_list.set)
        self.list_targets.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_list.config(command=self.list_targets.yview)
        
        frame_btns = tk.Frame(container_list, bg=COLOR_BG)
        frame_btns.pack(side=tk.RIGHT, fill=tk.Y, padx=(5, 0))
        
        tk.Button(frame_btns, text="⬆️", bg="#444", fg="white", width=4, command=self.move_up).pack(pady=2)
        tk.Button(frame_btns, text="⬇️", bg="#444", fg="white", width=4, command=self.move_down).pack(pady=2)
        tk.Button(frame_btns, text="🗑️", bg="#8b0000", fg="white", width=4, command=self.delete_item).pack(pady=(10, 2))

        saved_targets_str = self.saved_data.get("targets", "")
        if saved_targets_str:
            for line in saved_targets_str.split("\n"):
                if line.strip() and not line.startswith("#"):
                    self.list_targets.insert(tk.END, line.strip())

        # 5. BLACKLIST
        lbl_frame_blacklist = ttk.LabelFrame(content_frame, text="5. Ignorované termíny (Blacklist)", padding="10")
        lbl_frame_blacklist.pack(fill=tk.X, pady=5)
        
        ttk.Label(lbl_frame_blacklist, text="Zde napiš co nechceš (odděl středníkem). Např: 24.01; 8:00; Novák").pack(anchor=tk.W)
        self.entry_blacklist = tk.Entry(lbl_frame_blacklist, bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, insertbackground='white')
        self.entry_blacklist.pack(fill=tk.X, pady=2)
        self.entry_blacklist.insert(0, self.saved_data.get("blacklist", ""))

        # 6. OVLÁDÁNÍ
        lbl_frame_control = ttk.LabelFrame(content_frame, text="6. Ovládání", padding="10")
        lbl_frame_control.pack(fill=tk.X, pady=5)

        self.chk_outlook = ttk.Checkbutton(lbl_frame_control, text="📧 Aktivovat Outlook Watcher (Čekání na email)", variable=self.outlook_mode, onvalue=True, offvalue=False)
        self.chk_outlook.pack(anchor=tk.W, pady=(0, 5))
        ttk.Label(lbl_frame_control, text="Pozor: E-maily mají zpoždění. Vhodné jen pro nové termíny.", font=("Segoe UI", 8), foreground="gray").pack(anchor=tk.W, pady=(0, 10))

        btn_frame = ttk.Frame(lbl_frame_control)
        btn_frame.pack(fill=tk.X)

        self.btn_start = tk.Button(btn_frame, text="🚀 SPUSTIT SNIPER", bg=COLOR_BTN_START, fg="white", font=("Segoe UI", 12, "bold"), command=self.start_sniper)
        self.btn_start.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.btn_dog = tk.Button(btn_frame, text="🐶 NASTAVIT HLÍDACÍHO PSA", bg=COLOR_BTN_DOG, fg="white", font=("Segoe UI", 12, "bold"), command=self.start_dog_mode)
        self.btn_dog.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        self.btn_stop = tk.Button(btn_frame, text="🛑 ZASTAVIT", bg=COLOR_BTN_STOP, fg="white", font=("Segoe UI", 12, "bold"), command=self.stop_sniper, state=tk.DISABLED)
        self.btn_stop.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # LOG
        lbl_frame_log = ttk.LabelFrame(content_frame, text="Log (Průběh)", padding="10")
        lbl_frame_log.pack(fill=tk.BOTH, expand=True, pady=5)

        self.txt_log = scrolledtext.ScrolledText(lbl_frame_log, height=8, state='normal', bg="#000000", fg="#00ff00", font=("Consolas", 9))
        self.txt_log.pack(fill=tk.BOTH, expand=True)

        btn_coffee = tk.Button(content_frame, text="☕ Líbi se ti aplikace? Podpoř autora na Buy Me a Coffee", bg=COLOR_ACCENT, fg="black", font=("Segoe UI", 10, "bold"), command=lambda: webbrowser.open(COFFEE_URL))
        btn_coffee.pack(fill=tk.X, pady=10)

    # --- UI METODY ---
    def log(self, msg):
        try:
            self.txt_log.insert(tk.END, f"{msg}\n")
            self.txt_log.see(tk.END)
        except: pass

    def save_config(self):
        targets = "\n".join(self.list_targets.get(0, tk.END))
        study_info_text = self.lbl_study_info.cget("text")
        data = {
            "username": self.entry_user.get(),
            "targets": targets,
            "blacklist": self.entry_blacklist.get(),
            "scanned_data": self.scanned_data,
            "all_subjects": self.all_subjects,
            "study_info": study_info_text
        }
        self.config.save(data)

    def on_teacher_selected(self, event):
        t = self.cb_teacher.get()
        if t in self.scanned_data:
            self.cb_subject['values'] = sorted(list(self.scanned_data[t]))
            if self.scanned_data[t]: self.cb_subject.current(0)
        else:
            self.cb_subject['values'] = sorted(self.all_subjects)

    def add_target(self):
        subj = self.cb_subject.get().strip()
        teach = self.cb_teacher.get().strip()
        date = self.entry_date.get().strip()
        
        if not subj:
            messagebox.showwarning("Chyba", "Musíš vybrat nebo napsat název předmětu!")
            return

        line = f"{subj};{date};{teach}"
        self.list_targets.insert(tk.END, line)
        
        self.cb_subject.set('')
        self.cb_teacher.set('')
        self.entry_date.delete(0, tk.END)
        self.save_config()

    def move_up(self):
        idx = self.list_targets.curselection()
        if not idx or idx[0] == 0: return
        text = self.list_targets.get(idx[0])
        self.list_targets.delete(idx[0])
        self.list_targets.insert(idx[0]-1, text)
        self.list_targets.selection_set(idx[0]-1)
        self.save_config()
    
    def move_down(self):
        idx = self.list_targets.curselection()
        if not idx or idx[0] == self.list_targets.size()-1: return
        text = self.list_targets.get(idx[0])
        self.list_targets.delete(idx[0])
        self.list_targets.insert(idx[0]+1, text)
        self.list_targets.selection_set(idx[0]+1)
        self.save_config()

    def delete_item(self):
        idx = self.list_targets.curselection()
        if idx: 
            self.list_targets.delete(idx[0])
            self.save_config()

    def get_targets(self):
        raw = self.list_targets.get(0, tk.END)
        targets = []
        for line in raw:
            line = line.strip()
            if not line: continue
            parts = line.split(";")
            if len(parts) >= 1:
                targets.append({"subject": parts[0].strip(), "date": parts[1].strip() if len(parts)>1 else "", "filter": parts[2].strip() if len(parts)>2 else "", "original_line": line})
        return targets
    
    def remove_target_from_gui(self, original_line):
        def _remove():
            try:
                items = self.list_targets.get(0, tk.END)
                if original_line in items:
                    idx = items.index(original_line)
                    self.list_targets.delete(idx)
                    self.save_config()
            except: pass
        self.root.after(0, _remove)

    def update_study_info_ui(self, info_text):
        def _update():
            self.lbl_study_info.config(text=info_text)
        self.root.after(0, _update)

    # --- SELENIUM & LOGIC ---
    def init_driver(self):
        options = Options()
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option('useAutomationExtension', False)
        try:
            driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
            driver.maximize_window()
            return driver
        except Exception as e:
            self.log(f"CHYBA DRIVERU: {e}")
            return None
    
    def detect_study_info(self, driver):
        """Funkce ze staré verze pro detekci fakulty a oboru."""
        try:
            try:
                titulek_elem = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.ID, "titulek"))
                )
                full_text = titulek_elem.text
            except:
                full_text = driver.find_element(By.TAG_NAME, "body").text

            match = re.search(r"Studium\s*[-–—]?\s*(.+?)(?:,|$|\sobdobí)", full_text, re.IGNORECASE)
            
            if match:
                study_part = match.group(1).strip()
                study_part = study_part.split('[')[0].split('(')[0].strip()
                study_part = re.sub(r'\s+', ' ', study_part)
                self.update_study_info_ui(study_part)
                self.root.after(0, self.save_config)
        except Exception:
            pass

    def login_process(self, driver, user, pwd):
        self.log("🔵 Přihlašuji do UIS...")
        driver.get(UIS_LOGIN_URL)
        time.sleep(2)
        try: driver.find_element(By.XPATH, "//a[contains(@href, 'lang=cz')]").click(); time.sleep(2)
        except: pass
        try: driver.find_element(By.XPATH, "//div[@data-sysid='email']").click()
        except: pass
        try:
            driver.find_element(By.ID, "credential_0").send_keys(user)
            driver.find_element(By.ID, "credential_1").send_keys(pwd)
            driver.find_element(By.ID, "credential_1").send_keys(Keys.RETURN)
            time.sleep(5)
            if len(driver.find_elements(By.ID, "credential_1")) > 0: return False
            return True
        except: return False

    def navigate_to_exams(self, driver):
        try:
            if "moje_studium" not in driver.current_url:
                try: driver.find_element(By.PARTIAL_LINK_TEXT, "Portál studenta").click(); time.sleep(2)
                except: driver.find_element(By.XPATH, "//span[contains(text(), 'Moje studium')]").click(); time.sleep(2)
            
            self.detect_study_info(driver) # Přidána detekce pro UI

            try: 
                driver.find_element(By.XPATH, "//span[@data-sysid='prihlasovani-zkousky']/..").click()
            except:
                driver.get("https://is.czu.cz/auth/student/terminy_seznam.pl?lang=cz")
            time.sleep(2)
            return True
        except: return False

    def run_sniper_process(self, user, pwd, targets, use_outlook):
        driver = self.init_driver()
        if not driver: return
        
        try:
            if not self.login_process(driver, user, pwd):
                driver.quit(); self.reset_ui(); return
            self.navigate_to_exams(driver)
            uis_handle = driver.current_window_handle
            
            # --- OUTLOOK SETUP (DUAL LOGIN) ---
            active_checking_mode = not use_outlook 
            
            if use_outlook:
                driver.switch_to.new_window('tab')
                self.log("📧 Otevírám Outlook v novém tabu...")
                driver.get(OUTLOOK_URL)
                outlook_handle = driver.current_window_handle
                self.log("⏳ Čekám na tvé přihlášení do Outlooku...")
                try: WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.XPATH, "//div[@role='tree']")))
                except: driver.quit(); return
                self.log("✅ Outlook připraven. Sleduji poštu...")

            blacklist_val = self.entry_blacklist.get()
            blacklist = [b.strip() for b in blacklist_val.split(";") if b.strip()]
            
            while self.is_running:
                check_uis = True
                
                # REŽIM: ČEKÁM NA EMAIL
                if use_outlook and not active_checking_mode:
                    driver.switch_to.window(outlook_handle)
                    found_mail = False
                    for t in targets:
                        subj = t["subject"]
                        # Hledáme mail o vypsání NEBO uvolnění
                        xpath = f"//div[@role='option' and contains(@aria-label, 'Unread') and (contains(@aria-label, 'Vypsání termínu') or contains(@aria-label, 'Uvolnění místa')) and contains(@aria-label, '{subj}')]"
                        if driver.find_elements(By.XPATH, xpath):
                            self.log(f"🚨 MAIL: {subj}! Přepínám do UIS!")
                            found_mail = True
                            break
                    
                    if found_mail:
                        active_checking_mode = True # Jakmile najdeme mail, přepneme na aktivní režim a už v něm zůstaneme
                        check_uis = True
                    else:
                        check_uis = False # Žádný mail -> čekáme
                        time.sleep(5)
                
                # REŽIM: AKTIVNÍ SKENOVÁNÍ UIS
                if check_uis:
                    if use_outlook: driver.switch_to.window(uis_handle)
                    
                    # Refresh UIS
                    driver.refresh()
                    try: WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.ID, "table_2")))
                    except: pass
                    
                    # 1. Zjistit, kde jsem přihlášen (table_1) pro kontrolu priorit
                    my_reg_subjects = []
                    try:
                        rows1 = driver.find_elements(By.XPATH, "//table[@id='table_1']//tbody/tr")
                        for r in rows1: my_reg_subjects.append(r.text)
                    except: pass
                    
                    # 2. Hledat v table_2 (volné)
                    current_targets = self.get_targets() 
                    target_action_done = False

                    for i, t in enumerate(current_targets):
                        subj = t["subject"]
                        date = t["date"]
                        filtr = t["filter"]
                        original_line = t["original_line"]
                        
                        xpath = f"//table[@id='table_2']//tr[contains(., '{subj}')]"
                        if date: xpath += f"[contains(., '{date}')]"
                        if filtr: xpath += f"[contains(., '{filtr}')]"
                        
                        rows = driver.find_elements(By.XPATH, xpath)
                        for row in rows:
                            if any(b in row.text for b in blacklist): continue
                            
                            # --- LOGIKA PRIORIT (SWAP) ---
                            # Pokud už mám tento předmět zapsaný, ale našel jsem ho znovu tady (což znamená, že jsem našel
                            # předmět, který je v mém seznamu 'targets' výše = vyšší priorita), tak se odhlásím z toho starého.
                            already_have_this_subject = any(subj in s for s in my_reg_subjects)
                            
                            if already_have_this_subject:
                                self.log(f"⚠️ Mám {subj} zapsaný, ale našel jsem lepší prioritu! Zkouším přehlásit...")
                                try:
                                    # Najdi řádek v table_1 pro tento předmět a klikni "Odhlásit ihned"
                                    unreg_xpath = f"//table[@id='table_1']//tr[contains(., '{subj}')]//a[contains(@href, 'odhlasit_ihned=1')]"
                                    driver.find_element(By.XPATH, unreg_xpath).click()
                                    try: driver.switch_to.alert.accept()
                                    except: pass
                                    time.sleep(1)
                                    driver.refresh() 
                                    # Znovu najít řádek v table_2
                                    rows = driver.find_elements(By.XPATH, xpath)
                                    row = rows[0] 
                                except Exception as e:
                                    self.log(f"❌ Chyba při přehlašování: {e}")
                                    continue

                            # --- ZÁPIS ---
                            try:
                                btn = row.find_element(By.XPATH, ".//a[contains(@href, 'prihlasit_ihned=1')] | .//span[@data-sysid='small-arrow-right-double']/..")
                                self.log(f"🔥 VOLNO: {subj}! Klikám...")
                                driver.execute_script("arguments[0].click();", btn)
                                try: driver.switch_to.alert.accept()
                                except: pass
                                self.log(f"🎉 ZAPSÁNO: {subj}")
                                
                                # Pokud jsme v Outlook módu, NEKONČÍME, vracíme se hlídat další maily/termíny
                                # Pokud v klasickém módu, odstraníme ze seznamu
                                if not use_outlook:
                                    self.remove_target_from_gui(original_line)
                                
                                target_action_done = True
                                break 
                            except: pass
                        if target_action_done: break # Jdeme na nový refresh
                    
                    if not use_outlook:
                        time.sleep(random.uniform(3, 8))

        except Exception as e: self.log(f"CHYBA: {e}")
        finally: 
            if driver: driver.quit()
            self.root.after(0, self.reset_ui)

    def start_sniper(self):
        self.is_running = True
        self.btn_start.config(state="disabled")
        self.btn_dog.config(state="disabled")
        self.btn_stop.config(state="normal")
        self.thread = threading.Thread(target=self.run_sniper_process, args=(self.entry_user.get(), self.entry_pass.get(), self.get_targets(), self.outlook_mode.get()))
        self.thread.daemon = True
        self.thread.start()
    
    def start_scan(self):
        self.btn_scan.config(state="disabled", text="⏳ Načítám...")
        self.thread = threading.Thread(target=self.scan_process, args=(self.entry_user.get(), self.entry_pass.get())).start()
    
    def scan_process(self, user, pwd):
        driver = self.init_driver()
        if not driver:
            self.root.after(0, lambda: self.btn_scan.config(state="normal", text="🔄 Načíst data z UIS"))
            return
        try:
            self.login_process(driver, user, pwd)
            self.navigate_to_exams(driver)
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "table_2")))
                rows = driver.find_elements(By.XPATH, "//table[@id='table_2']//tbody/tr")
                data_map = {}
                all_s = set()
                for row in rows:
                    cells = row.find_elements(By.TAG_NAME, "td")
                    if len(cells) > 9:
                        s = cells[4].text.strip()
                        t = cells[9].text.strip()
                        if s: 
                            all_s.add(s)
                            if t:
                                if t not in data_map: data_map[t] = set()
                                data_map[t].add(s)
                self.scanned_data = {k: sorted(list(v)) for k, v in data_map.items()}
                self.all_subjects = sorted(list(all_s))
                self.root.after(0, lambda: [self.save_config(), messagebox.showinfo("OK", "Data načtena"), self.update_comboboxes()])
            except: pass
        finally: 
            driver.quit()
            self.root.after(0, lambda: self.btn_scan.config(state="normal", text="🔄 Načíst data z UIS"))

    def update_comboboxes(self):
        self.cb_teacher['values'] = sorted(list(self.scanned_data.keys()))
        self.cb_subject['values'] = sorted(self.all_subjects)

    def start_dog_mode(self):
        self.is_running = True
        self.btn_dog.config(state="disabled")
        self.btn_start.config(state="disabled")
        self.btn_stop.config(state="normal")
        threading.Thread(target=self.run_dog, args=(self.entry_user.get(), self.entry_pass.get(), self.get_targets())).start()

    def run_dog(self, u, p, targets):
        driver = self.init_driver()
        try:
            self.login_process(driver, u, p)
            self.navigate_to_exams(driver)
            blacklist_val = self.entry_blacklist.get()
            blacklist = [b.strip() for b in blacklist_val.split(";") if b.strip()]
            
            for t in targets:
                if not self.is_running: break
                subj = t["subject"]; date = t["date"]; filtr = t["filter"]
                self.log(f"Hledám psa pro: {subj}")
                xpath = f"//table[@id='table_2']//tr[contains(., '{subj}')]"
                if date: xpath += f"[contains(., '{date}')]"
                if filtr: xpath += f"[contains(., '{filtr}')]"
                
                while self.is_running:
                    found_action = False
                    rows = driver.find_elements(By.XPATH, xpath)
                    for row in rows:
                        if any(b in row.text for b in blacklist): continue
                        try:
                            # Hledáme odkaz, který v sobě má psa
                            dog = row.find_element(By.XPATH, ".//a[.//span[@data-sysid='terminy-pes'] or .//use[contains(@href, 'glyph1561')]]")
                            self.log("🐶 Klikám na psa...")
                            driver.execute_script("arguments[0].click();", dog)
                            time.sleep(2)
                            driver.back()
                            driver.refresh()
                            try: WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "table_2")))
                            except: pass
                            found_action = True
                            self.log("✅ Pes nastaven.")
                            break
                        except: pass
                    if not found_action: break
            self.log("Hotovo.")
        finally: driver.quit(); self.root.after(0, self.reset_ui)

    def stop_sniper(self): self.is_running = False
    
    def reset_ui(self):
        self.is_running = False
        self.btn_start.config(state="normal")
        self.btn_stop.config(state="disabled")
        self.btn_dog.config(state="normal")
        self.log("--- ZASTAVENO ---")

# =============================================================================
# TŘÍDA: TC SNIPER (Moodle)
# =============================================================================
class TCSniperApp:
    def __init__(self, root):
        self.root = root
        self.root.title("TC Sniper - Moodle Dark")
        self.root.geometry("500x600")
        self.root.configure(bg=COLOR_BG)
        self.driver = None
        self.is_running = False
        self.config = ConfigManager()
        self.saved_data = self.config.load()

        # Styl pro Dark Mode TC Sniper
        style = ttk.Style()
        style.theme_use('clam') 
        style.configure("TFrame", background=COLOR_BG)
        style.configure("TLabelframe", background=COLOR_BG, foreground=COLOR_TEXT)
        style.configure("TLabelframe.Label", background=COLOR_BG, foreground=COLOR_ACCENT)
        style.configure("TLabel", background=COLOR_BG, foreground=COLOR_TEXT, font=("Segoe UI", 10))
        style.configure("TButton", padding=6, font=("Segoe UI", 10), background="#444", foreground="white", borderwidth=0)
        style.map("TButton", background=[('active', '#555')])
        style.configure("TCheckbutton", background=COLOR_BG, foreground=COLOR_TEXT, font=("Segoe UI", 10))
        
        lbl = ttk.LabelFrame(root, text="Nastavení", padding=10)
        lbl.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        tk.Label(lbl, text="URL Testu:", bg=COLOR_BG, fg=COLOR_TEXT).grid(row=0, column=0)
        self.e_url = tk.Entry(lbl, width=40, bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, insertbackground='white'); self.e_url.grid(row=0, column=1)
        self.e_url.insert(0, self.saved_data.get("tc_url", ""))

        tk.Label(lbl, text="Dny (např. 15,16):", bg=COLOR_BG, fg=COLOR_TEXT).grid(row=1, column=0)
        self.e_days = tk.Entry(lbl, bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, insertbackground='white'); self.e_days.grid(row=1, column=1)
        self.e_days.insert(0, self.saved_data.get("tc_days", "15"))
        
        tk.Label(lbl, text="Čas od (HH:MM):", bg=COLOR_BG, fg=COLOR_TEXT).grid(row=2, column=0)
        self.e_t1 = tk.Entry(lbl, bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, insertbackground='white'); self.e_t1.grid(row=2, column=1); self.e_t1.insert(0, "18:00")
        
        tk.Label(lbl, text="Čas do (HH:MM):", bg=COLOR_BG, fg=COLOR_TEXT).grid(row=3, column=0)
        self.e_t2 = tk.Entry(lbl, bg=COLOR_ENTRY_BG, fg=COLOR_TEXT, insertbackground='white'); self.e_t2.grid(row=3, column=1); self.e_t2.insert(0, "19:00")

        self.chk_book = tk.BooleanVar(value=True)
        tk.Checkbutton(lbl, text="Zarezervovat", variable=self.chk_book, bg=COLOR_BG, fg=COLOR_TEXT, selectcolor=COLOR_BG, activebackground=COLOR_BG, activeforeground=COLOR_TEXT).grid(row=4, columnspan=2)

        self.btn_run = tk.Button(root, text="START", bg=COLOR_BTN_START, fg="white", command=self.run)
        self.btn_run.pack(fill=tk.X, padx=10)
        self.btn_stop = tk.Button(root, text="STOP", bg=COLOR_BTN_STOP, fg="white", command=self.stop, state="disabled")
        self.btn_stop.pack(fill=tk.X, padx=10, pady=5)
        
        self.txt = scrolledtext.ScrolledText(root, height=8, bg="black", fg="#00ff00", font=("Consolas", 9))
        self.txt.pack(fill=tk.BOTH, padx=10)

    def log(self, m): self.txt.insert(tk.END, m+"\n"); self.txt.see(tk.END)
    
    def run(self):
        self.is_running = True
        self.btn_run.config(state="disabled")
        self.btn_stop.config(state="normal")
        # Save config
        self.config.save({"tc_url": self.e_url.get(), "tc_days": self.e_days.get()})
        threading.Thread(target=self.process).start()

    def stop(self): self.is_running = False

    def process(self):
        user = self.saved_data.get("username", "") 
        
        url = self.e_url.get()
        days = [d.strip() for d in self.e_days.get().split(",")]
        t1 = datetime.strptime(self.e_t1.get(), "%H:%M").time()
        t2 = datetime.strptime(self.e_t2.get(), "%H:%M").time()
        
        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=Options())
        try:
            self.log("Jdu na Moodle Login...")
            driver.get(MOODLE_LOGIN_URL)
            creds = self.config.load()
            if "username" in creds:
                driver.find_element(By.ID, "username").send_keys(creds["username"])
                self.log("❗ Prosím přihlas se ručně, pokud to neproběhlo.")
            
            time.sleep(5) # Wait for login
            
            while self.is_running:
                driver.get(url)
                try:
                    WebDriverWait(driver, 5).until(lambda d: len(d.find_elements(By.CSS_SELECTOR, "td.alert")) > 0)
                    for td in driver.find_elements(By.CSS_SELECTOR, "td.alert.alert-success"):
                        txt = td.text.strip()
                        for d in days:
                            if txt.startswith(str(d)):
                                self.log(f"Datum {d} je volné!")
                                td.click()
                                time.sleep(1)
                                # Check time
                                found_time = False
                                for a in driver.find_elements(By.TAG_NAME, "a"):
                                    if " - " in a.text:
                                        ct_str = a.text.split(" - ")[0].strip()
                                        try:
                                            ct = datetime.strptime(ct_str, "%H:%M").time()
                                            if t1 <= ct <= t2:
                                                self.log(f"Čas {ct_str} vyhovuje!")
                                                winsound.Beep(1000, 500)
                                                if self.chk_book.get():
                                                    a.click()
                                                    try: driver.switch_to.alert.accept()
                                                    except: pass
                                                    self.log("Hotovo!")
                                                    self.is_running = False
                                                found_time = True
                                                break
                                        except: pass
                                if found_time: break
                        if not self.is_running: break
                except: pass
                time.sleep(3)
        except Exception as e: self.log(f"Err: {e}")
        finally: driver.quit(); self.root.after(0, lambda: self.btn_run.config(state="normal"))

if __name__ == "__main__":
    root = tk.Tk()
    app = LauncherApp(root)
    root.mainloop()
