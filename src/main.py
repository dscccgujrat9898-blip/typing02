#!/usr/bin/env python3
# typing_trainer_for_build.py
# Modified for PyInstaller EXE packaging:
# - resource_path / get_base_dir to support frozen exe
# - writable user data directory for DB, replays, certificates, MainFolder
# - sound loader tries packaged resources then user data folder
# - minimal other compatibility fixes
#
# Requirements: Python 3.8+, pip install PyQt5 reportlab
# Build with PyInstaller (instructions below)

import sys
import os
import time
import json
import hashlib
import sqlite3
import random
import smtplib
from datetime import datetime, timezone
from email.message import EmailMessage
from functools import partial

from PyQt5.QtCore import Qt, QTimer, QThread, pyqtSignal, QSize, QUrl, QPropertyAnimation, QRect
from PyQt5.QtGui import QTextCharFormat, QColor, QFont, QTextCursor
from PyQt5.QtWidgets import (
    QApplication, QWidget, QMainWindow, QFileDialog, QMessageBox,
    QVBoxLayout, QHBoxLayout, QPushButton, QLabel, QListWidget,
    QTextEdit, QComboBox, QSpinBox, QDialog, QFormLayout,
    QLineEdit, QAction, QToolBar, QStatusBar, QGroupBox, QTabWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QSlider, QFrame, QCheckBox, QInputDialog
)
from PyQt5.QtMultimedia import QSoundEffect
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas

# ---------------------------
# Helper: resource path and user data dir
# ---------------------------
def get_base_dir():
    """
    Return base directory for packaged resources (read-only) and
    user data directory for writable files.
    - When running as PyInstaller frozen exe, sys._MEIPASS contains packaged files.
    - For writable data (DB, replays, certs), use a folder in user's profile (APPDATA or home).
    """
    # packaged base (read-only resources inside exe)
    if getattr(sys, "frozen", False):
        packaged_base = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        packaged_base = os.path.dirname(os.path.abspath(__file__))
    # user data base (writable)
    if os.name == "nt":
        appdata = os.getenv("APPDATA") or os.path.expanduser("~")
        user_base = os.path.join(appdata, "TypingTrainer")
    else:
        user_base = os.path.join(os.path.expanduser("~"), ".typing_trainer")
    return packaged_base, user_base

def resource_path_packaged(relative_path):
    """
    Return path to packaged resource (read-only). Use for resources bundled into exe.
    """
    packaged_base, _ = get_base_dir()
    return os.path.join(packaged_base, relative_path)

def resource_path_user(relative_path):
    """
    Return path inside user data folder (writable). Create directories as needed.
    """
    _, user_base = get_base_dir()
    full = os.path.join(user_base, relative_path)
    parent = os.path.dirname(full)
    if parent and not os.path.exists(parent):
        try:
            os.makedirs(parent, exist_ok=True)
        except Exception:
            pass
    return full

# ---------------------------
# App constants & institute info
# ---------------------------
APP_TITLE = "Digital Skill — Typing Trainer"

# Use user data folder for writable items
PACKAGED_BASE, USER_BASE = get_base_dir()
DEFAULT_MAIN_FOLDER = resource_path_user("MainFolder")
SUBFOLDERS = ["TypeA", "TypeB", "TypeC", "TypeD"]
DB_FILE = resource_path_user("typing_trainer.db")
REPLAY_DIR = resource_path_user("replays")
CERT_DIR = resource_path_user("certificates")
# For sounds: prefer packaged resources, but allow user folder fallback
SOUNDS_PACKAGED_DIR = resource_path_packaged("sounds")
SOUNDS_USER_DIR = resource_path_user("sounds")

INSTITUTE_BIG = "Digital Skill"
INSTITUTE_SMALL = "DS Computer and Coaching Classes"
INSTITUTE_ADDRESS = "Sai Ashirwad, Opp Modi Hospital, Near Manjulaben Hospital and KBC Children Hospital, Kadodara to Bardoli Road, Kadodara, Surat"
CONTACTS = ["9898710036", "7622953445"]
HONOUR_NAME = "Divya Kumari"
DEVELOPER_NAME = "Sanjeev Singh"

# ---------------------------
# Email (SMTP) configuration
# ---------------------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
EMAIL_FROM = "dscccgujrat9898@gmail.com"
# Provide password via environment variable APP_EMAIL_PASSWORD or set below (not recommended to hardcode)
EMAIL_PASSWORD = os.environ.get("APP_EMAIL_PASSWORD", "")

# ---------------------------
# Utilities & setup
# ---------------------------
def ensure_dirs():
    # Create writable user directories
    os.makedirs(DEFAULT_MAIN_FOLDER, exist_ok=True)
    for sf in SUBFOLDERS:
        os.makedirs(os.path.join(DEFAULT_MAIN_FOLDER, sf), exist_ok=True)
    os.makedirs(REPLAY_DIR, exist_ok=True)
    os.makedirs(CERT_DIR, exist_ok=True)
    # Ensure user sounds folder exists (so users can drop custom sounds)
    os.makedirs(SOUNDS_USER_DIR, exist_ok=True)

def file_sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        while True:
            chunk = f.read(8192)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()

def visible_whitespace(text):
    return text.replace(" ", "·").replace("\t", "→\t").replace("\n", "¶\n")

# ---------------------------
# Database (users + sessions) with recipient_email support
# ---------------------------
def init_db():
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        class TEXT,
        created_at TEXT
    )""")
    c.execute("""
    CREATE TABLE IF NOT EXISTS sessions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        user_id INTEGER,
        user_name TEXT,
        user_class TEXT,
        file_name TEXT,
        file_hash TEXT,
        folder_type TEXT,
        start_time TEXT,
        end_time TEXT,
        duration_seconds INTEGER,
        wpm REAL,
        accuracy REAL,
        errors INTEGER,
        replay_path TEXT,
        certificate_path TEXT,
        recipient_email TEXT
    )""")
    conn.commit()
    conn.close()

def save_user(name, cls):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("INSERT INTO users (name, class, created_at) VALUES (?, ?, ?)", (name, cls, datetime.now(timezone.utc).isoformat()))
    conn.commit()
    uid = c.lastrowid
    conn.close()
    return uid

def find_users_by_name(prefix):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("SELECT id, name, class FROM users WHERE name LIKE ? ORDER BY name LIMIT 50", (f"%{prefix}%",))
    rows = c.fetchall()
    conn.close()
    return rows

def save_session_db(user_id, user_name, user_class, file_name, file_hash, folder_type, start_time, end_time, duration, wpm, accuracy, errors, replay_path=None, certificate_path=None, recipient_email=None):
    conn = sqlite3.connect(DB_FILE)
    c = conn.cursor()
    c.execute("""
    INSERT INTO sessions (user_id, user_name, user_class, file_name, file_hash, folder_type, start_time, end_time, duration_seconds, wpm, accuracy, errors, replay_path, certificate_path, recipient_email)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (user_id, user_name, user_class, file_name, file_hash, folder_type, start_time, end_time, duration, wpm, accuracy, errors, replay_path, certificate_path, recipient_email))
    conn.commit()
    sid = c.lastrowid
    conn.close()
    return sid

# ---------------------------
# Paragraph generation helpers (Hindi & Gujarati)
# ---------------------------
def generate_paragraphs_from(base_sentences, count=15, target_words=320):
    paras = []
    for i in range(count):
        parts = []
        while sum(len(p.split()) for p in parts) < target_words:
            s = random.choice(base_sentences)
            parts.append(s)
        paras.append(" ".join(parts))
    return paras

HI_BASE = [
    "यह कहानी धैर्य और ईमानदारी के महत्व के बारे में है, जो छोटे‑छोटे कर्मों से बनती है।",
    "कंप्यूटर की दुनिया में समस्या सुलझाने की कला अभ्यास और ध्यान से आती है।",
    "एक विद्यार्थी ने रोज़ाना थोड़ी मेहनत से अपनी क्षमता बढ़ाई और अंततः सफलता पाई।",
    "मोरल कहानियाँ अक्सर सरल घटनाओं में गहरे सबक छुपाती हैं।",
    "टेक्नोलॉजी ने जीवन को आसान बनाया है, पर नैतिकता और जिम्मेदारी भी ज़रूरी है।",
    "एक प्रोग्रामर ने धैर्य से बग ढूँढा और टीम के लिए बड़ा समाधान निकाला।",
    "सहयोग और साझा ज्ञान से समुदाय मजबूत बनता है।",
    "छोटी‑छोटी आदतें समय के साथ बड़ी उपलब्धियाँ बनाती हैं।",
    "कंप्यूटर सीखना निरंतर अभ्यास और जिज्ञासा मांगता है।",
    "सही मार्गदर्शन और मेहनत से कोई भी कठिन काम आसान हो सकता है।",
    "कहानी में दिखाया गया कि ईमानदारी अंततः सम्मान और विश्वास दिलाती है।",
    "डिजिटल दुनिया में सुरक्षा और सतर्कता का महत्व बढ़ गया है।",
    "एक शिक्षक ने सरल उदाहरणों से जटिल विचारों को समझाया और छात्रों का आत्मविश्वास बढ़ा।",
    "नैतिकता और तकनीक का संतुलन समाज के लिए आवश्यक है।",
    "छात्रों को छोटे‑छोटे प्रयोगों से सीखने की प्रेरणा मिलती है।",
    "किसी भी समस्या का समाधान अक्सर धैर्य और बार‑बार प्रयास में छिपा होता है।",
    "कंप्यूटर प्रोग्रामिंग में त्रुटियों से सीखना ही असली विकास है।",
    "समय का सदुपयोग और अनुशासन सफलता की कुंजी हैं।",
    "एक छोटी मदद किसी के जीवन में बड़ा बदलाव ला सकती है।",
    "ज्ञान बांटने से वह बढ़ता है; साझा सीखना समुदाय को आगे बढ़ाता है।"
]

GU_BASE = [
    "આ વાર્તા ધીરજ અને સત્યની મહત્તા વિશે છે, જે નાના પગલાંઓથી બને છે.",
    "કમ્પ્યુટરની દુનિયામાં સમસ્યા ઉકેલવાની કળા અભ્યાસ અને ધ્યાનથી આવે છે.",
    "એક વિદ્યાર્થીએ રોજ થોડી મહેનતથી પોતાની ક્ષમતા વધારી અને સફળતા મેળવી.",
    "નૈતિક વાર્તાઓ ઘણીવાર સરળ ઘટનાઓમાં ઊંડા પાઠ છુપાવે છે.",
    "ટેકનોલોજીએ જીવન સરળ બનાવ્યું છે, પરંતુ જવાબદારી અને નૈતિકતા જરૂરી છે.",
    "એક પ્રોગ્રામરે ધીરજથી બગ શોધી અને ટીમ માટે મોટું ઉકેલ લાવ્યો.",
    "સહયોગ અને જ્ઞાન વહેંચવાથી સમુદાય મજબૂત બને છે.",
    "નાની‑નાની આદતો સમય સાથે મોટી સિદ્ધિઓ બનાવે છે.",
    "કમ્પ્યુટર શીખવું સતત અભ્યાસ અને જિજ્ઞાસા માંગે છે.",
    "સાચા માર્ગદર્શન અને મહેનતથી કોઈ પણ મુશ્કેલી સરળ બની શકે છે.",
    "વાર્તામાં બતાવવામાં આવ્યું કે સત્ય અને ઈમાનદારી અંતે માન અને વિશ્વાસ લાવે છે.",
    "ડિજિટલ દુનિયામાં સુરક્ષા અને સાવચેતીનું મહત્વ વધ્યું છે.",
    "શિક્ષકે સરળ ઉદાહરણોથી જટિલ વિચારો સમજાવ્યા અને વિદ્યાર્થીઓનો આત્મવિશ્વાસ વધ્યો.",
    "નૈતિકતા અને ટેકનોલોજીનું સંતુલન સમાજ માટે જરૂરી છે.",
    "વિદ્યાર્થીઓને નાના પ્રયોગોથી શીખવાની પ્રેરણા મળે છે.",
    "કોઈપણ સમસ્યાનું ઉકેલ ઘણીવાર ધીરજ અને પુનરાવર્તન માં છુપાયેલી હોય છે.",
    "પ્રોગ્રામિંગમાં ભૂલોથી શીખવું જ સાચું વિકાસ છે.",
    "સમયનો યોગ્ય ઉપયોગ અને શિસ્ત સફળતાની ચાવી છે.",
    "એક નાની મદદ કોઈના જીવનમાં મોટો ફેરફાર લાવી શકે છે.",
    "જ્ઞાન વહેંચવાથી તે વધે છે; શેર કરવું સમુદાયને આગળ વધારશે."
]

HINDI_PARAGRAPHS = generate_paragraphs_from(HI_BASE, count=15, target_words=320)
GUJARATI_PARAGRAPHS = generate_paragraphs_from(GU_BASE, count=15, target_words=320)

# ---------------------------
# Built-in drills & paragraphs (English)
# ---------------------------
LETTERS = list("abcdefghijklmnopqrstuvwxyz")
SIMPLE_WORDS = [
    "the", "and", "is", "in", "it", "you", "that", "he", "was", "for",
    "on", "are", "with", "as", "I", "his", "they", "be", "at", "one",
    "have", "this", "from", "or", "had", "by", "not", "word", "but", "what"
]

def generate_word_drill(length=20):
    return " ".join(random.choice(SIMPLE_WORDS) for _ in range(length))

def generate_letter_drill(length=60):
    return " ".join(random.choice(LETTERS) for _ in range(length))

def generate_number_drill(single_digit=True, count=60):
    if single_digit:
        return " ".join(str(random.randint(0,9)) for _ in range(count))
    else:
        return " ".join(str(random.randint(10, 999999)) for _ in range(count))

PARAGRAPHS_EN = [
    "This is a story about technology and human effort, where learning to type becomes a bridge between ideas and execution.",
    "In modern computing, clarity of thought is often matched by clarity of typing; practice builds both speed and precision.",
    "A developer's day can be long, but focused typing practice sharpens the mind and reduces friction when expressing algorithms.",
    "Moral lessons often hide in small routines: patience, repetition, and attention to detail lead to mastery.",
    "Stories about collaboration show that typing fast is useful, but typing accurately is what keeps teams in sync."
]
def generate_paragraph_en(length_words=320):
    parts = []
    while sum(len(p.split()) for p in parts) < length_words:
        parts.append(random.choice(PARAGRAPHS_EN))
    return " ".join(parts)

BUILTIN_TEXTS = {
    "English": [generate_paragraph_en(320) for _ in range(15)],
    "Hindi": HINDI_PARAGRAPHS,
    "Gujarati": GUJARATI_PARAGRAPHS,
    "Numbers": [generate_number_drill(single_digit=False, count=200) for _ in range(15)]
}

# ---------------------------
# Sound helpers (try packaged then user folder)
# ---------------------------
def load_loop_sound(filename, volume=0.35):
    # Try packaged path first
    packaged = os.path.join(SOUNDS_PACKAGED_DIR, filename)
    user_path = os.path.join(SOUNDS_USER_DIR, filename)
    path = None
    if os.path.exists(packaged):
        path = packaged
    elif os.path.exists(user_path):
        path = user_path
    else:
        # no sound available
        return None
    try:
        s = QSoundEffect()
        s.setSource(QUrl.fromLocalFile(path))
        s.setLoopCount(QSoundEffect.Infinite)
        s.setVolume(volume)
        return s
    except Exception:
        return None

def load_one_shot(filename, volume=0.6):
    packaged = os.path.join(SOUNDS_PACKAGED_DIR, filename)
    user_path = os.path.join(SOUNDS_USER_DIR, filename)
    path = None
    if os.path.exists(packaged):
        path = packaged
    elif os.path.exists(user_path):
        path = user_path
    else:
        return None
    try:
        s = QSoundEffect()
        s.setSource(QUrl.fromLocalFile(path))
        s.setLoopCount(1)
        s.setVolume(volume)
        return s
    except Exception:
        return None

# ---------------------------
# Certificate generator
# ---------------------------
def generate_certificate_pdf(candidate_name, candidate_class, honour_name, developer_name, institute_big, institute_small, address, contacts, exam_title, wpm, accuracy, out_path):
    try:
        c = canvas.Canvas(out_path, pagesize=landscape(A4))
        width, height = landscape(A4)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width/2, height - 80, institute_big)
        c.setFont("Helvetica", 18)
        c.drawCentredString(width/2, height - 110, institute_small)
        c.setFont("Helvetica", 12)
        c.drawCentredString(width/2, height - 140, address)
        c.setFont("Helvetica-Bold", 20)
        c.drawCentredString(width/2, height - 200, "Certificate of Typing Achievement")
        c.setFont("Helvetica", 16)
        c.drawString(80, height - 260, f"Awarded to: {candidate_name}")
        c.drawString(80, height - 290, f"Class: {candidate_class}")
        c.drawString(80, height - 320, f"Honour: {honour_name}")
        c.drawString(80, height - 350, f"Developer: {developer_name}")
        c.drawString(80, height - 380, f"Exam: {exam_title}")
        c.drawString(80, height - 410, f"WPM: {wpm:.2f}    Accuracy: {accuracy:.2f}%")
        c.setFont("Helvetica", 10)
        c.drawString(80, 60, f"Contact: {', '.join(contacts)}")
        c.drawString(width - 300, 60, f"Issued on: {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M UTC')}")
        c.drawString(width - 300, 40, "Signature: ____________________")
        c.showPage()
        c.save()
        return True
    except Exception as e:
        print("Certificate error:", e)
        return False

# ---------------------------
# Email helper (unchanged)
# ---------------------------
def send_email_with_attachment(to_email: str, subject: str, body: str, attachment_path: str, from_email: str = EMAIL_FROM, password: str = EMAIL_PASSWORD):
    if not password:
        raise RuntimeError("SMTP password not configured. Set APP_EMAIL_PASSWORD environment variable or edit EMAIL_PASSWORD in the script.")
    msg = EmailMessage()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = subject
    msg.set_content(body)
    # attach file
    with open(attachment_path, "rb") as f:
        data = f.read()
    maintype = "application"
    subtype = "pdf"
    filename = os.path.basename(attachment_path)
    msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=filename)
    # send
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.login(from_email, password)
        smtp.send_message(msg)

# ---------------------------
# Folder watcher (polling)
# ---------------------------
class FolderWatcher(QThread):
    folders_changed = pyqtSignal(dict)
    def __init__(self, main_folder, interval=2000):
        super().__init__()
        self.main_folder = main_folder
        self.interval = interval
        self._running = True
        self._snapshot = {}
    def run(self):
        while self._running:
            snapshot = {}
            for sf in SUBFOLDERS:
                folder_path = os.path.join(self.main_folder, sf)
                files = []
                if os.path.exists(folder_path):
                    for fname in sorted(os.listdir(folder_path)):
                        if fname.lower().endswith((".txt", ".md")):
                            full = os.path.join(folder_path, fname)
                            try:
                                mtime = os.path.getmtime(full)
                                files.append((fname, mtime))
                            except Exception:
                                pass
                snapshot[sf] = files
            if snapshot != self._snapshot:
                self._snapshot = snapshot
                self.folders_changed.emit(snapshot)
            self.msleep(self.interval)
    def stop(self):
        self._running = False
        self.wait()

# ---------------------------
# Replay Player
# ---------------------------
class ReplayPlayer(QDialog):
    def __init__(self, replay_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Keystroke Replay")
        self.setMinimumSize(700, 400)
        self.replay_path = replay_path
        layout = QVBoxLayout()
        self.display = QTextEdit()
        self.display.setReadOnly(True)
        self.display.setFontFamily("Courier")
        layout.addWidget(self.display)
        controls = QHBoxLayout()
        self.btn_play = QPushButton("Play")
        self.btn_play.clicked.connect(self.play)
        self.btn_stop = QPushButton("Stop")
        self.btn_stop.clicked.connect(self.stop)
        self.btn_stop.setEnabled(False)
        controls.addWidget(self.btn_play)
        controls.addWidget(self.btn_stop)
        layout.addLayout(controls)
        self.setLayout(layout)
        self._timer = QTimer()
        self._timer.timeout.connect(self._tick)
        self._events = []
        self._idx = 0
        self._running = False
        self.load_replay()
    def load_replay(self):
        try:
            with open(self.replay_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            self._events = data
            self.display.clear()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Cannot load replay: {e}")
    def play(self):
        if not self._events:
            return
        self._idx = 0
        self._running = True
        self.btn_play.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self._timer.start(50)
    def stop(self):
        self._timer.stop()
        self._running = False
        self.btn_play.setEnabled(True)
        self.btn_stop.setEnabled(False)
    def _tick(self):
        if not self._running:
            return
        if self._idx >= len(self._events):
            self.stop()
            return
        ev = self._events[self._idx]
        self.display.setPlainText(ev.get("text", ""))
        self._idx += 1

# ---------------------------
# Game Dialog (many drills)
# ---------------------------
class GameDialog(QDialog):
    def __init__(self, title, mode, source_text, max_seconds=60, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Game: {title}")
        self.setMinimumSize(900, 600)
        self.mode = mode
        self.source_text = source_text
        self.max_seconds = min(max_seconds, 15*60)
        self.remaining = self.max_seconds
        layout = QVBoxLayout()
        top = QHBoxLayout()
        self.lbl_timer = QLabel(f"Time: {self._format_time(self.remaining)}")
        self.lbl_timer.setStyleSheet("font-size:18px; font-weight:bold;")
        top.addWidget(self.lbl_timer)
        self.lbl_score = QLabel("Score: 0")
        self.lbl_score.setStyleSheet("font-size:18px; font-weight:bold;")
        top.addWidget(self.lbl_score)
        top.addStretch()
        layout.addLayout(top)

        # Animated background area
        self.anim_area = QFrame()
        self.anim_area.setMinimumHeight(120)
        self.anim_area.setStyleSheet("background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #87CEFA, stop:1 #ffffff); border-radius:8px;")
        anim_layout = QVBoxLayout()
        self.cloud_label = QLabel(self.anim_area)
        self.cloud_label.setText("☁️ ☁️")
        self.cloud_label.setStyleSheet("font-size:28px;")
        self.cloud_label.move(20, 20)
        anim_layout.addStretch()
        self.anim_area.setLayout(anim_layout)
        layout.addWidget(self.anim_area)

        # Source and input
        self.source_display = QTextEdit()
        self.source_display.setReadOnly(True)
        self.source_display.setPlainText(self.source_text)
        self.source_display.setFontFamily("Courier")
        self.source_display.setStyleSheet("font-size:16px;")
        layout.addWidget(QLabel("Source"))
        layout.addWidget(self.source_display)
        self.input_area = QTextEdit()
        self.input_area.setFontFamily("Courier")
        self.input_area.setStyleSheet("font-size:18px;")
        self.input_area.textChanged.connect(self.on_input)
        layout.addWidget(QLabel("Type here"))
        layout.addWidget(self.input_area)
        controls = QHBoxLayout()
        self.btn_start = QPushButton("Start")
        self.btn_start.clicked.connect(self.start)
        self.btn_stop = QPushButton("Stop")
        self.btn_stop.clicked.connect(self.stop)
        self.btn_stop.setEnabled(False)
        controls.addWidget(self.btn_start)
        controls.addWidget(self.btn_stop)
        layout.addLayout(controls)
        self.setLayout(layout)

        # Timer and animation
        self.timer = QTimer()
        self.timer.setInterval(1000)
        self.timer.timeout.connect(self._tick)
        self.score = 0
        self.started = False

        # bubble timer
        self.bubble_timer = QTimer()
        self.bubble_timer.setInterval(700)
        self.bubble_timer.timeout.connect(self.spawn_bubble)

        # cloud animation
        self.cloud_anim = QPropertyAnimation(self.cloud_label, b"geometry")
        self.cloud_anim.setDuration(12000)
        self.cloud_anim.setStartValue(QRect(-100, 10, 200, 50))
        self.cloud_anim.setEndValue(QRect(1000, 10, 200, 50))
        self.cloud_anim.setLoopCount(-1)

        # sound
        self.click_sound = load_one_shot("click.wav", volume=0.6)
        self.game_click = load_one_shot("game_click.wav", volume=0.6)
        self.game_bg = load_loop_sound("game_bg.wav", volume=0.18)

    def _format_time(self, s):
        m = s // 60
        sec = s % 60
        return f"{int(m):02d}:{int(sec):02d}"

    def start(self):
        # stop welcome sound when entering games
        try:
            parent = self.parent()
            if parent and getattr(parent, "welcome_sound", None):
                try:
                    parent.welcome_sound.stop()
                except Exception:
                    pass
        except Exception:
            pass

        self.remaining = self.max_seconds
        self.lbl_timer.setText(f"Time: {self._format_time(self.remaining)}")
        self.timer.start()
        self.bubble_timer.start()
        self.cloud_anim.start()
        # play game bg only if app sounds not muted
        try:
            parent = self.parent()
            if self.game_bg and not (parent and getattr(parent, "app_sounds_muted", False)):
                try: self.game_bg.play()
                except Exception: pass
        except Exception:
            pass
        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.input_area.setReadOnly(False)
        self.input_area.setFocus()
        self.started = True
        self.score = 0
        self.lbl_score.setText("Score: 0")
        self.input_area.clear()
        try:
            parent = self.parent()
            if self.game_click and not (parent and getattr(parent, "app_sounds_muted", False)):
                try: self.game_click.play()
                except Exception: pass
        except Exception:
            pass

    def stop(self):
        self.timer.stop()
        self.bubble_timer.stop()
        self.cloud_anim.stop()
        if self.game_bg:
            try: self.game_bg.stop()
            except Exception: pass
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.input_area.setReadOnly(True)
        self.started = False
        QMessageBox.information(self, "Game Over", f"Game ended. Score: {self.score}")

    def _tick(self):
        self.remaining -= 1
        self.lbl_timer.setText(f"Time: {self._format_time(self.remaining)}")
        if self.remaining <= 0:
            self.stop()

    def on_input(self):
        if not self.started:
            return
        typed = self.input_area.toPlainText()
        if self.mode == "letter_drill":
            src = "".join(self.source_text.split())
            typed_compact = "".join(typed.split())
            correct = sum(1 for i, ch in enumerate(typed_compact) if i < len(src) and ch == src[i])
            self.score = correct
        elif self.mode == "word_drill":
            src_words = self.source_text.split()
            typed_words = typed.split()
            correct = sum(1 for i, w in enumerate(typed_words) if i < len(src_words) and w == src_words[i])
            self.score = correct
        elif self.mode in ("number_single", "number_multi"):
            src = self.source_text.split()
            typed_words = typed.split()
            correct = sum(1 for i, w in enumerate(typed_words) if i < len(src) and w == src[i])
            self.score = correct
        elif self.mode == "paragraph":
            src = self.source_text
            correct = sum(1 for i, ch in enumerate(typed) if i < len(src) and ch == src[i])
            self.score = correct
        elif self.mode == "speed":
            src = self.source_text
            correct = sum(1 for i, ch in enumerate(typed) if i < len(src) and ch == src[i])
            self.score = int(correct / 5)
        else:
            self.score = 0
        self.lbl_score.setText(f"Score: {self.score}")
        try:
            parent = self.parent()
            if self.click_sound and self.score and self.score % 50 == 0 and not (parent and getattr(parent, "app_sounds_muted", False)):
                try:
                    self.click_sound.play()
                except Exception:
                    pass
        except Exception:
            pass

    def spawn_bubble(self):
        bubble = QLabel(self.anim_area)
        bubble.setText("●")
        size = random.randint(12, 28)
        bubble.setStyleSheet(f"color: rgba(255,255,255,0.9); font-size:{size}px;")
        start_x = random.randint(20, max(40, self.anim_area.width() - 40))
        start_y = self.anim_area.height() - 20
        bubble.setGeometry(start_x, start_y, 40, 40)
        bubble.show()
        anim = QPropertyAnimation(bubble, b"geometry")
        anim.setDuration(random.randint(3000, 7000))
        anim.setStartValue(QRect(start_x, start_y, 40, 40))
        anim.setEndValue(QRect(start_x + random.randint(-50,50), -40, 40, 40))
        anim.start()
        def cleanup():
            try:
                bubble.deleteLater()
            except Exception:
                pass
        anim.finished.connect(cleanup)

# ---------------------------
# Main Window (unchanged logic, but uses new paths)
# ---------------------------
class MainWindow(QMainWindow):
    def __init__(self, user_tuple):
        super().__init__()
        self.setWindowTitle(APP_TITLE)
        self.resize(1200, 820)
        ensure_dirs()
        init_db()
        # user info
        self.user_id, self.user_name, self.user_class = user_tuple
        # state
        self.main_folder = DEFAULT_MAIN_FOLDER
        self.settings = {"language": "English", "theme": "Light", "font_size": 14, "mute_app_sounds": False}
        self.session_running = False
        self.session_start_time = None
        self.keystroke_log = []
        self.current_file_path = None
        self.current_file_hash = None
        self.current_folder_type = None
        self.source_text = ""
        self.kiosk_mode = False
        # sounds
        self.welcome_sound = load_loop_sound("welcome.wav")
        self.typing_sound = load_loop_sound("typing_loop.wav")
        self.click_sound = load_one_shot("click.wav")
        # flags
        self.typing_muted = False
        self.app_sounds_muted = False
        # exam email controls
        self.require_student_email = False
        self.student_email_value = ""
        # UI
        self._create_toolbar()
        self._create_statusbar()
        self._create_central()
        # watcher
        self.watcher = FolderWatcher(self.main_folder)
        self.watcher.folders_changed.connect(self.on_folders_changed)
        self.watcher.start()
        # timer
        self.timer = QTimer()
        self.timer.setInterval(200)
        self.timer.timeout.connect(self.update_metrics)
        # last replay/cert
        self.last_replay_path = None
        self.last_certificate_path = None
        self.status.showMessage(f"Welcome {self.user_name} ({self.user_class}) — Ready")
        # play welcome sound if available
        if self.welcome_sound:
            try:
                self.welcome_sound.play()
            except Exception:
                pass

    def closeEvent(self, event):
        try:
            self.watcher.stop()
        except Exception:
            pass
        # stop sounds
        try:
            if self.welcome_sound:
                self.welcome_sound.stop()
            if self.typing_sound:
                self.typing_sound.stop()
        except Exception:
            pass
        event.accept()

    # ---------------------------
    # Settings dialog and apply helpers
    # ---------------------------
    def open_settings(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Settings")
        dlg.setMinimumSize(420, 220)
        layout = QFormLayout()
        lang = QComboBox()
        lang.addItems(["English", "Hindi", "Gujarati"])
        lang.setCurrentText(self.settings.get("language", "English"))
        theme = QComboBox()
        theme.addItems(["Light", "Dark"])
        theme.setCurrentText(self.settings.get("theme", "Light"))
        font_size = QSpinBox()
        font_size.setRange(10, 36)
        font_size.setValue(self.settings.get("font_size", 14))
        mute_checkbox = QCheckBox("Mute App Sounds (keep welcome sound)")
        mute_checkbox.setChecked(self.settings.get("mute_app_sounds", False))

        layout.addRow("Language", lang)
        layout.addRow("Theme", theme)
        layout.addRow("Font Size", font_size)
        layout.addRow(mute_checkbox)

        btn_save = QPushButton("Save")
        def save_and_close():
            self.settings["language"] = lang.currentText()
            self.settings["theme"] = theme.currentText()
            self.settings["font_size"] = font_size.value()
            self.settings["mute_app_sounds"] = mute_checkbox.isChecked()
            self.apply_settings()
            self.app_sounds_muted = self.settings["mute_app_sounds"]
            self.update_sound_state()
            dlg.accept()
        btn_save.clicked.connect(save_and_close)
        layout.addRow(btn_save)
        dlg.setLayout(layout)
        dlg.exec_()

    def apply_settings(self):
        size = self.settings.get("font_size", 14)
        font = QFont()
        font.setPointSize(size)
        try:
            self.source_display.setFont(font)
            self.input_area.setFont(font)
            self.preview.setFont(font)
        except Exception:
            pass
        if self.settings.get("theme") == "Dark":
            self.setStyleSheet("""
                QWidget { background: #2b2b2b; color: #e6e6e6; }
                QTextEdit { background: #1e1e1e; color: #e6e6e6; }
                QLineEdit { background: #1e1e1e; color: #e6e6e6; }
                QPushButton { background: #3a3a3a; color: #e6e6e6; }
            """)
        else:
            self.setStyleSheet("")
        try:
            self.status.showMessage("Settings applied")
        except Exception:
            pass

    def update_sound_state(self):
        try:
            if self.app_sounds_muted:
                if self.typing_sound:
                    try: self.typing_sound.stop()
                    except Exception: pass
            else:
                if self.typing_sound and self.session_running and not self.typing_muted:
                    try: self.typing_sound.play()
                    except Exception: pass
        except Exception:
            pass

    # ---------------------------
    # UI creation
    # ---------------------------
    def _create_toolbar(self):
        toolbar = QToolBar("Main")
        toolbar.setIconSize(QSize(16, 16))
        self.addToolBar(toolbar)
        act_open = QAction("Open MainFolder", self)
        act_open.triggered.connect(self.select_main_folder)
        toolbar.addAction(act_open)
        act_settings = QAction("Settings", self)
        act_settings.triggered.connect(self.open_settings)
        toolbar.addAction(act_settings)
        act_games = QAction("Games", self)
        act_games.triggered.connect(self.open_games_dialog)
        toolbar.addAction(act_games)
        act_kiosk = QAction("Toggle Secure Exam Mode", self)
        act_kiosk.triggered.connect(self.toggle_kiosk_mode)
        toolbar.addAction(act_kiosk)
        act_replay = QAction("Play Last Replay", self)
        act_replay.triggered.connect(self.play_last_replay)
        toolbar.addAction(act_replay)
        act_rules = QAction("Typing Rules & Notes", self)
        act_rules.triggered.connect(self.show_rules)
        toolbar.addAction(act_rules)

    def _create_statusbar(self):
        self.status = QStatusBar()
        self.setStatusBar(self.status)

    def _create_central(self):
        central = QWidget()
        main_layout = QHBoxLayout()
        left = QVBoxLayout()
        right = QVBoxLayout()

        # Left: library + time + user + exam email controls
        lib_box = QGroupBox("Text Library & Session")
        l_layout = QVBoxLayout()
        self.folder_select = QComboBox()
        self.folder_select.addItems(SUBFOLDERS + ["BuiltIn-English","BuiltIn-Hindi","BuiltIn-Gujarati","BuiltIn-Numbers"])
        self.folder_select.currentTextChanged.connect(self.refresh_file_lists)
        l_layout.addWidget(self.folder_select)
        self.file_list = QListWidget()
        self.file_list.itemClicked.connect(self.on_file_selected)
        l_layout.addWidget(self.file_list)
        time_layout = QHBoxLayout()
        time_layout.addWidget(QLabel("Set Time (minutes):"))
        self.time_spin = QSpinBox()
        self.time_spin.setRange(1, 60)
        self.time_spin.setValue(5)
        time_layout.addWidget(self.time_spin)
        self.btn_mute = QPushButton("Mute Typing Music")
        self.btn_mute.setCheckable(True)
        self.btn_mute.toggled.connect(self.toggle_typing_mute)
        time_layout.addWidget(self.btn_mute)
        l_layout.addLayout(time_layout)

        # Exam email controls
        exam_layout = QHBoxLayout()
        self.chk_require_email = QCheckBox("Require student email for exam")
        self.chk_require_email.stateChanged.connect(self.on_require_email_changed)
        exam_layout.addWidget(self.chk_require_email)
        l_layout.addLayout(exam_layout)
        email_layout = QHBoxLayout()
        email_layout.addWidget(QLabel("Student Email:"))
        self.le_student_email = QLineEdit()
        self.le_student_email.setPlaceholderText("student@example.com")
        email_layout.addWidget(self.le_student_email)
        l_layout.addLayout(email_layout)

        user_box = QLabel(f"User: {self.user_name}   Class: {self.user_class}")
        l_layout.addWidget(user_box)
        lib_box.setLayout(l_layout)
        left.addWidget(lib_box)

        # Preview
        preview_box = QGroupBox("Preview (visible whitespace)")
        pv_layout = QVBoxLayout()
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setFontFamily("Courier")
        self.preview.setMinimumHeight(160)
        pv_layout.addWidget(self.preview)
        preview_box.setLayout(pv_layout)
        left.addWidget(preview_box)

        # Right: typing area, slider, controls, reports
        tabs = QTabWidget()
        practice_tab = QWidget()
        p_layout = QVBoxLayout()
        self.source_display = QTextEdit(); self.source_display.setReadOnly(True); self.source_display.setFontFamily("Courier"); self.source_display.setMinimumHeight(180)
        p_layout.addWidget(QLabel("Source Text (exact)")); p_layout.addWidget(self.source_display)
        self.input_area = QTextEdit(); self.input_area.setFontFamily("Courier"); self.input_area.setPlaceholderText("Type here exactly as the source text.")
        self.input_area.textChanged.connect(self.on_input_changed)
        p_layout.addWidget(QLabel("Your Input")); p_layout.addWidget(self.input_area)
        # Font size slider
        slider_layout = QHBoxLayout()
        slider_layout.addWidget(QLabel("Font Size"))
        self.font_slider = QSlider(Qt.Horizontal)
        self.font_slider.setRange(10, 28)
        self.font_slider.setValue(self.settings.get("font_size", 14))
        self.font_slider.valueChanged.connect(self.on_font_slider_changed)
        slider_layout.addWidget(self.font_slider)
        self.font_label = QLabel(str(self.font_slider.value()))
        slider_layout.addWidget(self.font_label)
        p_layout.addLayout(slider_layout)
        controls = QHBoxLayout()
        self.btn_start = QPushButton("Start Session")
        self.btn_start.clicked.connect(self.start_session)
        self.btn_stop = QPushButton("Stop Session")
        self.btn_stop.clicked.connect(self.stop_session)
        self.btn_stop.setEnabled(False)
        controls.addWidget(self.btn_start); controls.addWidget(self.btn_stop)
        p_layout.addLayout(controls)
        metrics_layout = QHBoxLayout()
        self.lbl_wpm = QLabel("WPM: 0.00"); self.lbl_accuracy = QLabel("Accuracy: 0.00%"); self.lbl_errors = QLabel("Errors: 0")
        metrics_layout.addWidget(self.lbl_wpm); metrics_layout.addWidget(self.lbl_accuracy); metrics_layout.addWidget(self.lbl_errors)
        p_layout.addLayout(metrics_layout)
        replay_controls = QHBoxLayout()
        self.btn_export_replay = QPushButton("Export Replay (JSON)"); self.btn_export_replay.clicked.connect(self.export_replay); self.btn_export_replay.setEnabled(False)
        self.btn_play_replay = QPushButton("Play Last Replay"); self.btn_play_replay.clicked.connect(self.play_last_replay); self.btn_play_replay.setEnabled(False)
        self.btn_generate_cert = QPushButton("Generate Certificate (PDF)"); self.btn_generate_cert.clicked.connect(self.generate_cert_for_last); self.btn_generate_cert.setEnabled(False)
        replay_controls.addWidget(self.btn_export_replay); replay_controls.addWidget(self.btn_play_replay); replay_controls.addWidget(self.btn_generate_cert)
        p_layout.addLayout(replay_controls)
        practice_tab.setLayout(p_layout)
        tabs.addTab(practice_tab, "Practice")

        # Reports tab with certificate/email actions
        reports_tab = QWidget()
        r_layout = QVBoxLayout()
        self.table_sessions = QTableWidget(0, 10)
        self.table_sessions.setHorizontalHeaderLabels(["ID","User","Class","File","Folder","Start","End","WPM","Accuracy","Email"])
        self.table_sessions.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        r_layout.addWidget(self.table_sessions)
        reports_btns = QHBoxLayout()
        btn_load_reports = QPushButton("Load Reports"); btn_load_reports.clicked.connect(self.load_reports)
        btn_generate_cert_for_selected = QPushButton("Generate Certificate for Selected"); btn_generate_cert_for_selected.clicked.connect(self.generate_cert_for_selected)
        btn_email_cert_for_selected = QPushButton("Email Certificate for Selected"); btn_email_cert_for_selected.clicked.connect(self.email_cert_for_selected)
        reports_btns.addWidget(btn_load_reports); reports_btns.addWidget(btn_generate_cert_for_selected); reports_btns.addWidget(btn_email_cert_for_selected)
        r_layout.addLayout(reports_btns)
        reports_tab.setLayout(r_layout)
        tabs.addTab(reports_tab, "Reports / History")

        right.addWidget(tabs)
        main_layout.addLayout(left, 3); main_layout.addLayout(right, 7)
        central.setLayout(main_layout); self.setCentralWidget(central)
        self.refresh_file_lists()
        # apply initial font size
        self.on_font_slider_changed(self.font_slider.value())

    # ---------------------------
    # Exam email checkbox handler
    # ---------------------------
    def on_require_email_changed(self, state):
        self.require_student_email = (state == Qt.Checked)
        # keep field enabled so teacher can enter email; validation enforced on start
        self.le_student_email.setEnabled(True)

    # ---------------------------
    # Folder & file handling
    # ---------------------------
    def select_main_folder(self):
        dlg = QFileDialog(self); dlg.setFileMode(QFileDialog.Directory); dlg.setOption(QFileDialog.ShowDirsOnly, True)
        if dlg.exec_():
            selected = dlg.selectedFiles()[0]; self.main_folder = selected; self.refresh_file_lists()

    def refresh_file_lists(self):
        sel = self.folder_select.currentText(); self.file_list.clear()
        if sel.startswith("BuiltIn-"):
            lang = sel.split("-",1)[1]
            for idx in range(1,16):
                self.file_list.addItem(f"{lang} Paragraph {idx}")
        else:
            folder_path = os.path.join(self.main_folder, sel)
            if os.path.exists(folder_path):
                for fname in sorted(os.listdir(folder_path)):
                    if fname.lower().endswith((".txt", ".md")):
                        self.file_list.addItem(fname)

    def on_folders_changed(self, snapshot):
        current = self.folder_select.currentText()
        if current.startswith("BuiltIn-"): return
        files = snapshot.get(current, []); names = [f[0] for f in files]
        existing = [self.file_list.item(i).text() for i in range(self.file_list.count())]
        if existing != names:
            self.file_list.clear()
            for n in names: self.file_list.addItem(n)

    def on_file_selected(self, item):
        sel = self.folder_select.currentText()
        if sel.startswith("BuiltIn-"):
            lang = sel.split("-",1)[1]; idx = int(item.text().split()[-1]) - 1
            raw = BUILTIN_TEXTS.get(lang, [""]*15)[idx]
            self.current_file_path = None; self.current_file_hash = hashlib.sha256(raw.encode("utf-8")).hexdigest(); self.current_folder_type = sel; self.source_text = raw
            self.source_display.setPlainText(raw); self.preview.setPlainText(visible_whitespace(raw)); self.status.showMessage(f"Loaded built-in {sel} paragraph {idx+1}")
        else:
            sf = sel; folder_path = os.path.join(self.main_folder, sf); path = os.path.join(folder_path, item.text())
            try:
                with open(path, "r", encoding="utf-8") as f: raw = f.read()
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Cannot read file: {e}"); return
            self.current_file_path = path; self.current_file_hash = file_sha256(path); self.current_folder_type = sf; self.source_text = raw
            self.source_display.setPlainText(raw); self.preview.setPlainText(visible_whitespace(raw)); self.status.showMessage(f"Loaded {item.text()} ({sf})")

    # ---------------------------
    # Session control & metrics (exam email integrated)
    # ---------------------------
    def start_session(self):
        if not self.source_text:
            QMessageBox.information(self, "Select Text", "Please select a text (file or built-in) first.")
            return
        # if exam requires student email, validate
        if self.require_student_email:
            email = self.le_student_email.text().strip()
            if not email:
                QMessageBox.warning(self, "Email required", "Exam mode requires a student email. Enter student email before starting.")
                return
            # basic validation
            if "@" not in email or "." not in email:
                QMessageBox.warning(self, "Invalid email", "Enter a valid student email address.")
                return
            self.student_email_value = email

        # stop welcome sound when typing starts
        try:
            if self.welcome_sound:
                self.welcome_sound.stop()
        except Exception:
            pass

        minutes = self.time_spin.value()
        self.session_duration_seconds = minutes * 60
        self.input_area.clear(); self.input_area.setReadOnly(False); self.input_area.setFocus()
        self.session_running = True; self.session_start_time = time.time(); self.keystroke_log = []
        self.timer.start(); self.btn_start.setEnabled(False); self.btn_stop.setEnabled(True)
        self._session_start_dt = datetime.now(timezone.utc).isoformat()
        self.input_area.setContextMenuPolicy(Qt.NoContextMenu)
        # start typing sound only if not muted globally and not typing_muted
        if self.typing_sound and not self.typing_muted and not self.app_sounds_muted:
            try: self.typing_sound.play()
            except Exception: pass
        self.countdown = QTimer(); self.countdown.setInterval(1000); self.countdown.timeout.connect(self._countdown_tick); self.remaining_seconds = self.session_duration_seconds; self.countdown.start()
        self.status.showMessage(f"Session started for {minutes} minute(s)")

    def _countdown_tick(self):
        self.remaining_seconds -= 1
        if self.remaining_seconds <= 0:
            self.countdown.stop()
            self.stop_session()

    def stop_session(self):
        if not self.session_running:
            return
        self.session_running = False; self.timer.stop(); self.btn_start.setEnabled(True); self.btn_stop.setEnabled(False); self.input_area.setReadOnly(True)
        if self.typing_sound:
            try: self.typing_sound.stop()
            except Exception: pass
        end_time = time.time(); duration = int(end_time - self.session_start_time)
        wpm, accuracy, errors = self.compute_metrics(final=True)
        replay_path = None
        try:
            replay_name = f"replay_{int(time.time())}.json"; replay_path = os.path.join(REPLAY_DIR, replay_name)
            data = [{"t": t, "text": txt} for (t, txt) in self.keystroke_log]
            with open(replay_path, "w", encoding="utf-8") as f: json.dump(data, f, ensure_ascii=False, indent=2)
            self.last_replay_path = replay_path; self.btn_export_replay.setEnabled(True); self.btn_play_replay.setEnabled(True)
        except Exception:
            replay_path = None

        # generate certificate and optionally email if exam mode with student email
        cert_path = None
        try:
            cert_name = f"certificate_{int(time.time())}.pdf"
            cert_path = os.path.join(CERT_DIR, cert_name)
            generate_certificate_pdf(self.user_name, self.user_class, HONOUR_NAME, DEVELOPER_NAME, INSTITUTE_BIG, INSTITUTE_SMALL, INSTITUTE_ADDRESS, CONTACTS, os.path.basename(self.current_file_path) if self.current_file_path else "BuiltIn", wpm or 0.0, accuracy or 0.0, cert_path)
            self.last_certificate_path = cert_path
        except Exception:
            cert_path = None

        # save session with recipient_email if provided
        recipient = self.student_email_value if self.require_student_email else None
        save_session_db(self.user_id, self.user_name, self.user_class, os.path.basename(self.current_file_path) if self.current_file_path else "BuiltIn", self.current_file_hash, self.current_folder_type, self._session_start_dt, datetime.now(timezone.utc).isoformat(), duration, wpm, accuracy, errors, replay_path, cert_path, recipient_email=recipient)

        # if exam mode and student email present, send email with certificate (if generated)
        if self.require_student_email and recipient:
            if cert_path and os.path.exists(cert_path):
                try:
                    subject = f"Typing Exam Result - {self.user_name}"
                    body = f"Dear Student,\n\nAttached is your typing exam certificate.\n\nWPM: {wpm:.2f}\nAccuracy: {accuracy:.2f}%\n\nRegards,\n{INSTITUTE_BIG}"
                    send_email_with_attachment(recipient, subject, body, cert_path)
                except Exception as e:
                    QMessageBox.warning(self, "Email Error", f"Could not send email to student: {e}")

        self.btn_generate_cert.setEnabled(True)
        self.status.showMessage(f"Session stopped. WPM: {wpm:.2f}, Accuracy: {accuracy:.2f}%")
        QMessageBox.information(self, "Session Result", f"WPM: {wpm:.2f}\nAccuracy: {accuracy:.2f}%\nErrors: {errors}")

    def update_metrics(self):
        if not self.session_running:
            return
        wpm, accuracy, errors = self.compute_metrics()
        self.lbl_wpm.setText(f"WPM: {wpm:.2f}"); self.lbl_accuracy.setText(f"Accuracy: {accuracy:.2f}%"); self.lbl_errors.setText(f"Errors: {errors}")

    def compute_metrics(self, final=False):
        typed = self.input_area.toPlainText(); source = self.source_text
        total_typed = len(typed); correct = 0; errors = 0
        for i, ch in enumerate(typed):
            if i < len(source) and ch == source[i]:
                correct += 1
            else:
                errors += 1
        if final and len(typed) < len(source):
            errors += (len(source) - len(typed))
        elapsed = max(1, time.time() - self.session_start_time) if self.session_running else 1
        minutes = elapsed / 60.0
        net_wpm = (correct / 5.0) / minutes
        accuracy = (correct / total_typed * 100.0) if total_typed > 0 else 0.0
        net_wpm = max(0.0, net_wpm)
        return net_wpm, accuracy, errors

    # ---------------------------
    # Input handling & replay logging
    # ---------------------------
    def on_input_changed(self):
        if not self.session_running:
            return
        now = time.time(); text = self.input_area.toPlainText(); self.keystroke_log.append((now, text))
        self.highlight_input(text)

    def highlight_input(self, typed):
        cursor = self.input_area.textCursor(); pos = cursor.position()
        fmt_ok = QTextCharFormat(); fmt_ok.setForeground(QColor("black")); fmt_ok.setBackground(QColor("white"))
        fmt_err = QTextCharFormat(); fmt_err.setForeground(QColor("white")); fmt_err.setBackground(QColor("#d9534f"))
        self.input_area.blockSignals(True); self.input_area.selectAll(); self.input_area.setCurrentCharFormat(fmt_ok)
        source = self.source_text; doc = self.input_area.document()
        for i in range(len(typed)):
            if i < len(source) and typed[i] == source[i]:
                continue
            cursor = QTextCursor(doc); cursor.setPosition(i); cursor.movePosition(QTextCursor.Right, QTextCursor.KeepAnchor, 1); cursor.setCharFormat(fmt_err)
        cursor = self.input_area.textCursor(); cursor.setPosition(pos); self.input_area.setTextCursor(cursor); self.input_area.blockSignals(False)

    # ---------------------------
    # Replay export & play
    # ---------------------------
    def export_replay(self):
        if not self.last_replay_path:
            QMessageBox.information(self, "No Replay", "No replay available.")
            return
        dlg = QFileDialog(self); dlg.setAcceptMode(QFileDialog.AcceptSave); dlg.setDefaultSuffix("json"); dlg.setNameFilter("JSON files (*.json)"); dlg.setDirectory(REPLAY_DIR)
        if dlg.exec_():
            dest = dlg.selectedFiles()[0]
            try:
                with open(self.last_replay_path, "r", encoding="utf-8") as src: data = src.read()
                with open(dest, "w", encoding="utf-8") as dst: dst.write(data)
                QMessageBox.information(self, "Exported", f"Replay exported to {dest}")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Cannot export: {e}")

    def play_last_replay(self):
        if not self.last_replay_path:
            QMessageBox.information(self, "No Replay", "No replay available.")
            return
        dlg = ReplayPlayer(self.last_replay_path, self); dlg.exec_()

    # ---------------------------
    # Certificate generation & emailing from Reports tab
    # ---------------------------
    def load_reports(self):
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT id, user_name, user_class, file_name, folder_type, start_time, end_time, wpm, accuracy, recipient_email FROM sessions ORDER BY id DESC LIMIT 500")
        rows = c.fetchall(); conn.close()
        self.table_sessions.setRowCount(0)
        for r in rows:
            row_idx = self.table_sessions.rowCount(); self.table_sessions.insertRow(row_idx)
            for col_idx, val in enumerate(r):
                self.table_sessions.setItem(row_idx, col_idx, QTableWidgetItem(str(val) if val is not None else ""))
        self.status.showMessage(f"Loaded {len(rows)} sessions")

    def generate_cert_for_selected(self):
        sel = self.table_sessions.currentRow()
        if sel < 0:
            QMessageBox.information(self, "Select", "Select a session row first.")
            return
        sid_item = self.table_sessions.item(sel, 0)
        if not sid_item:
            QMessageBox.information(self, "Select", "Select a session row first.")
            return
        sid = int(sid_item.text())
        # fetch session details
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT user_name, user_class, file_name, wpm, accuracy FROM sessions WHERE id=?", (sid,))
        row = c.fetchone(); conn.close()
        if not row:
            QMessageBox.warning(self, "Error", "Session not found.")
            return
        user_name, user_class, file_name, wpm, accuracy = row
        cert_name = f"certificate_{sid}.pdf"; cert_path = os.path.join(CERT_DIR, cert_name)
        ok = generate_certificate_pdf(user_name, user_class, HONOUR_NAME, DEVELOPER_NAME, INSTITUTE_BIG, INSTITUTE_SMALL, INSTITUTE_ADDRESS, CONTACTS, file_name, wpm or 0.0, accuracy or 0.0, cert_path)
        if ok:
            QMessageBox.information(self, "Certificate Created", f"Certificate saved: {cert_path}")
            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("UPDATE sessions SET certificate_path=? WHERE id=?", (cert_path, sid)); conn.commit(); conn.close()
            except Exception:
                pass
        else:
            QMessageBox.warning(self, "Error", "Certificate generation failed.")

    def email_cert_for_selected(self):
        sel = self.table_sessions.currentRow()
        if sel < 0:
            QMessageBox.information(self, "Select", "Select a session row first.")
            return
        sid_item = self.table_sessions.item(sel, 0)
        if not sid_item:
            QMessageBox.information(self, "Select", "Select a session row first.")
            return
        sid = int(sid_item.text())
        conn = sqlite3.connect(DB_FILE); c = conn.cursor()
        c.execute("SELECT user_name, user_class, file_name, wpm, accuracy, certificate_path, recipient_email FROM sessions WHERE id=?", (sid,))
        row = c.fetchone(); conn.close()
        if not row:
            QMessageBox.warning(self, "Error", "Session not found.")
            return
        user_name, user_class, file_name, wpm, accuracy, cert_path, recipient_email = row
        # if no certificate, generate
        if not cert_path or not os.path.exists(cert_path):
            cert_name = f"certificate_{sid}.pdf"; cert_path = os.path.join(CERT_DIR, cert_name)
            ok = generate_certificate_pdf(user_name, user_class, HONOUR_NAME, DEVELOPER_NAME, INSTITUTE_BIG, INSTITUTE_SMALL, INSTITUTE_ADDRESS, CONTACTS, file_name, wpm or 0.0, accuracy or 0.0, cert_path)
            if ok:
                try:
                    conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("UPDATE sessions SET certificate_path=? WHERE id=?", (cert_path, sid)); conn.commit(); conn.close()
                except Exception:
                    pass
            else:
                QMessageBox.warning(self, "Error", "Certificate generation failed.")
                return
        # determine recipient email
        if not recipient_email:
            # ask user to input recipient email
            email, ok = QInputDialog.getText(self, "Recipient Email", "Enter recipient email to send certificate:")
            if not ok or not email.strip():
                QMessageBox.information(self, "Cancelled", "Email cancelled.")
                return
            recipient_email = email.strip()
            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("UPDATE sessions SET recipient_email=? WHERE id=?", (recipient_email, sid)); conn.commit(); conn.close()
            except Exception:
                pass
        # send email
        try:
            subject = f"Typing Certificate - {user_name}"
            body = f"Dear {user_name},\n\nAttached is your typing certificate for {file_name}.\nWPM: {wpm:.2f}\nAccuracy: {accuracy:.2f}%\n\nRegards,\n{INSTITUTE_BIG}"
            send_email_with_attachment(recipient_email, subject, body, cert_path)
            QMessageBox.information(self, "Email Sent", f"Certificate emailed to {recipient_email}")
        except Exception as e:
            QMessageBox.warning(self, "Email Error", f"Could not send email: {e}")

    # ---------------------------
    # NEW: Generate certificate for last session (fix for missing method)
    # ---------------------------
    def generate_cert_for_last(self):
        try:
            conn = sqlite3.connect(DB_FILE); c = conn.cursor()
            c.execute("SELECT id, user_name, user_class, file_name, wpm, accuracy FROM sessions WHERE user_id=? ORDER BY id DESC LIMIT 1", (self.user_id,))
            row = c.fetchone(); conn.close()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"DB error: {e}")
            return
        if not row:
            QMessageBox.information(self, "No Session", "No previous session found for this user.")
            return
        sid, user_name, user_class, file_name, wpm, accuracy = row
        cert_name = f"certificate_{sid}.pdf"; cert_path = os.path.join(CERT_DIR, cert_name)
        ok = generate_certificate_pdf(user_name, user_class, HONOUR_NAME, DEVELOPER_NAME, INSTITUTE_BIG, INSTITUTE_SMALL, INSTITUTE_ADDRESS, CONTACTS, file_name, wpm or 0.0, accuracy or 0.0, cert_path)
        if ok:
            QMessageBox.information(self, "Certificate Created", f"Certificate saved: {cert_path}")
            try:
                conn = sqlite3.connect(DB_FILE); c = conn.cursor(); c.execute("UPDATE sessions SET certificate_path=? WHERE id=?", (cert_path, sid)); conn.commit(); conn.close()
                self.last_certificate_path = cert_path
            except Exception:
                pass
        else:
            QMessageBox.warning(self, "Error", "Certificate generation failed.")

    # ---------------------------
    # Typing music mute toggle
    # ---------------------------
    def toggle_typing_mute(self, checked):
        self.typing_muted = checked
        if checked:
            self.btn_mute.setText("Unmute Typing Music")
            if self.typing_sound:
                try: self.typing_sound.stop()
                except Exception: pass
        else:
            self.btn_mute.setText("Mute Typing Music")
            if self.typing_sound and self.session_running and not self.app_sounds_muted:
                try: self.typing_sound.play()
                except Exception: pass

    # ---------------------------
    # Kiosk / Secure exam mode
    # ---------------------------
    def toggle_kiosk_mode(self):
        self.kiosk_mode = not self.kiosk_mode
        if self.kiosk_mode:
            self.showFullScreen(); self.status.showMessage("Secure Exam Mode: ON (best-effort)")
            for act in self.findChildren(QAction):
                if act.text() != "Toggle Secure Exam Mode": act.setEnabled(False)
            self.file_list.setEnabled(False); self.folder_select.setEnabled(False)
        else:
            self.showNormal(); self.status.showMessage("Secure Exam Mode: OFF")
            for act in self.findChildren(QAction): act.setEnabled(True)
            self.file_list.setEnabled(True); self.folder_select.setEnabled(True)

    def keyPressEvent(self, event):
        if self.kiosk_mode:
            if event.modifiers() & Qt.ControlModifier:
                if event.key() in (Qt.Key_C, Qt.Key_V, Qt.Key_X, Qt.Key_T, Qt.Key_N): return
            if event.key() in (Qt.Key_Print, Qt.Key_SysReq, Qt.Key_F11): return
        super().keyPressEvent(event)

    # ---------------------------
    # Reports & rules
    # ---------------------------
    def show_rules(self):
        rules = (
            "Typing Rules & Notes:\n\n"
            "1. Exact matching: Type exactly as the source text. Spaces, punctuation, and newlines matter.\n"
            "2. WPM calculation: Net WPM = (correct characters / 5) / minutes.\n"
            "3. Accuracy: (correct characters / total typed) * 100.\n"
            "4. Corrections: Backspace allowed; final untyped characters count as errors when session ends.\n"
            "5. Numbers mode: Use BuiltIn-Numbers or MainFolder/TypeD for numeric drills.\n"
            "6. Secure Exam Mode: App attempts to enforce kiosk-like behavior (fullscreen, block common shortcuts). OS-level actions (Alt+Tab, PrintScreen) may still work depending on OS.\n"
            "7. Replay: Keystrokes are saved as snapshots (timestamp + full text). Replay shows typing progression.\n"
            "8. Certificate: Generate PDF certificate after a session; it will include institute & user details automatically.\n"
            "9. Games: Many drills available: letter drills, word drills, number drills (single & multi-digit), paragraph and speed games.\n"
            "10. File handling: Place UTF-8 .txt files in MainFolder/TypeA..TypeD. App watches folder for changes.\n"
        )
        dlg = QDialog(self); dlg.setWindowTitle("Typing Rules & Notes"); dlg.setMinimumSize(700, 500)
        layout = QVBoxLayout(); te = QTextEdit(); te.setReadOnly(True); te.setPlainText(rules); layout.addWidget(te)
        btn = QPushButton("Close"); btn.clicked.connect(dlg.accept); layout.addWidget(btn); dlg.setLayout(layout); dlg.exec_()

    # ---------------------------
    # Games dialog launcher (15+ games)
    # ---------------------------
    def open_games_dialog(self):
        games = []
        for i in range(5):
            title = f"Letter Drill {i+1}"
            src = generate_letter_drill(length=60 + i*20)
            games.append((title, "letter_drill", src, 60))
        for i in range(4):
            title = f"Word Drill {i+1}"
            src = generate_word_drill(length=20 + i*10)
            games.append((title, "word_drill", src, 60))
        for i in range(3):
            title = f"Number Single Digit Drill {i+1}"
            src = generate_number_drill(single_digit=True, count=80 + i*20)
            games.append((title, "number_single", src, 60))
        for i in range(3):
            title = f"Number Multi Digit Drill {i+1}"
            src = generate_number_drill(single_digit=False, count=40 + i*20)
            games.append((title, "number_multi", src, 60))
        for i in range(2):
            title = f"Paragraph Drill {i+1}"
            src = generate_paragraph_en(length_words=150 + i*100) if i % 2 == 0 else random.choice(HINDI_PARAGRAPHS)
            games.append((title, "paragraph", src, 120))
        for i in range(2):
            title = f"Speed Game {i+1}"
            src = " ".join(generate_word_drill(length=80).split()[:80])
            games.append((title, "speed", src, 60))

        dlg = QDialog(self); dlg.setWindowTitle("Games"); dlg.setMinimumSize(420, 520)
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Select a game (15+ drills available)"))
        for title, mode, src, secs in games:
            btn = QPushButton(title)
            def make_open(t=title, m=mode, s=src, sec=secs):
                def _open():
                    g = GameDialog(t, m, s, max_seconds=sec, parent=self)
                    g.exec_()
                return _open
            btn.clicked.connect(make_open())
            layout.addWidget(btn)
        btn_close = QPushButton("Close"); btn_close.clicked.connect(dlg.accept); layout.addWidget(btn_close)
        dlg.setLayout(layout); dlg.exec_()

    # ---------------------------
    # Font slider handler
    # ---------------------------
    def on_font_slider_changed(self, val):
        self.font_label.setText(str(val))
        font = QFont()
        font.setPointSize(val)
        try:
            self.source_display.setFont(font)
            self.input_area.setFont(font)
            self.preview.setFont(font)
        except Exception:
            pass
        self.settings["font_size"] = val

# ---------------------------
# Welcome / User selection dialog
# ---------------------------
class WelcomeDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Welcome")
        self.setMinimumSize(700, 420)
        layout = QVBoxLayout()
        lbl_big = QLabel(INSTITUTE_BIG)
        lbl_big.setStyleSheet("font-size:36px; font-weight:bold; color:#2b6cb0;")
        lbl_small = QLabel(INSTITUTE_SMALL)
        lbl_small.setStyleSheet("font-size:16px; color:#333333;")
        layout.addWidget(lbl_big, alignment=Qt.AlignCenter)
        layout.addWidget(lbl_small, alignment=Qt.AlignCenter)
        addr = QLabel(INSTITUTE_ADDRESS)
        addr.setWordWrap(True)
        layout.addWidget(addr, alignment=Qt.AlignCenter)
        contacts = QLabel("Call / WhatsApp: " + "  |  ".join(CONTACTS) + "   📞  💬")
        layout.addWidget(contacts, alignment=Qt.AlignCenter)
        honour = QLabel(f"Honour: {HONOUR_NAME}    Developer: {DEVELOPER_NAME}")
        layout.addWidget(honour, alignment=Qt.AlignCenter)
        layout.addSpacing(10)
        form = QHBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search existing user by name (or leave blank to create new)")
        self.btn_search = QPushButton("Search")
        self.btn_search.clicked.connect(self.search_users)
        form.addWidget(self.search_input)
        form.addWidget(self.btn_search)
        layout.addLayout(form)
        self.results_list = QListWidget()
        self.results_list.itemDoubleClicked.connect(self.select_existing_user)
        layout.addWidget(self.results_list)
        new_box = QGroupBox("New User / Start")
        nb = QFormLayout()
        self.new_name = QLineEdit()
        self.new_class = QLineEdit()
        nb.addRow("Name", self.new_name)
        nb.addRow("Class", self.new_class)
        btn_create = QPushButton("Create & Start")
        btn_create.clicked.connect(self.create_and_start)
        nb.addRow(btn_create)
        new_box.setLayout(nb)
        layout.addWidget(new_box)
        self.setLayout(layout)
        # welcome sound (safe)
        self.welcome_sound = load_loop_sound("welcome.wav")
        if self.welcome_sound:
            try:
                self.welcome_sound.play()
            except Exception:
                pass
        self.selected_user = None

    def search_users(self):
        q = self.search_input.text().strip()
        self.results_list.clear()
        if not q:
            return
        rows = find_users_by_name(q)
        for r in rows:
            uid, name, cls = r
            self.results_list.addItem(f"{uid} | {name} | {cls}")

    def select_existing_user(self, item):
        try:
            text = item.text()
            parts = text.split("|")
            uid = int(parts[0].strip())
            name = parts[1].strip()
            cls = parts[2].strip()
            self.selected_user = (uid, name, cls)
            try:
                if self.welcome_sound:
                    self.welcome_sound.stop()
            except Exception:
                pass
            self.accept()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Selection error: {e}")

    def create_and_start(self):
        name = self.new_name.text().strip()
        cls = self.new_class.text().strip()
        if not name:
            QMessageBox.warning(self, "Input", "Enter name")
            return
        uid = save_user(name, cls)
        if uid is None:
            QMessageBox.warning(self, "Error", "Could not create user. Try again.")
            return
        self.selected_user = (uid, name, cls)
        try:
            if self.welcome_sound:
                self.welcome_sound.stop()
        except Exception:
            pass
        self.accept()

# ---------------------------
# Main entry
# ---------------------------
def main():
    ensure_dirs()
    init_db()
    app = QApplication(sys.argv)
    w = WelcomeDialog()
    if w.exec_() != QDialog.Accepted or not w.selected_user:
        QMessageBox.information(None, "Exit", "No user selected. Exiting.")
        return
    user_tuple = w.selected_user
    window = MainWindow(user_tuple)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
