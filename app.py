import sys
import os
import webbrowser
import tempfile
import json
import pandas as pd
import requests
import threading
import subprocess
from PIL import Image  # Yüksek kaliteli görsel işleme için
from PIL.ImageQt import ImageQt
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                             QLineEdit, QPushButton, QTextEdit, QFileDialog,
                             QMessageBox, QGroupBox, QFormLayout, QSplitter, QLabel, QDialog,
                             QFrame, QComboBox, QScrollArea)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QIcon, QClipboard, QPixmap, QDragEnterEvent, QDropEvent

VERSION = "v5.2"

# PyInstaller kaynak yolu çözücü
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# Harf sütununu indekse çeviren yardımcı fonksiyon
def col_to_index(col):
    col = col.upper().strip()
    if not col: return 0
    if col.isdigit(): return int(col) - 1
    index = 0
    for char in col:
        index = index * 26 + (ord(char) - ord('A') + 1)
    return index - 1

# --- GÜNCELLEME İŞLEMCİSİ (GITHUB ENTEGRASYONU) ---
class UpdateThread(QThread):
    update_found = pyqtSignal(str, str) # Yeni Versiyon, İndirme Linki
    def run(self):
        try:
            # GitHub'daki son sürümü kontrol et (Release üzerinden)
            repo_url = "https://api.github.com/repos/habibdogan/GAS-Sinav-Tasarimcisi/releases/latest"
            response = requests.get(repo_url, timeout=5)
            if response.status_code == 200:
                data = response.json()
                latest_v = data.get("tag_name", VERSION)
                html_url = data.get("html_url", "")
                if latest_v != VERSION:
                    self.update_found.emit(latest_v, html_url)
        except: pass

# --- SÜRÜKLE BIRAK ALANI ---
class DropArea(QFrame):
    file_dropped = pyqtSignal(str)
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setFrameStyle(QFrame.Shape.StyledPanel | QFrame.Shadow.Sunken)
        self.setMinimumHeight(100)
        self.setStyleSheet("""
            QFrame {
                border: 2px dashed #aaa;
                border-radius: 10px;
                background-color: rgba(0,0,0,0.03);
            }
        """)
        layout = QVBoxLayout(self)
        self.label = QLabel("Excel veya CSV Dosyasını Buraya Sürükleyin\n(Başlıkları Otomatik Çekmek İçin)")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls(): event.accept()
        else: event.ignore()

    def dropEvent(self, event: QDropEvent):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.endswith(('.xlsx', '.xls', '.csv')):
                self.file_dropped.emit(file_path)

# --- YARDIM PENCERESİ (OPTIMIZE EDILDI - AKILLI SIĞDIRMA) ---
class TutorialDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Nasıl Kullanılır?")
        self.resize(1000, 750)
        self.setModal(False)
        self.current_step = 1
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout(self)
        
        self.info_label = QLabel(f"Adım {self.current_step}")
        self.info_label.setStyleSheet("font-size: 15px; font-weight: bold; color: #2c3e50; margin-bottom: 5px;")
        self.info_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.info_label)

        # Görsel Alanı - Pencereye Sığan ve Net
        self.img_container = QWidget()
        self.img_container_layout = QVBoxLayout(self.img_container)
        self.img_label = QLabel("Görsel Yükleniyor...")
        self.img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.img_label.setStyleSheet("border: 1px solid #eee; background-color: #fff;")
        self.img_container_layout.addWidget(self.img_label)
        layout.addWidget(self.img_container)
        
        btns_layout = QHBoxLayout()
        self.btn_prev = QPushButton("⬅️ Geri")
        self.btn_open_ext = QPushButton("🔍 Orijinal Boyutta Aç")
        self.btn_next = QPushButton("İleri ➡️")
        
        # Buton stilleri
        self.btn_open_ext.setStyleSheet("""
            QPushButton { background-color: #3498db; color: white; font-weight: bold; border-radius: 5px; padding: 5px 15px; }
            QPushButton:hover { background-color: #2980b9; }
        """)
        for b in [self.btn_prev, self.btn_next]:
            b.setFixedHeight(40)
            b.setFixedWidth(120)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
        
        self.btn_open_ext.setFixedHeight(40)
        self.btn_open_ext.setCursor(Qt.CursorShape.PointingHandCursor)

        self.btn_prev.clicked.connect(self.prev)
        self.btn_next.clicked.connect(self.next)
        self.btn_open_ext.clicked.connect(self.open_external)

        btns_layout.addWidget(self.btn_prev)
        btns_layout.addStretch()
        btns_layout.addWidget(self.btn_open_ext)
        btns_layout.addStretch()
        btns_layout.addWidget(self.btn_next)
        layout.addLayout(btns_layout)
        
        self.update_step()

    def update_step(self):
        path = resource_path(f"step{self.current_step}.png")
        self.info_label.setText(f"Adım {self.current_step} / 5")
        
        if os.path.exists(path):
            try:
                with Image.open(path) as img:
                    # Akıllı Sığdırma: Pencere boyutuna göre yüksek kaliteli (LANCZOS) küçültme
                    # 980x600 makul bir sığdırma alanıdır
                    img.thumbnail((980, 600), Image.Resampling.LANCZOS)
                    
                    qimage = ImageQt(img)
                    pixmap = QPixmap.fromImage(qimage)
                    self.img_label.setPixmap(pixmap)
                    # Sabit boyutu kaldırıyoruz ki sığdırma aktif olsun
                    self.img_label.setFixedSize(pixmap.size()) 
            except Exception as e:
                self.img_label.setText(f"Görsel yükleme hatası: {str(e)}")
        else:
            self.img_label.setText(f"Adım {self.current_step}\n(Görsel bulunamadı)")
        
        self.btn_prev.setEnabled(self.current_step > 1)
        self.btn_next.setText("İleri ➡️" if self.current_step < 5 else "Anladım ✅")

    def open_external(self):
        path = resource_path(f"step{self.current_step}.png")
        if os.path.exists(path):
            if sys.platform == 'win32': os.startfile(path)
            elif sys.platform == 'darwin': subprocess.call(['open', path])
            else: subprocess.call(['xdg-open', path])

    def next(self):
        if self.current_step < 5: self.current_step += 1; self.update_step()
        else: self.hide()

    def prev(self):
        if self.current_step > 1: self.current_step -= 1; self.update_step()

# --- ANA UYGULAMA ---
class AppGenerator(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings_file = "settings.json"
        self.is_dark_theme = False
        self.headers = []
        self.tutorial_window = None 
        self.initUI()
        self.load_settings()
        self.apply_theme()
        self.check_updates()

    def check_updates(self):
        self.update_thread = UpdateThread()
        self.update_thread.update_found.connect(self.show_update_notification)
        self.update_thread.start()

    def show_update_notification(self, version, url):
        msg = QMessageBox(self)
        msg.setWindowTitle("Yeni Güncelleme Mevcut!")
        msg.setText(f"Yeni bir sürüm bulundu: {version}\n\nŞu anki sürümünüz: {VERSION}\n\nYeni özellikleri kullanmak için güncellemeyi indirmenizi öneririz.")
        msg.setIcon(QMessageBox.Icon.Information)
        btn_download = msg.addButton("Hemen İndir", QMessageBox.ButtonRole.AcceptRole)
        msg.addButton("Daha Sonra", QMessageBox.ButtonRole.RejectRole)
        msg.exec()
        if msg.clickedButton() == btn_download:
            webbrowser.open(url)

    def initUI(self):
        self.setWindowTitle(f"GAS Sınav Sistemi Tasarımcısı {VERSION}")
        self.setMinimumSize(1200, 850)
        
        icon_path = resource_path("app_icon.ico")
        if os.path.exists(icon_path): self.setWindowIcon(QIcon(icon_path))

        self.main_widget = QWidget()
        self.setCentralWidget(self.main_widget)
        self.main_layout = QVBoxLayout(self.main_widget)
        self.main_layout.setContentsMargins(15, 15, 15, 15)

        # --- Üst Bar ---
        top_bar = QHBoxLayout()
        self.header_label = QLabel("MAKÜ | GAS Sınav Tasarımcısı Pro")
        self.header_label.setStyleSheet("font-size: 22px; font-weight: bold; color: #2c3e50;")
        
        top_buttons = QHBoxLayout()
        self.btn_how_to = QPushButton("❓ Nasıl Kullanılır?")
        self.btn_theme = QPushButton("🌓 Tema Değiştir")
        
        self.btn_go_sheets = QPushButton("📊 Google E-Tablolara Git")
        self.btn_go_sheets.setStyleSheet("""
            QPushButton { background-color: #0F9D58; color: white; padding: 8px 15px; font-weight: bold; border-radius: 6px; }
            QPushButton:hover { background-color: #0b8043; }
        """)
        self.btn_go_sheets.clicked.connect(lambda: webbrowser.open("https://sheets.google.com"))
        
        self.btn_how_to.clicked.connect(self.show_tutorial)
        self.btn_theme.clicked.connect(self.toggle_theme)

        top_buttons.addWidget(self.btn_how_to)
        top_buttons.addWidget(self.btn_theme)
        top_buttons.addWidget(self.btn_go_sheets)
        
        top_bar.addWidget(self.header_label)
        top_bar.addStretch()
        top_bar.addLayout(top_buttons)
        self.main_layout.addLayout(top_bar)

        # --- Sürükle Bırak ---
        self.drop_area = DropArea()
        self.drop_area.file_dropped.connect(self.load_excel_headers)
        self.main_layout.addWidget(self.drop_area)

        # --- Ana İçerik ---
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- Sol Panel ---
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 10, 0)

        self.input_group = QGroupBox("Sistem Yapılandırması")
        self.form_layout = QFormLayout()
        self.form_layout.setSpacing(12)

        self.exam_title_input = QLineEdit("Sınav Başlığı")
        self.sheet_name_input = QLineEdit("Sayfa1")
        self.start_row_input = QLineEdit("2")
        
        self.tc_combo = QComboBox(); self.tc_combo.setEditable(True)
        self.no_combo = QComboBox(); self.no_combo.setEditable(True)
        self.name_combo = QComboBox(); self.name_combo.setEditable(True)
        self.grade_combo = QComboBox(); self.grade_combo.setEditable(True)
        
        self.error_msg_input = QLineEdit("Kayıt bulunamadı.")
        self.logo_url_input = QLineEdit("https://gubyo.mehmetakif.edu.tr/favicon.ico")

        self.form_layout.addRow("Sınav Başlığı:", self.exam_title_input)
        self.form_layout.addRow("Sayfa İsmi:", self.sheet_name_input)
        self.form_layout.addRow("Veri Başlama Satırı:", self.start_row_input)
        self.form_layout.addRow("TC Sütunu (Harf/Başlık):", self.tc_combo)
        self.form_layout.addRow("Öğrenci No Sütunu:", self.no_combo)
        self.form_layout.addRow("Ad Soyad Sütunu:", self.name_combo)
        self.form_layout.addRow("Sınav Notu Sütunu:", self.grade_combo)
        self.form_layout.addRow("Hata Mesajı:", self.error_msg_input)
        self.form_layout.addRow("Özel Logo URL:", self.logo_url_input)

        self.input_group.setLayout(self.form_layout)
        left_layout.addWidget(self.input_group)

        # Butonlar
        self.btn_generate = QPushButton("🚀 Kodları Üret")
        self.btn_preview = QPushButton("👁️ Arayüzü Önizle")
        self.btn_save = QPushButton("💾 Dosya Olarak Kaydet")
        
        for btn in [self.btn_generate, self.btn_preview, self.btn_save]:
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setFixedHeight(45)
            btn.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
            left_layout.addWidget(btn)

        self.btn_generate.clicked.connect(self.generate_code)
        self.btn_preview.clicked.connect(self.preview_html)
        self.btn_save.clicked.connect(self.save_files)
        
        left_layout.addStretch()

        # --- Sağ Panel ---
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        splitter_v = QSplitter(Qt.Orientation.Vertical)
        
        self.gs_output = QTextEdit(); self.gs_output.setReadOnly(True)
        self.html_output = QTextEdit(); self.html_output.setReadOnly(True)
        
        splitter_v.addWidget(self.create_code_box("Kod.gs (Google Script)", self.gs_output))
        splitter_v.addWidget(self.create_code_box("Index.html (Arayüz)", self.html_output))
        right_layout.addWidget(splitter_v)

        self.main_splitter.addWidget(left_panel); self.main_splitter.addWidget(right_panel)
        self.main_splitter.setSizes([450, 750])
        self.main_layout.addWidget(self.main_splitter)

    def show_tutorial(self):
        if not self.tutorial_window:
            self.tutorial_window = TutorialDialog(self)
        self.tutorial_window.show()
        self.tutorial_window.raise_()
        self.tutorial_window.activateWindow()

    def create_code_box(self, title, widget):
        cont = QWidget(); lay = QVBoxLayout(cont)
        head = QHBoxLayout(); head.addWidget(QLabel(title))
        btn = QPushButton("📋 Kopyala"); btn.clicked.connect(lambda ch, w=widget: self.copy_to(w.toPlainText()))
        head.addStretch(); head.addWidget(btn)
        lay.addLayout(head); lay.addWidget(widget)
        return cont

    def copy_to(self, text):
        if text: 
            QApplication.clipboard().setText(text)
            QMessageBox.information(self, "Başarılı", "Kod başarıyla kopyalandı!")

    def load_excel_headers(self, path):
        try:
            df = pd.read_csv(path, nrows=1) if path.endswith('.csv') else pd.read_excel(path, nrows=1)
            self.headers = df.columns.tolist()
            for c in [self.tc_combo, self.no_combo, self.name_combo, self.grade_combo]:
                c.clear(); c.addItems(self.headers)
            QMessageBox.information(self, "Bilgi", "Sütun başlıkları Excel'den çekildi.")
        except Exception as e: QMessageBox.critical(self, "Hata", str(e))

    def get_column_index(self, combo):
        text = combo.currentText().strip()
        if text in self.headers: return self.headers.index(text)
        return col_to_index(text)

    def generate_code(self):
        title = self.exam_title_input.text()
        sheet = self.sheet_name_input.text()
        row = self.start_row_input.text()
        logo = self.logo_url_input.text()
        
        tc_idx = self.get_column_index(self.tc_combo)
        no_idx = self.get_column_index(self.no_combo)
        nm_idx = self.get_column_index(self.name_combo)
        gr_idx = self.get_column_index(self.grade_combo)

        gs_code = f"""// GAS Sınav Tasarımcısı {VERSION} Pro
function doGet() {{
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('{title}')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}}

function searchStudent(tcNo, ogrNo) {{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("{sheet}");
  if (!sheet) return {{ success: false, message: "Sayfa bulunamadı." }};

  var data = sheet.getDataRange().getValues();
  for (var i = {row} - 1; i < data.length; i++) {{
    if (String(data[i][{tc_idx}]).trim() === String(tcNo).trim() && 
        String(data[i][{no_idx}]).trim() === String(ogrNo).trim()) {{
      return {{ success: true, ad: data[i][{nm_idx}], no: data[i][{no_idx}], not: data[i][{gr_idx}] }};
    }}
  }}
  return {{ success: false, message: "{self.error_msg_input.text()}" }};
}}"""

        html_code = f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>{title}</title>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    :root {{ 
        --p: #004a99; --p-d: #003366; --bg: #f4f7f9; --txt: #1a202c; 
        --card: #ffffff; --border: #e2e8f0; --title: #004a99; --sub: #718096;
    }}
    @media (prefers-color-scheme: dark) {{ 
        :root {{ 
            --bg: #1a202c; --txt: #f7fafc; --card: #2d3748; 
            --border: #4a5568; --title: #ffffff; --sub: #a0aec0;
        }} 
    }}
    body {{ font-family: 'Inter', sans-serif; background: var(--bg); color: var(--txt); display: flex; justify-content: center; align-items: center; min-height: 100vh; margin: 0; }}
    .card {{ background: var(--card); padding: 40px; border-radius: 20px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); max-width: 420px; width: 90%; text-align: center; border: 1px solid var(--border); transition: 0.3s; }}
    .logo {{ width: 80px; height: 80px; margin-bottom: 20px; }}
    h1 {{ font-size: 24px; font-weight: 700; margin: 0; color: var(--title); }}
    .subtitle {{ font-size: 14px; color: var(--sub); margin: 10px 0 30px; }}
    .input-group {{ text-align: left; margin-bottom: 20px; }}
    label {{ display: block; font-size: 13px; font-weight: 600; margin-bottom: 8px; color: var(--txt); }}
    input {{ width: 100%; padding: 12px; border: 1px solid var(--border); border-radius: 10px; background: var(--bg); color: var(--txt); box-sizing: border-box; outline: none; transition: 0.2s; }}
    input:focus {{ border-color: var(--p); box-shadow: 0 0 0 3px rgba(0, 74, 153, 0.2); }}
    button {{ width: 100%; padding: 15px; background: var(--p); color: #fff; border: none; border-radius: 10px; font-weight: 600; cursor: pointer; transition: 0.2s; }}
    button:hover {{ background: var(--p-d); transform: translateY(-1px); }}
    #res {{ margin-top: 25px; padding: 20px; border-radius: 12px; display: none; text-align: left; border: 1px solid var(--border); }}
    .success {{ background: rgba(72, 187, 120, 0.1); border-color: #48bb78 !important; }}
    .error {{ background: rgba(245, 101, 101, 0.1); border-color: #f56565 !important; color: #f56565; }}
    .grade {{ background: var(--p); color: #fff; padding: 20px; border-radius: 10px; margin-top: 15px; text-align: center; }}
    .g-val {{ font-size: 32px; font-weight: 800; }}
  </style>
</head>
<body>
<div class="card">
  <img src="{logo}" alt="Logo" class="logo">
  <h1>{title}</h1>
  <p class="subtitle">Öğrenci Sınav Sonuç Sorgulama</p>
  <div class="input-group">
    <label>TC Kimlik Numarası</label>
    <input type="text" id="tc" placeholder="11 Haneli TC" maxlength="11">
  </div>
  <div class="input-group">
    <label>Öğrenci Numarası</label>
    <input type="text" id="no" placeholder="Öğrenci No Giriniz">
  </div>
  <button onclick="sorgula()" id="btn">Sonucu Sorgula</button>
  <div id="res"></div>
</div>
<script>
  function sorgula() {{
    var tc = document.getElementById('tc').value;
    var no = document.getElementById('no').value;
    if(tc.length < 11 || no === "") {{ alert('Lütfen bilgileri eksiksiz girin.'); return; }}
    document.getElementById('btn').disabled = true;
    google.script.run.withSuccessHandler(show).searchStudent(tc, no);
  }}
  function show(r) {{
    document.getElementById('btn').disabled = false;
    var b = document.getElementById('res'); b.style.display = 'block';
    if(r.success) {{
      b.className = 'success';
      b.innerHTML = '<div><b>Ad Soyad:</b> '+r.ad+'</div>' +
                    '<div><b>No:</b> '+r.no+'</div>' +
                    '<div class="grade"><div>SINAV NOTU</div><div class="g-val">'+r.not+'</div></div>';
    }} else {{ b.className = 'error'; b.innerHTML = r.message; }}
  }}
</script>
</body>
</html>"""
        self.gs_output.setPlainText(gs_code); self.html_output.setPlainText(html_code)

    def preview_html(self):
        content = self.html_output.toPlainText()
        if not content: return
        with tempfile.NamedTemporaryFile('w', delete=False, suffix='.html', encoding='utf-8') as f:
            f.write(content); webbrowser.open(f"file://{f.name}")

    def save_files(self):
        if not self.gs_output.toPlainText(): return
        folder = QFileDialog.getExistingDirectory(self, "Kaydedilecek Klasörü Seçin")
        if folder:
            try:
                with open(os.path.join(folder, "Kod.gs"), "w", encoding="utf-8") as f: f.write(self.gs_output.toPlainText())
                with open(os.path.join(folder, "Index.html"), "w", encoding="utf-8") as f: f.write(self.html_output.toPlainText())
                QMessageBox.information(self, "Başarılı", "Dosyalar kaydedildi!")
            except Exception as e: QMessageBox.critical(self, "Hata", str(e))

    def toggle_theme(self):
        self.is_dark_theme = not self.is_dark_theme; self.apply_theme()

    def apply_theme(self):
        bg, txt = ("#1e1e1e", "#e0e0e0") if self.is_dark_theme else ("#f5f6f7", "#2c3e50")
        self.setStyleSheet(f"background: {bg}; color: {txt};")
        self.main_widget.setStyleSheet(f"QLineEdit, QComboBox, QTextEdit {{ background: {'#2d2d2d' if self.is_dark_theme else '#fff'}; color: {txt}; border: 1px solid #ccc; border-radius: 4px; padding: 5px; }}")

    def load_settings(self):
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    s = json.load(f)
                    self.exam_title_input.setText(s.get("t","Sınav Başlığı"))
            except: pass
    def save_settings(self):
        with open(self.settings_file, "w", encoding="utf-8") as f: json.dump({"t":self.exam_title_input.text()}, f)
    def closeEvent(self, e): self.save_settings(); e.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv); app.setStyle('Fusion'); ex = AppGenerator(); ex.show(); sys.exit(app.exec())
