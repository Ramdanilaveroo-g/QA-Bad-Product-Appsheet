"""
================================================================
  QA INBOUND BOT — Telegram → Google Sheets
  Stack: Tesseract OCR + OpenCV QR + Regex + gspread
  100% Gratis | No AI API | No Rate Limit | Railway Ready
================================================================
  Kolom Sheet "HUB":
  A: HUB | B: Time Stamp | C: QR Code | D: Qty |
  E: MFG Date | F: Foto Product | G: VID | H: Email

  Sheet "Users":
  A: TelegramID | B: Name | C: Email
================================================================
"""

import os, re, json, logging, io
from datetime import datetime

import cv2
import numpy as np
import pytesseract
import gspread
from PIL import Image
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardRemove
from telegram.ext import (
    Application, CommandHandler, MessageHandler,
    CallbackQueryHandler, ConversationHandler, filters, ContextTypes,
)

# ════════════════════════════════════════════
#  CONFIG
# ════════════════════════════════════════════
TELEGRAM_TOKEN    = os.getenv("TELEGRAM_TOKEN", "")
GOOGLE_CREDS_JSON = os.getenv("GOOGLE_CREDENTIALS_JSON", "")

SPREADSHEET_ID = "1kdxdLz4PUxCDZjaLGrKRITPZTR0yETfh7Z7lCpXT7gk"
SHEET_DATA     = "HUB"
SHEET_USERS    = "Users"
GDRIVE_FOLDER  = ""

DEFAULT_HUB = "MTG - Menteng"
DEFAULT_QTY = "1"

# ════════════════════════════════════════════
#  STATUS
# ════════════════════════════════════════════
STATUS_IDX = {
    "s0":  "INBOUND Busuk / Berjamur / Bau Menyengat",
    "s1":  "INBOUND Tidak Segar / Layu / Keriput",
    "s2":  "INBOUND Memar / Pecah / Luka Mekanis",
    "s3":  "INBOUND Mencair / Thawing",
    "s4":  "INBOUND Terlalu Matang / Melebihi Step Kematangan",
    "s5":  "INBOUND Tidak Matang / Dibawah Step Kematangan / Gagal matang",
    "s6":  "INBOUND Kontaminasi : Rambut, Batu, Hama, Parasit, Benda Asing, Pestisida, Kotor",
    "s7":  "INBOUND Kemasan Rusak / Robek / Berlubang / Loss Vacum",
    "s8":  "INBOUND Melewati MSLTC / Mendekati Kadaluarsa",
    "s9":  "INBOUND Produk Kadaluarsa",
    "s10": "Badstock Karantina",
    "s11": "Badstock Training",
    "s12": "Dispose di HUB",
}
STATUS_KW = {
    "busuk":"s0","jamur":"s0","bau":"s0",
    "layu":"s1","segar":"s1","keriput":"s1",
    "memar":"s2","pecah":"s2","luka":"s2","mekanis":"s2",
    "cair":"s3","mencair":"s3","thawing":"s3",
    "matang":"s4","overripe":"s4",
    "belum":"s5","mentah":"s5","gagal":"s5",
    "kontaminasi":"s6","hama":"s6","rambut":"s6","parasit":"s6",
    "kemasan":"s7","rusak":"s7","robek":"s7","berlubang":"s7","vacum":"s7",
    "msltc":"s8",
    "kadaluarsa":"s9","expired":"s9","exp":"s9",
    "karantina":"s10","training":"s11",
    "dispose":"s12","buang":"s12",
}

# ════════════════════════════════════════════
#  STATES
# ════════════════════════════════════════════
(
    WAIT_EMAIL, WAIT_EMAIL_CONFIRM,
    WAIT_LABEL, WAIT_STATUS, WAIT_STATUS_TYPE,
    WAIT_CONFIRM, WAIT_EDIT_FIELD, WAIT_EDIT_VALUE,
    WAIT_PRODUCT_PHOTO,
) = range(9)

# ════════════════════════════════════════════
#  LOGGING
# ════════════════════════════════════════════
logging.basicConfig(format="%(asctime)s | %(levelname)s | %(message)s", level=logging.INFO)
log = logging.getLogger(__name__)

# ════════════════════════════════════════════
#  GOOGLE CLIENTS
# ════════════════════════════════════════════
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def get_creds():
    return Credentials.from_service_account_info(json.loads(GOOGLE_CREDS_JSON), scopes=SCOPES)

def get_gc():
    return gspread.authorize(get_creds())

def get_drive():
    return build("drive", "v3", credentials=get_creds())

# ════════════════════════════════════════════
#  USER STORAGE — Sheet "Users"
# ════════════════════════════════════════════
def ensure_users_sheet():
    try:
        gc = get_gc()
        sh = gc.open_by_key(SPREADSHEET_ID)
        try:
            sh.worksheet(SHEET_USERS)
        except gspread.WorksheetNotFound:
            ws = sh.add_worksheet(title=SHEET_USERS, rows=1000, cols=3)
            ws.append_row(["TelegramID", "Name", "Email"])
    except Exception as e:
        log.error(f"ensure_users_sheet: {e}")

def get_user(tid: int):
    try:
        ws = get_gc().open_by_key(SPREADSHEET_ID).worksheet(SHEET_USERS)
        for r in ws.get_all_records():
            if str(r.get("TelegramID")) == str(tid):
                return r
    except Exception as e:
        log.error(f"get_user: {e}")
    return None

def save_user(tid: int, name: str, email: str):
    try:
        ws   = get_gc().open_by_key(SPREADSHEET_ID).worksheet(SHEET_USERS)
        rows = ws.get_all_records()
        for i, r in enumerate(rows, start=2):
            if str(r.get("TelegramID")) == str(tid):
                ws.update(f"A{i}:C{i}", [[str(tid), name, email]])
                return
        ws.append_row([str(tid), name, email])
    except Exception as e:
        log.error(f"save_user: {e}")

# ════════════════════════════════════════════
#  CORE: Preprocess image untuk OCR
# ════════════════════════════════════════════
def preprocess(image_bytes: bytes) -> np.ndarray:
    """Konversi foto ke grayscale + threshold → teks lebih jelas dibaca."""
    nparr = np.frombuffer(image_bytes, np.uint8)
    img   = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

    # Resize kalau terlalu besar
    h, w  = img.shape[:2]
    if max(h, w) > 2000:
        scale = 2000 / max(h, w)
        img   = cv2.resize(img, (int(w*scale), int(h*scale)))

    # Upscale 2x buat akurasi OCR
    img   = cv2.resize(img, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)

    # Grayscale
    gray  = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Sharpen
    kernel = np.array([[0,-1,0],[-1,5,-1],[0,-1,0]])
    sharp  = cv2.filter2D(gray, -1, kernel)

    # Adaptive threshold
    thresh = cv2.adaptiveThreshold(
        sharp, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 31, 10
    )
    return thresh

# ════════════════════════════════════════════
#  CORE: Scan QR Code pakai OpenCV
# ════════════════════════════════════════════
def scan_qr(image_bytes: bytes) -> str:
    """Scan QR/barcode dari foto. Return isi QR atau string kosong."""
    try:
        nparr = np.frombuffer(image_bytes, np.uint8)
        img   = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

        detector = cv2.QRCodeDetector()

        # Coba beberapa ukuran
        for scale in [1.0, 1.5, 2.0, 0.75]:
            h, w    = img.shape[:2]
            resized = cv2.resize(img, (int(w*scale), int(h*scale)))
            data, _, _ = detector.detectAndDecode(resized)
            if data:
                log.info(f"QR scan OK: {data}")
                return data.strip()

        # Coba juga pakai WeChatQRCode kalau tersedia
        try:
            wechat = cv2.wechat_qrcode_WeChatQRCode()
            texts, _ = wechat.detectAndDecode(img)
            if texts:
                log.info(f"WeChatQR: {texts[0]}")
                return texts[0].strip()
        except:
            pass

    except Exception as e:
        log.warning(f"QR scan error: {e}")
    return ""

# ════════════════════════════════════════════
#  CORE: OCR + Regex Parsing label Astro
# ════════════════════════════════════════════
MONTH_MAP = {
    "jan":"01","feb":"02","mar":"03","apr":"04",
    "may":"05","mei":"05","jun":"06","jul":"07",
    "aug":"08","agu":"08","sep":"09","oct":"10",
    "okt":"10","nov":"11","dec":"12","des":"12",
}

def parse_date(s: str) -> str:
    """Konversi berbagai format tanggal ke DD/MM/YYYY."""
    s = s.strip()
    # Format: 16-May-2026 atau 16-05-2026
    m = re.match(r"(\d{1,2})[-/](\w+)[-/](\d{4})", s)
    if m:
        day, mon, year = m.group(1), m.group(2).lower(), m.group(3)
        if mon.isdigit():
            return f"{day.zfill(2)}/{mon.zfill(2)}/{year}"
        mon_num = MONTH_MAP.get(mon[:3], "")
        if mon_num:
            return f"{day.zfill(2)}/{mon_num}/{year}"
    # Format: DD/MM/YYYY
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{4})", s)
    if m:
        return f"{m.group(1).zfill(2)}/{m.group(2).zfill(2)}/{m.group(3)}"
    return s

def ocr_and_parse(image_bytes: bytes, sku_from_qr: str = "") -> dict:
    """
    OCR foto label pakai Tesseract → parse dengan regex.
    Format label Astro:
      Tomat Merah Astro Farm 500gram   ← product name
      ExpDate 29-May-2026               ← abaikan
      VID: 5051                         ← vendor_id
      485802;29052026                   ← SKU;ddmmyyyy
      MFG 16-May-2026                   ← mfg_date
    """
    result = {
        "sku_number":   sku_from_qr,
        "product_name": "",
        "mfg_date":     "",
        "vendor_id":    "",
    }

    try:
        processed = preprocess(image_bytes)
        # OCR config: PSM 6 = assume uniform block of text
        config    = "--oem 3 --psm 6 -l eng"
        raw_text  = pytesseract.image_to_string(processed, config=config)
        log.info(f"OCR raw:\n{raw_text}")

        lines = [l.strip() for l in raw_text.splitlines() if l.strip()]

        # 1. SKU: dari QR atau cari pola angka;angka
        if not result["sku_number"]:
            for line in lines:
                m = re.search(r"(\d{5,})\s*[;:]\s*\d+", line)
                if m:
                    result["sku_number"] = m.group(1)
                    break
            # Fallback: cari baris yang isinya angka panjang saja
            if not result["sku_number"]:
                for line in lines:
                    m = re.match(r"^(\d{5,})\s*$", line)
                    if m:
                        result["sku_number"] = m.group(1)
                        break

        # 2. VID: cari "VID" diikuti angka
        for line in lines:
            m = re.search(r"VID\s*[:\-]?\s*(\d{3,6})", line, re.IGNORECASE)
            if m:
                result["vendor_id"] = m.group(1)
                break

        # 3. MFG Date
        for line in lines:
            m = re.search(r"MFG\s+(\d{1,2}[-/]\w+[-/]\d{4})", line, re.IGNORECASE)
            if m:
                result["mfg_date"] = parse_date(m.group(1))
                break
        # Fallback MFG format lain
        if not result["mfg_date"]:
            for line in lines:
                m = re.search(r"MFG\s+(\d{1,2}[/-]\d{1,2}[/-]\d{4})", line, re.IGNORECASE)
                if m:
                    result["mfg_date"] = parse_date(m.group(1))
                    break

        # 4. Product name: baris pertama yang bukan angka & bukan keyword
        skip_kw = {"expdate","exp","vid","mfg","astro","date","barcode"}
        for line in lines:
            line_low = line.lower()
            if any(kw in line_low for kw in skip_kw):
                continue
            if re.match(r"^\d+", line):
                continue
            if len(line) > 5:
                # Bersihkan noise OCR
                clean = re.sub(r"[^a-zA-Z0-9 /.]", "", line).strip()
                if len(clean) > 5:
                    result["product_name"] = clean
                    break

        log.info(f"Parsed: {result}")

    except Exception as e:
        log.error(f"OCR error: {e}")

    return result

# ════════════════════════════════════════════
#  CORE: Upload foto ke Google Drive
# ════════════════════════════════════════════
def upload_photo(image_bytes: bytes, filename: str) -> str:
    try:
        svc   = get_drive()
        meta  = {"name": filename}
        if GDRIVE_FOLDER:
            meta["parents"] = [GDRIVE_FOLDER]
        media = MediaIoBaseUpload(io.BytesIO(image_bytes), mimetype="image/jpeg", resumable=True)
        f     = svc.files().create(body=meta, media_body=media, fields="id").execute()
        fid   = f.get("id")
        svc.permissions().create(fileId=fid, body={"type":"anyone","role":"reader"}).execute()
        url   = f"https://drive.google.com/uc?id={fid}"
        log.info(f"Uploaded: {url}")
        return url
    except Exception as e:
        log.error(f"Upload error: {e}")
        return ""

# ════════════════════════════════════════════
#  CORE: Append ke Google Sheets
# ════════════════════════════════════════════
def append_sheet(data: dict, photo_url: str = "", email: str = "") -> tuple[bool, str]:
    try:
        ws  = get_gc().open_by_key(SPREADSHEET_ID).worksheet(SHEET_DATA)
        ts  = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        row = [
            DEFAULT_HUB,
            ts,
            data.get("sku_number", ""),
            data.get("qty", DEFAULT_QTY),
            data.get("mfg_date", ""),
            photo_url,
            data.get("vendor_id", ""),
            email,
        ]
        ws.append_row(row, value_input_option="USER_ENTERED")
        log.info("Row appended OK")
        return True, "OK"
    except Exception as e:
        log.error(f"Sheets error: {e}")
        return False, str(e)

# ════════════════════════════════════════════
#  KEYBOARDS & UTILS
# ════════════════════════════════════════════
def status_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🦠 Busuk/Jamur/Bau",   callback_data="s0"),
         InlineKeyboardButton("🥀 Layu/Keriput",       callback_data="s1")],
        [InlineKeyboardButton("🤕 Memar/Pecah",        callback_data="s2"),
         InlineKeyboardButton("💧 Mencair/Thawing",    callback_data="s3")],
        [InlineKeyboardButton("🍊 Terlalu Matang",     callback_data="s4"),
         InlineKeyboardButton("🌱 Belum Matang",       callback_data="s5")],
        [InlineKeyboardButton("🐛 Kontaminasi",        callback_data="s6"),
         InlineKeyboardButton("📦 Kemasan Rusak",      callback_data="s7")],
        [InlineKeyboardButton("⏰ Mend. Kadaluarsa",   callback_data="s8"),
         InlineKeyboardButton("💀 Produk Kadaluarsa",  callback_data="s9")],
        [InlineKeyboardButton("🔒 Badstock Karantina", callback_data="s10"),
         InlineKeyboardButton("📚 Badstock Training",  callback_data="s11")],
        [InlineKeyboardButton("🗑️ Dispose di HUB",    callback_data="s12")],
        [InlineKeyboardButton("⌨️ Ketik Manual",       callback_data="stype")],
    ])

def confirm_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Konfirmasi — Kirim Foto Produk", callback_data="ok")],
        [InlineKeyboardButton("🔄 Ganti Status", callback_data="chg"),
         InlineKeyboardButton("✏️ Edit Data",    callback_data="edit")],
        [InlineKeyboardButton("❌ Batal",         callback_data="batal")],
    ])

def edit_kb():
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("🔢 SKU / QR Code", callback_data="ed_sku")],
        [InlineKeyboardButton("📦 Nama Produk",   callback_data="ed_name")],
        [InlineKeyboardButton("📅 MFG Date",      callback_data="ed_mfg")],
        [InlineKeyboardButton("🏭 VID",           callback_data="ed_vid")],
        [InlineKeyboardButton("📈 Qty",           callback_data="ed_qty")],
        [InlineKeyboardButton("◀️ Kembali",       callback_data="back")],
    ])

EDIT_FIELDS = {
    "ed_sku":  ("sku_number",   "SKU / QR Code"),
    "ed_name": ("product_name", "Nama Produk"),
    "ed_mfg":  ("mfg_date",     "MFG Date (DD/MM/YYYY)"),
    "ed_vid":  ("vendor_id",    "VID"),
    "ed_qty":  ("qty",          "Qty"),
}

def find_status(text: str):
    tl = text.lower().strip()
    for v in STATUS_IDX.values():
        if tl == v.lower(): return v
    for kw, idx in STATUS_KW.items():
        if kw in tl: return STATUS_IDX[idx]
    for v in STATUS_IDX.values():
        if tl in v.lower(): return v
    return None

def summary(d: dict) -> str:
    return (
        "📋 *Ringkasan Data:*\n\n"
        f"🏪 Hub:      `{DEFAULT_HUB}`\n"
        f"🔢 QR/SKU:   `{d.get('sku_number') or '⚠️ kosong'}`\n"
        f"📦 Produk:   `{d.get('product_name') or '⚠️ kosong'}`\n"
        f"📈 Qty:      `{d.get('qty', DEFAULT_QTY)}`\n"
        f"🚦 Status:   `{d.get('status') or '⚠️ belum diisi'}`\n"
        f"📅 MFG Date: `{d.get('mfg_date') or '⚠️ kosong'}`\n"
        f"🏭 VID:      `{d.get('vendor_id') or '⚠️ kosong'}`\n"
    )

# ════════════════════════════════════════════
#  HANDLERS
# ════════════════════════════════════════════
async def cmd_start(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    uid  = update.effective_user.id
    user = get_user(uid)
    if user:
        await update.message.reply_text(
            f"👋 Halo *{user['Name']}*!\n"
            f"📧 Email: `{user['Email']}`\n\n"
            "📷 Kirim *foto label produk* untuk mulai.\n"
            "_/tips untuk panduan foto | /gantiemail untuk ganti email_",
            parse_mode="Markdown",
        )
        return WAIT_LABEL
    await update.message.reply_text(
        "🌿 *QA Inbound Bot — MTG Menteng*\n\n"
        "Registrasi dulu, *sekali aja* ya!\n\n"
        "📧 Ketik *email Google* lo yang dipakai login AppSheet:",
        parse_mode="Markdown",
    )
    return WAIT_EMAIL

async def cmd_tips(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📷 *Tips Foto Label:*\n\n"
        "✅ Foto label dari *dekat & fokus*\n"
        "✅ Label isi hampir seluruh frame\n"
        "✅ Cahaya cukup terang\n"
        "✅ Label tidak terlipat\n\n"
        "❌ Jangan foto dari jauh\n"
        "❌ Jangan blur\n\n"
        "💡 Data salah → pakai ✏️ *Edit Data*",
        parse_mode="Markdown",
    )

async def cmd_ganti_email(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    user = get_user(uid)
    old  = user["Email"] if user else "-"
    await update.message.reply_text(
        f"📧 Email sekarang: `{old}`\n\nKetik email baru:",
        parse_mode="Markdown",
    )
    return WAIT_EMAIL

async def handle_email_input(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    email = update.message.text.strip().lower()
    if not re.match(r"^[\w\.-]+@[\w\.-]+\.\w+$", email):
        await update.message.reply_text("❌ Format email salah. Contoh: _nama@gmail.com_\n\nCoba lagi:", parse_mode="Markdown")
        return WAIT_EMAIL
    ctx.user_data["pending_email"] = email
    await update.message.reply_text(
        f"📧 Email: `{email}`\n\nSudah benar?",
        parse_mode="Markdown",
        reply_markup=InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Benar", callback_data="email_ok"),
             InlineKeyboardButton("✏️ Ketik Ulang", callback_data="email_retry")],
        ]),
    )
    return WAIT_EMAIL_CONFIRM

async def handle_email_confirm(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == "email_retry":
        await q.edit_message_text("📧 Ketik ulang email lo:")
        return WAIT_EMAIL
    uid   = update.effective_user.id
    name  = update.effective_user.first_name or "QA User"
    email = ctx.user_data.pop("pending_email", "")
    save_user(uid, name, email)
    await q.edit_message_text(
        f"✅ *Registrasi berhasil!*\n\n👤 `{name}`\n📧 `{email}`\n\n"
        "📷 Kirim *foto label produk* untuk mulai!",
        parse_mode="Markdown",
    )
    return WAIT_LABEL

async def handle_label(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    uid  = update.effective_user.id
    user = get_user(uid)
    if not user:
        await update.message.reply_text("⚠️ Belum registrasi. Ketik /start dulu!")
        return WAIT_LABEL

    wait = await update.message.reply_text("🔍 Scanning QR & membaca label...")
    try:
        pf   = await update.message.photo[-1].get_file()
        img  = bytes(await pf.download_as_bytearray())

        # Step 1: Scan QR dulu (paling akurat untuk SKU)
        qr_data = scan_qr(img)
        sku_qr  = ""
        if qr_data:
            # QR Astro format: "angka;ddmmyyyy" atau langsung angka
            m = re.match(r"(\d+)[;:]", qr_data)
            sku_qr = m.group(1) if m else qr_data.split(";")[0].strip()
            log.info(f"SKU dari QR: {sku_qr}")

        # Step 2: OCR + regex parsing
        data = ocr_and_parse(img, sku_from_qr=sku_qr)
        data["qty"] = DEFAULT_QTY
        ctx.user_data["data"] = data

        qr_info = f"✅ QR scan: `{sku_qr}`" if sku_qr else "⚠️ QR tidak terbaca, pakai OCR"
        await wait.delete()
        await update.message.reply_text(
            f"📷 *Label dibaca!*\n_{qr_info}_\n\n"
            f"🔢 QR/SKU:    `{data.get('sku_number') or '⚠️ tidak terbaca'}`\n"
            f"📦 Produk:    `{data.get('product_name') or '⚠️ tidak terbaca'}`\n"
            f"📅 MFG Date:  `{data.get('mfg_date') or '⚠️ tidak terbaca'}`\n"
            f"🏭 VID:       `{data.get('vendor_id') or '⚠️ tidak terbaca'}`\n\n"
            "━━━━━━━━━━━━━━━━━━\n🚦 *Pilih STATUS produk:*",
            parse_mode="Markdown",
            reply_markup=status_kb(),
        )
        return WAIT_STATUS
    except Exception as e:
        await wait.delete()
        log.error(e)
        await update.message.reply_text(f"❌ Error: `{e}`\nCoba foto ulang.", parse_mode="Markdown")
        return WAIT_LABEL

async def handle_status_btn(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == "stype":
        await q.edit_message_text("⌨️ *Ketik keyword status:*\n\nContoh: `memar`, `busuk`, `layu`...", parse_mode="Markdown")
        return WAIT_STATUS_TYPE
    status = STATUS_IDX.get(q.data)
    if not status:
        await q.answer("Tidak valid.", show_alert=True)
        return WAIT_STATUS
    ctx.user_data["data"]["status"] = status
    await q.edit_message_text(
        summary(ctx.user_data["data"]) + "\n📷 _Cek data, lalu kirim foto produk._",
        parse_mode="Markdown", reply_markup=confirm_kb(),
    )
    return WAIT_CONFIRM

async def handle_status_type(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    status = find_status(update.message.text)
    if not status:
        await update.message.reply_text("❓ Tidak ditemukan. Coba: `memar`, `busuk`, `layu`, `kemasan`...", parse_mode="Markdown")
        return WAIT_STATUS_TYPE
    ctx.user_data["data"]["status"] = status
    await update.message.reply_text(
        summary(ctx.user_data["data"]) + "\n📷 _Cek data, lalu kirim foto produk._",
        parse_mode="Markdown", reply_markup=confirm_kb(),
    )
    return WAIT_CONFIRM

async def handle_confirm_btn(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == "ok":
        await q.edit_message_text(
            summary(ctx.user_data["data"]) + "\n✅ *Dikonfirmasi!*\n\n📷 Kirim *foto produk* sekarang.",
            parse_mode="Markdown",
        )
        return WAIT_PRODUCT_PHOTO
    if q.data == "chg":
        await q.edit_message_text("🚦 *Pilih STATUS baru:*", parse_mode="Markdown", reply_markup=status_kb())
        return WAIT_STATUS
    if q.data == "edit":
        await q.edit_message_text("✏️ *Field mana yang mau diedit?*", parse_mode="Markdown", reply_markup=edit_kb())
        return WAIT_EDIT_FIELD
    if q.data == "batal":
        ctx.user_data.clear()
        await q.edit_message_text("❌ Dibatalin. /start untuk mulai lagi.")
        return ConversationHandler.END

async def handle_edit_field(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    if q.data == "back":
        await q.edit_message_text(
            summary(ctx.user_data["data"]) + "\n📷 _Cek data._",
            parse_mode="Markdown", reply_markup=confirm_kb(),
        )
        return WAIT_CONFIRM
    fk, fl = EDIT_FIELDS[q.data]
    ctx.user_data["editing"] = fk
    await q.edit_message_text(f"✏️ Ketik nilai baru untuk *{fl}*:", parse_mode="Markdown")
    return WAIT_EDIT_VALUE

async def handle_edit_val(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    fk = ctx.user_data.pop("editing", None)
    if fk:
        ctx.user_data["data"][fk] = update.message.text.strip()
    await update.message.reply_text(
        summary(ctx.user_data["data"]) + "\n📷 _Cek data._",
        parse_mode="Markdown", reply_markup=confirm_kb(),
    )
    return WAIT_CONFIRM

async def handle_product_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    wait = await update.message.reply_text("⏳ Upload foto & simpan ke Sheets...")
    try:
        uid       = update.effective_user.id
        user      = get_user(uid)
        email     = user["Email"] if user else ""
        pf        = await update.message.photo[-1].get_file()
        img       = bytes(await pf.download_as_bytearray())
        data      = ctx.user_data.get("data", {})
        ts        = datetime.now().strftime("%Y%m%d_%H%M%S")
        sku       = data.get("sku_number", "unknown")
        photo_url = upload_photo(img, f"QA_{sku}_{ts}.jpg")
        ok, msg   = append_sheet(data, photo_url, email)
        await wait.delete()
        if ok:
            foto_info = f"\n🖼️ [Lihat Foto]({photo_url})" if photo_url else ""
            await update.message.reply_text(
                "✅ *BERHASIL input ke Google Sheets!*\n\n"
                + summary(data)
                + f"📧 Email: `{email}`"
                + foto_info
                + "\n\n🔄 Kirim foto label lagi untuk input berikutnya.",
                parse_mode="Markdown",
                disable_web_page_preview=True,
            )
        else:
            await update.message.reply_text(f"❌ *Gagal!*\n`{msg}`", parse_mode="Markdown")
        ctx.user_data.clear()
        return WAIT_LABEL
    except Exception as e:
        await wait.delete()
        log.error(e)
        await update.message.reply_text(f"❌ Error: `{e}`", parse_mode="Markdown")
        return WAIT_LABEL

async def ask_photo(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("📷 Kirim *foto* ya, bukan teks.", parse_mode="Markdown")

async def cmd_cancel(update: Update, ctx: ContextTypes.DEFAULT_TYPE):
    ctx.user_data.clear()
    await update.message.reply_text("❌ Dibatalin. /start untuk mulai lagi.", reply_markup=ReplyKeyboardRemove())
    return ConversationHandler.END

# ════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════
def main():
    ensure_users_sheet()
    app = (
        Application.builder()
        .token(TELEGRAM_TOKEN)
        .connect_timeout(30)
        .read_timeout(30)
        .write_timeout(30)
        .pool_timeout(30)
        .build()
    )
    conv = ConversationHandler(
        entry_points=[
            CommandHandler("start", cmd_start),
            MessageHandler(filters.PHOTO, handle_label),
        ],
        states={
            WAIT_EMAIL:         [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_email_input)],
            WAIT_EMAIL_CONFIRM: [CallbackQueryHandler(handle_email_confirm, pattern="^email_")],
            WAIT_LABEL:         [MessageHandler(filters.PHOTO, handle_label),
                                 MessageHandler(filters.TEXT & ~filters.COMMAND, ask_photo)],
            WAIT_STATUS:        [CallbackQueryHandler(handle_status_btn, pattern=r"^s")],
            WAIT_STATUS_TYPE:   [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_status_type)],
            WAIT_CONFIRM:       [CallbackQueryHandler(handle_confirm_btn, pattern="^(ok|chg|edit|batal)$")],
            WAIT_EDIT_FIELD:    [CallbackQueryHandler(handle_edit_field, pattern="^(ed_|back)")],
            WAIT_EDIT_VALUE:    [MessageHandler(filters.TEXT & ~filters.COMMAND, handle_edit_val)],
            WAIT_PRODUCT_PHOTO: [MessageHandler(filters.PHOTO, handle_product_photo),
                                 MessageHandler(filters.TEXT & ~filters.COMMAND, ask_photo)],
        },
        fallbacks=[CommandHandler("cancel", cmd_cancel)],
        allow_reentry=True,
    )
    app.add_handler(conv)
    app.add_handler(CommandHandler("tips", cmd_tips))
    app.add_handler(CommandHandler("gantiemail", cmd_ganti_email))
    log.info("🤖 QA Bot jalan!")
    app.run_polling(drop_pending_updates=True)

if __name__ == "__main__":
    main()
