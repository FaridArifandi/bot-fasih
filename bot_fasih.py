"""
bot_fasih.py
Bot Approve FASIH SM (Playwright, terhubung ke Chrome debug).

Alur per link di file Excel (kolom: nama | link_fasih):
  1. Buka link assignment FASIH (login SSO otomatis dari file 'sso' bila perlu)
  2. Cek tombol di kanan bawah: sudah Approve atau belum
  3. Belum -> klik Approve (+ konfirmasi). Kalau gagal / ada GALAT,
     buka menu CATATAN, aktifkan switch 'Tampilkan Anomali Usaha dan Keluarga',
     lalu coba Approve lagi.
  4. Hasil dicatat ke hasil_approve.csv (bisa dilanjut kalau bot berhenti).
"""
import os
import re
import sys
import csv
import time
import glob
import random
import socket
import subprocess
from datetime import datetime

from playwright.sync_api import sync_playwright

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# ============================================================
#  KONFIGURASI
# ============================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FILE_INPUT_DEFAULT = "daftar.xlsx"          # file Excel: kolom A = nama, kolom B = link
CDP_PORT = 9223                             # beda dari Bot Nerwana (9222) agar tidak bentrok
PROFILE_DIR = os.path.join(BASE_DIR, "chrome_debug_profile")
SCREENSHOTS_DIR = os.path.join(BASE_DIR, "screenshots")
HASIL_CSV = os.path.join(BASE_DIR, "hasil_approve.csv")
HASIL_XLSX = os.path.join(BASE_DIR, "hasil_approve.xlsx")

JEDA_ANTAR_LINK_MIN = 2.0                   # detik, acak antar link
JEDA_ANTAR_LINK_MAX = 3.5
TIMEOUT_SIAP = 45                           # detik menunggu halaman assignment siap
MAX_PERCOBAAN_APPROVE = 3                   # percobaan approve per link (dengan perbaikan switch di antaranya)
MAX_ERR_BERTURUT = 3                        # berhenti jika N error teknis berturut-turut
SIMPAN_XLSX_SETIAP = 50                     # tulis ulang hasil_approve.xlsx tiap N link

STATUS_FINAL = {"APPROVED", "SUDAH_APPROVE"}
STATUS_ERROR = {"GAGAL_APPROVE", "GALAT", "ERROR_HALAMAN", "TIDAK_DIKENALI"}

CHROME_CANDIDATES = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
]

os.makedirs(SCREENSHOTS_DIR, exist_ok=True)


class SSOGagal(Exception):
    pass


# ============================================================
#  UTIL
# ============================================================
def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def is_port_open(port):
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(("127.0.0.1", port)) == 0


def start_debug_chrome():
    if is_port_open(CDP_PORT):
        return True
    chrome = next((c for c in CHROME_CANDIDATES if os.path.exists(c)), None)
    if not chrome:
        print("[ERROR] Google Chrome tidak ditemukan di folder standar. Install Chrome dulu.")
        return False
    print(f"[CHROME] Membuka Chrome debug (port {CDP_PORT})...")
    subprocess.Popen([
        chrome,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={PROFILE_DIR}",
        "--start-maximized",
    ])
    for _ in range(20):
        time.sleep(1)
        if is_port_open(CDP_PORT):
            print("[CHROME] Siap terhubung.")
            return True
    print("[ERROR] Chrome gagal dibuka pada port debug.")
    return False


def get_sso_credentials():
    sso_file = next((os.path.join(BASE_DIR, n) for n in ("sso", ".sso", "sso.txt")
                     if os.path.exists(os.path.join(BASE_DIR, n))), None)
    if not sso_file:
        return None, None
    u, p = "", ""
    with open(sso_file, "r", encoding="utf-8-sig") as f:
        for line in f:
            if "=" in line:
                k, v = line.split("=", 1)
                k = k.strip().lower()
                v = v.strip()
                if "user" in k or "uname" in k:
                    u = v
                elif "pass" in k:
                    p = v
    return (u, p) if u and p else (None, None)


def screenshot(page, prefix):
    try:
        path = os.path.join(SCREENSHOTS_DIR, f"{prefix}_{int(time.time())}.png")
        page.screenshot(path=path, timeout=5000)
        # simpan max 60 screenshot terbaru
        files = sorted(glob.glob(os.path.join(SCREENSHOTS_DIR, "*.png")), key=os.path.getmtime)
        for old in files[:-60]:
            try:
                os.remove(old)
            except Exception:
                pass
        return path
    except Exception:
        return ""


# ============================================================
#  EXCEL & HASIL
# ============================================================
def cari_file_input():
    p = os.path.join(BASE_DIR, FILE_INPUT_DEFAULT)
    if os.path.exists(p):
        return p
    kandidat = [f for f in glob.glob(os.path.join(BASE_DIR, "*.xlsx"))
                if not os.path.basename(f).startswith(("~$", "hasil_"))]
    if len(kandidat) == 1:
        return kandidat[0]
    if kandidat:
        print("File Excel ditemukan:")
        for i, f in enumerate(kandidat, 1):
            print(f"  [{i}] {os.path.basename(f)}")
        pilih = input("Pilih nomor file input (default 1): ").strip() or "1"
        return kandidat[int(pilih) - 1]
    return None


def baca_excel(path):
    """Mengembalikan list dict {no, nama, link}. Kolom dikenali dari header; fallback kolom A/B."""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return []
    header = [str(c or "").strip().lower() for c in rows[0]]
    idx_link = next((i for i, h in enumerate(header) if "link" in h), None)
    idx_nama = next((i for i, h in enumerate(header) if i != idx_link and h), 0)
    if idx_link is None:
        idx_link = 1
    out = []
    for no, r in enumerate(rows[1:], start=2):
        if r is None or len(r) <= idx_link:
            continue
        link = str(r[idx_link] or "").strip()
        if not link.lower().startswith("http"):
            continue
        nama = str(r[idx_nama] or "").strip() if idx_nama < len(r) else ""
        out.append({"no": no, "nama": nama, "link": link})
    return out


def muat_hasil():
    """link -> (status, catatan) terakhir dari hasil_approve.csv"""
    hasil = {}
    if os.path.exists(HASIL_CSV):
        with open(HASIL_CSV, "r", encoding="utf-8-sig", newline="") as f:
            for r in csv.DictReader(f):
                hasil[r["link"]] = (r["status"], r.get("catatan", ""), r.get("waktu", ""))
    return hasil


def catat_hasil(row, status, catatan=""):
    baru = not os.path.exists(HASIL_CSV)
    with open(HASIL_CSV, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.writer(f)
        if baru:
            w.writerow(["waktu", "no", "nama", "link", "status", "catatan"])
        w.writerow([now(), row["no"], row["nama"], row["link"], status, catatan])


def simpan_xlsx(baris):
    try:
        import openpyxl
        hasil = muat_hasil()
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Hasil"
        ws.append(["No", "Nama", "Link", "Status", "Catatan", "Waktu"])
        for r in baris:
            st, cat, wk = hasil.get(r["link"], ("BELUM_DIPROSES", "", ""))
            ws.append([r["no"], r["nama"], r["link"], st, cat, wk])
        for col, w in zip("ABCDEF", (6, 40, 70, 18, 60, 20)):
            ws.column_dimensions[col].width = w
        wb.save(HASIL_XLSX)
    except PermissionError:
        print("[WARN] hasil_approve.xlsx sedang terbuka di Excel, dilewati (data aman di hasil_approve.csv).")
    except Exception as e:
        print(f"[WARN] Gagal menyimpan hasil_approve.xlsx: {e}")


# ============================================================
#  LOGIN SSO
# ============================================================
_sso_gagal = 0


def di_halaman_sso(page):
    try:
        if "sso.bps.go.id" in page.url.lower():
            return True
        return page.locator("#kc-login").count() > 0
    except Exception:
        return False


def di_layar_login_fasih(page):
    try:
        btn = page.locator('button:has-text("Lanjutkan dengan SSO"), a:has-text("Lanjutkan dengan SSO")')
        return (btn.count() > 0 and btn.first.is_visible()) or page.locator('text="Selamat Datang Kembali"').count() > 0
    except Exception:
        return False


def isi_form_sso(page):
    global _sso_gagal
    if _sso_gagal >= 2:
        raise SSOGagal("Login SSO gagal 2x berturut-turut. Cek username/password di file 'sso', "
                       "atau login manual dulu di Chrome yang terbuka.")
    uname, pwd = get_sso_credentials()
    if not uname or not pwd:
        raise SSOGagal("Kredensial tidak ditemukan. Isi file 'sso' (username=... / password=...).")

    print("[SSO] Form SSO BPS terdeteksi, login otomatis...")
    page.keyboard.press("Escape")
    time.sleep(0.5)
    inp_user = page.locator('#username, input[name="username"], input[type="text"]').first
    inp_pass = page.locator('#password, input[name="password"], input[type="password"]').first
    if inp_user.count() > 0:
        inp_user.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.keyboard.type(uname, delay=random.randint(60, 120))
        time.sleep(0.3)
    if inp_pass.count() > 0:
        inp_pass.click()
        page.keyboard.press("Control+A")
        page.keyboard.press("Backspace")
        page.keyboard.type(pwd, delay=random.randint(60, 120))
        time.sleep(0.3)
    time.sleep(random.uniform(1.0, 1.6))
    page.locator('#kc-login, button:has-text("Log In"), input[type="submit"]').first.click()
    try:
        page.wait_for_url(lambda u: "sso.bps.go.id" not in u.lower(), timeout=20000)
        _sso_gagal = 0
        print("[SSO] Login berhasil.")
    except Exception:
        _sso_gagal += 1
        print(f"[SSO] Masih di halaman SSO (gagal ke-{_sso_gagal}).")
    time.sleep(2.5)


# ============================================================
#  MEMBACA HALAMAN
# ============================================================
JS_BACA_TOMBOL = """() => {
  document.querySelectorAll('[data-bf]').forEach(e => e.removeAttribute('data-bf'));
  const vw = window.innerWidth, vh = window.innerHeight;
  const out = []; let i = 0;
  for (const el of document.querySelectorAll('button, [role="button"], a[class*="btn"], a[class*="button"]')) {
    const r = el.getBoundingClientRect();
    if (r.width < 8 || r.height < 8) continue;
    const cs = getComputedStyle(el);
    if (cs.visibility === 'hidden' || cs.display === 'none' || parseFloat(cs.opacity) === 0) continue;
    if (r.bottom < 0 || r.top > vh || r.right < 0 || r.left > vw) continue;
    const text = (el.innerText || el.textContent || '').trim().replace(/\\s+/g, ' ');
    el.setAttribute('data-bf', String(i));
    out.push({i: i, text: text, title: el.getAttribute('title') || '', aria: el.getAttribute('aria-label') || '',
              x: r.x + r.width / 2, y: r.y + r.height / 2,
              disabled: el.disabled === true || el.getAttribute('aria-disabled') === 'true'});
    i++;
  }
  return {vw: vw, vh: vh, items: out};
}"""

RE_APPROVE = re.compile(r"\b(approve|setujui)\b", re.I)


def baca_tombol(page):
    """Semua tombol terlihat di layar; ditandai kanan_bawah bila ada di sisi kanan-bawah viewport."""
    try:
        data = page.evaluate(JS_BACA_TOMBOL)
    except Exception:
        return []
    vw, vh = data["vw"], data["vh"]
    for it in data["items"]:
        it["kanan_bawah"] = it["x"] > vw * 0.55 and it["y"] > vh * 0.35
    return data["items"]


def cari_tombol_approve(items):
    """Pilih tombol Approve; utamakan yang di kanan-bawah."""
    cocok = []
    for it in items:
        for teks in (it["text"], it["title"], it["aria"]):
            if teks and len(teks) <= 30 and RE_APPROVE.search(teks):
                cocok.append(it)
                break
    if not cocok:
        return None
    kb = [c for c in cocok if c["kanan_bawah"]]
    pilih = kb or cocok
    pilih.sort(key=lambda c: -c["y"])
    return pilih[0]


def klik_item(page, item):
    loc = page.locator(f'[data-bf="{item["i"]}"]')
    try:
        loc.first.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    try:
        loc.first.click(timeout=5000)
    except Exception:
        page.mouse.click(item["x"], item["y"])


def teks_approved_terbaca(page):
    try:
        return bool(page.evaluate(
            "() => /\\bapproved\\b|\\bdisetujui\\b/i.test(document.body ? document.body.innerText : '')"))
    except Exception:
        return False


def loading_terlihat(page):
    try:
        l = page.locator('text="Loading Page..."')
        return l.count() > 0 and l.first.is_visible()
    except Exception:
        return False


def cek_fasih_error(page):
    try:
        return (page.locator('h1:has-text("Forbidden"), h2:has-text("Forbidden")').count() > 0 or
                page.locator('h1:has-text("403"), h2:has-text("403")').count() > 0 or
                page.locator('text="There\'s some error"').count() > 0 or
                page.locator('h1:has-text("Access Denied")').count() > 0)
    except Exception:
        return False


def tunggu_siap(page, link):
    """
    Menunggu halaman assignment siap sambil menangani SSO.
    Return: BELUM | SUDAH | TANPA_TOMBOL | ERROR_HALAMAN | TIMEOUT | CLOSED
    """
    mulai = time.time()
    goto_ulang = 0
    stabil = 0
    while time.time() - mulai < TIMEOUT_SIAP:
        try:
            if page.is_closed():
                return "CLOSED"

            if di_halaman_sso(page):
                isi_form_sso(page)
                mulai = time.time()
                continue

            if di_layar_login_fasih(page):
                print("[SSO] Layar 'Lanjutkan dengan SSO' terdeteksi, mengklik...")
                time.sleep(random.uniform(1.2, 1.8))
                btn = page.locator('button:has-text("Lanjutkan dengan SSO"), a:has-text("Lanjutkan dengan SSO")') \
                          .filter(has_not_text="Eksternal")
                if btn.count() == 0:
                    btn = page.locator('button:has-text("Lanjutkan dengan SSO"), a:has-text("Lanjutkan dengan SSO")')
                btn.first.click()
                time.sleep(3)
                mulai = time.time()
                continue

            url = page.url.lower()
            if "fasih-sm.bps.go.id" in url and "assignment" not in url and not loading_terlihat(page):
                # Setelah login FASIH sering mendarat di /app/ -> buka ulang link target
                if goto_ulang >= 3:
                    return "ERROR_HALAMAN"
                goto_ulang += 1
                print("[FASIH] Mendarat di beranda FASIH, membuka ulang link assignment...")
                try:
                    page.goto(link, timeout=60000, wait_until="domcontentloaded")
                except Exception:
                    pass
                time.sleep(1.5)
                continue

            if loading_terlihat(page):
                stabil = 0
                time.sleep(1)
                continue

            if cek_fasih_error(page):
                return "ERROR_HALAMAN"

            items = baca_tombol(page)
            if cari_tombol_approve(items):
                return "BELUM"

            if any(it["kanan_bawah"] for it in items):
                # toolbar kanan-bawah sudah ada tapi tanpa tombol Approve -> pastikan stabil 3x
                stabil += 1
                if stabil >= 3:
                    return "SUDAH" if teks_approved_terbaca(page) else "TANPA_TOMBOL"
            else:
                stabil = 0
        except SSOGagal:
            raise
        except Exception:
            pass
        time.sleep(1)
    return "TIMEOUT"


# ============================================================
#  MENU CATATAN & SWITCH ANOMALI  (diambil dari Bot Nerwana)
# ============================================================
def buka_menu_catatan(page):
    print("[FASIH] Membuka menu CATATAN...")
    try:
        page.evaluate("""() => {
            const sidebars = document.querySelectorAll('.fasih-form-sidebar, aside, [class*="sidebar"], [class*="drawer"]');
            sidebars.forEach(s => s.scrollTop = s.scrollHeight);
        }""")
        time.sleep(0.5)

        sidebar_catatan = page.locator('.fasih-form-sidebar [title*="CATATAN" i], [title="CATATAN"], div[title="CATATAN"]')
        if sidebar_catatan.count() > 0:
            for i in range(sidebar_catatan.count()):
                el = sidebar_catatan.nth(i)
                if el.is_visible():
                    el.scroll_into_view_if_needed()
                    time.sleep(0.3)
                    el.click()
                    time.sleep(2)
                    return True

        for el in page.locator("div, span, button, a").filter(has_text="CATATAN").all():
            try:
                if el.is_visible():
                    t = el.inner_text().strip().upper()
                    if t == "CATATAN" or ("CATATAN" in t and len(t) <= 15):
                        el.scroll_into_view_if_needed()
                        time.sleep(0.3)
                        el.click()
                        time.sleep(2)
                        return True
            except Exception:
                pass

        btn_next = page.locator('button:has-text("BERIKUTNYA")')
        if btn_next.count() > 0 and btn_next.first.is_visible():
            if "CATATAN" in btn_next.first.inner_text().upper():
                btn_next.first.click()
                time.sleep(2)
                return True
    except Exception as e:
        print(f"[WARN] Kendala navigasi ke CATATAN: {e}")
    return False


JS_CEK_SWITCH = """() => {
    const all = Array.from(document.querySelectorAll('*')).filter(el => {
        const t = (el.textContent || '').trim();
        return t === 'Tampilkan Anomali Usaha dan Keluarga' || (t.includes('Tampilkan Anomali') && !t.includes('admin') && !t.includes('diselesaikan'));
    });
    if (all.length === 0) return "not_found";
    all.sort((a, b) => a.children.length - b.children.length);
    const labelEl = all[0];

    let container = labelEl.parentElement;
    for (let i = 0; i < 5 && container; i++) {
        const txt = (container.textContent || '');
        if (txt.includes('Wajib diisi')) return "false";
        container = container.parentElement;
    }

    let parent = labelEl.parentElement;
    for (let i = 0; i < 5 && parent; i++) {
        const sw = parent.querySelector('[role="switch"], input[type="checkbox"], button[role="switch"], [id$="-control"]');
        if (sw) {
            const ariaChecked = sw.getAttribute("aria-checked");
            if (ariaChecked === "true") return "true";
            if (ariaChecked === "false") return "false";
            if (sw.checked === true) return "true";
            if (sw.checked === false) return "false";
        }
        parent = parent.parentElement;
    }
    return "true";
}"""

JS_KLIK_SWITCH = """() => {
    const all = Array.from(document.querySelectorAll('*')).filter(el => {
        const t = (el.textContent || '').trim();
        return t === 'Tampilkan Anomali Usaha dan Keluarga' || (t.includes('Tampilkan Anomali') && !t.includes('admin') && !t.includes('diselesaikan'));
    });
    if (all.length === 0) return false;
    all.sort((a, b) => a.children.length - b.children.length);
    const labelEl = all[0];
    let parent = labelEl.parentElement;
    for (let i = 0; i < 5 && parent; i++) {
        const sw = parent.querySelector('[role="switch"], input[type="checkbox"], button[role="switch"], [id$="-control"], span, label');
        if (sw && sw !== labelEl) { sw.click(); return true; }
        parent = parent.parentElement;
    }
    return false;
}"""


def aktifkan_switch_anomali(page):
    """Pastikan toggle 'Tampilkan Anomali Usaha dan Keluarga' ON. Return True bila terverifikasi."""
    print("[FASIH] Memeriksa switch 'Tampilkan Anomali Usaha dan Keluarga'...")
    try:
        status = page.evaluate(JS_CEK_SWITCH)
        if status == "true":
            print("[FASIH] Switch anomali sudah aktif.")
            return True
        if status == "not_found":
            print("[WARN] Label switch anomali tidak ditemukan di halaman ini.")
            return False

        for attempt in range(3):
            print(f"[FASIH] Switch anomali masih OFF, mengaktifkan (percobaan {attempt + 1})...")

            # Tiga metode klik; status dicek SETELAH tiap metode supaya switch
            # tidak ter-toggle dua kali (ON lalu OFF lagi).
            def metode_koordinat():
                label_loc = page.get_by_text("Tampilkan Anomali Usaha dan Keluarga", exact=False).first
                if label_loc.count() > 0 and label_loc.is_visible():
                    box = label_loc.bounding_box()
                    if box:
                        page.mouse.click(box["x"] + box["width"] + 50, box["y"] + box["height"] / 2)

            def metode_kontrol_baris():
                row_loc = page.locator("div, section, tr").filter(has_text="Tampilkan Anomali").filter(has_not_text="admin").last
                if row_loc.count() > 0:
                    for c in row_loc.locator('button, [role="switch"], label, input[type="checkbox"]').all():
                        try:
                            if c.is_visible():
                                c.click(force=True)
                                break
                        except Exception:
                            pass

            def metode_js():
                page.evaluate(JS_KLIK_SWITCH)

            for metode in (metode_koordinat, metode_kontrol_baris, metode_js):
                try:
                    metode()
                except Exception:
                    pass
                time.sleep(1.2)
                if page.evaluate(JS_CEK_SWITCH) == "true":
                    print("[FASIH] Switch anomali berhasil diaktifkan & terverifikasi.")
                    return True

        print("[WARN] Switch anomali belum terkonfirmasi aktif.")
        return False
    except Exception as e:
        print(f"[WARN] Kendala saat memeriksa switch anomali: {e}")
        return False


def perbaiki_anomali(page):
    """Recovery: tutup dialog, buka CATATAN, aktifkan switch anomali."""
    try:
        page.keyboard.press("Escape")
        time.sleep(0.8)
        page.keyboard.press("Escape")
        time.sleep(0.8)
    except Exception:
        pass
    buka_menu_catatan(page)
    time.sleep(1.5)
    ok = aktifkan_switch_anomali(page)
    time.sleep(1.5)
    return ok


# ============================================================
#  APPROVE
# ============================================================
RE_GALAT = re.compile(r"GALAT\s*[:\s]*(\d+)", re.I)
RE_BTN_KONFIRM = re.compile(r"^\s*(konfirmasi|ya|ya,?\s*approve|approve|setujui|lanjutkan|ok|setuju|kirim)\s*$", re.I)
SEL_DIALOG = 'div[role="dialog"], [role="alertdialog"], .modal.show'


def dialog_terlihat(page):
    try:
        dlg = page.locator(SEL_DIALOG)
        for i in range(dlg.count() - 1, -1, -1):
            if dlg.nth(i).is_visible():
                return dlg.nth(i)
    except Exception:
        pass
    return None


def tangani_dialog(page, tunggu=6):
    """
    Setelah klik Approve: tangani dialog konfirmasi (bisa berlapis).
    Return (hasil, teks, jumlah_galat); hasil: TANPA_DIALOG | OK | GALAT | TANPA_KONFIRM
    """
    teks_terakhir, dikonfirmasi = "", False
    for _ in range(3):
        d = None
        batas = time.time() + tunggu
        while time.time() < batas:
            d = dialog_terlihat(page)
            if d:
                break
            time.sleep(0.4)
        if not d:
            return ("OK" if dikonfirmasi else "TANPA_DIALOG"), teks_terakhir, 0

        try:
            teks = d.inner_text().strip()
        except Exception:
            teks = ""
        teks_terakhir = teks
        m = RE_GALAT.search(teks)
        if m and int(m.group(1)) > 0:
            return "GALAT", teks, int(m.group(1))

        klik = False
        try:
            for b in reversed(d.locator("button").all()):
                try:
                    if b.is_visible() and RE_BTN_KONFIRM.match(b.inner_text()):
                        print(f"[FASIH] Dialog konfirmasi: klik '{b.inner_text().strip()}'")
                        b.click()
                        klik = True
                        break
                except Exception:
                    pass
        except Exception:
            pass
        if not klik:
            return "TANPA_KONFIRM", teks, 0
        dikonfirmasi = True
        time.sleep(1.5)
    return "OK", teks_terakhir, 0


def verifikasi_approve(page, timeout=30, maks_masih_ada=20):
    """True bila Approve berhasil: tombol Approve hilang (atau teks 'Approved' muncul) dan tidak ada dialog galat."""
    mulai = time.time()
    masih_ada_sejak = None
    absen = 0
    while time.time() - mulai < timeout:
        try:
            d = dialog_terlihat(page)
            if d:
                m = RE_GALAT.search(d.inner_text())
                if m and int(m.group(1)) > 0:
                    return False
            if loading_terlihat(page):
                absen = 0
                time.sleep(1)
                continue
            items = baca_tombol(page)
            if cari_tombol_approve(items):
                absen = 0
                masih_ada_sejak = masih_ada_sejak or time.time()
                if time.time() - masih_ada_sejak >= maks_masih_ada:
                    return False
            else:
                masih_ada_sejak = None
                if teks_approved_terbaca(page):
                    return True
                absen += 1
                if absen >= 3:
                    return True
        except Exception:
            pass
        time.sleep(1)
    return False


def lakukan_approve(page):
    """Return (status, catatan): APPROVED | GALAT | GAGAL_APPROVE"""
    alasan_terakhir, status_terakhir = "", "GAGAL_APPROVE"
    for percobaan in range(1, MAX_PERCOBAAN_APPROVE + 1):
        print(f"[FASIH] Approve (percobaan {percobaan}/{MAX_PERCOBAAN_APPROVE})...")
        btn = cari_tombol_approve(baca_tombol(page))
        if btn is None:
            if teks_approved_terbaca(page):
                return "APPROVED", "tombol Approve sudah tidak ada, status Approved terbaca"
            return "GAGAL_APPROVE", "tombol Approve tidak ditemukan"

        if btn["disabled"]:
            alasan_terakhir, status_terakhir = "tombol Approve disabled", "GAGAL_APPROVE"
            print("[FASIH] Tombol Approve disabled.")
        else:
            klik_item(page, btn)
            time.sleep(1.2)
            hasil, teks, n = tangani_dialog(page)
            if hasil == "GALAT":
                alasan_terakhir = f"GALAT {n}: " + " | ".join(l.strip() for l in teks.splitlines() if l.strip())[:300]
                status_terakhir = "GALAT"
                print(f"[FASIH] Terdeteksi GALAT {n} saat Approve.")
                screenshot(page, "GALAT_APPROVE")
            elif hasil == "TANPA_KONFIRM":
                alasan_terakhir = "dialog tanpa tombol konfirmasi: " + teks[:200].replace("\n", " | ")
                status_terakhir = "GAGAL_APPROVE"
                screenshot(page, "DIALOG_ASING")
            else:
                if verifikasi_approve(page):
                    return "APPROVED", ("" if percobaan == 1 else f"berhasil setelah perbaikan anomali (percobaan {percobaan})")
                alasan_terakhir = "approve tidak terverifikasi (tombol masih ada / galat)"
                status_terakhir = "GAGAL_APPROVE"
                screenshot(page, "VERIFIKASI_GAGAL")

        if percobaan < MAX_PERCOBAAN_APPROVE:
            print("[FASIH] Mencoba perbaikan: aktifkan switch 'Tampilkan Anomali Usaha dan Keluarga' lalu Approve ulang...")
            perbaiki_anomali(page)

    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    return status_terakhir, alasan_terakhir


# ============================================================
#  PROSES SATU LINK
# ============================================================
def proses_link(page, row, dry_run=False, interaktif=False):
    link = row["link"]
    try:
        page.goto(link, timeout=60000, wait_until="domcontentloaded")
    except Exception as e:
        print(f"[WARN] goto: {str(e)[:100]}")

    state = tunggu_siap(page, link)

    if state == "BELUM":
        print("[FASIH] Status: BELUM approve (tombol Approve ada di layar).")
        if dry_run:
            return "DRY_RUN", "belum approve; tombol Approve ditemukan (tidak diklik)"
        if interaktif:
            if input("   Klik Approve sekarang? [y/N]: ").strip().lower() != "y":
                return "DILEWATI", "dibatalkan user"
        return lakukan_approve(page)

    if state == "SUDAH":
        print("[FASIH] Status: SUDAH approve.")
        return "SUDAH_APPROVE", ""

    if state == "ERROR_HALAMAN":
        screenshot(page, "ERROR_HALAMAN")
        return "ERROR_HALAMAN", "Forbidden/403/halaman error atau tidak bisa masuk ke assignment"

    if state == "TANPA_TOMBOL":
        items = baca_tombol(page)
        daftar = [f"{it['text'] or it['title'] or it['aria'] or '?'}" for it in items if it["kanan_bawah"]]
        screenshot(page, "TIDAK_DIKENALI")
        return "TIDAK_DIKENALI", "toolbar kanan-bawah tanpa tombol Approve & tanpa teks Approved. Tombol: " + "; ".join(daftar)[:250]

    screenshot(page, f"{state}")
    return ("ERROR_HALAMAN" if state == "TIMEOUT" else "GAGAL_APPROVE"), f"halaman tidak siap ({state})"


# ============================================================
#  MAIN
# ============================================================
def fmt_durasi(detik):
    detik = int(detik)
    return f"{detik // 3600}j {detik % 3600 // 60}m {detik % 60}d"


def main():
    print("=" * 65)
    print("              BOT FASIH - AUTO APPROVE ASSIGNMENT")
    print("=" * 65)

    path = cari_file_input()
    if not path:
        print(f"[ERROR] File Excel input tidak ditemukan. Taruh '{FILE_INPUT_DEFAULT}' di folder ini "
              "(kolom A = nama, kolom B = link_fasih).")
        return
    baris = baca_excel(path)
    if not baris:
        print("[ERROR] Tidak ada link FASIH yang terbaca dari Excel.")
        return

    hasil = muat_hasil()
    antrean, terlihat = [], set()
    dobel = 0
    for r in baris:
        if r["link"] in terlihat:
            dobel += 1
            continue
        terlihat.add(r["link"])
        if hasil.get(r["link"], ("",))[0] not in STATUS_FINAL:
            antrean.append(r)

    print(f"File input     : {os.path.basename(path)}")
    print(f"Total link     : {len(terlihat)}" + (f" ({dobel} link dobel diabaikan)" if dobel else ""))
    print(f"Sudah selesai  : {len(terlihat) - len(antrean)} (APPROVED / SUDAH_APPROVE di hasil_approve.csv)")
    print(f"Antrean        : {len(antrean)}")
    if not antrean:
        print("Semua link sudah selesai.")
        simpan_xlsx(baris)
        return

    u, _ = get_sso_credentials()
    if u:
        print(f"Akun SSO       : {u}")
    else:
        print("Akun SSO       : (file 'sso' belum diisi - bot hanya jalan kalau Chrome sudah login manual)")

    print("\nPilih Mode:")
    print("  [1] DRY-RUN  - cek status saja, TIDAK klik Approve [Default]")
    print("  [2] Approve 1 link, minta konfirmasi y/n sebelum klik")
    print("  [3] Batch   - approve N link pertama")
    print("  [4] Semua   - approve seluruh antrean")
    print("  [0] Keluar")
    pilih = input("Masukkan pilihan [1/2/3/4/0] (default 1): ").strip() or "1"

    dry_run, interaktif, limit = False, False, len(antrean)
    if pilih == "0":
        return
    elif pilih == "1":
        dry_run = True
        limit = max(1, int(input("Berapa link yang dicek? (default 3): ").strip() or "3"))
    elif pilih == "2":
        interaktif, limit = True, 1
    elif pilih == "3":
        limit = max(1, int(input("Berapa link? (default 10): ").strip() or "10"))
    elif pilih != "4":
        print("Pilihan tidak dikenal.")
        return
    antrean = antrean[:limit]

    if not start_debug_chrome():
        return

    stat = {}
    err_berturut = 0
    sukses_xlsx = 0
    t0 = time.time()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://localhost:{CDP_PORT}")
        context = browser.contexts[0]
        page = context.new_page()
        page.bring_to_front()

        try:
            for n, row in enumerate(antrean, 1):
                print("\n" + "-" * 65)
                print(f"[{n}/{len(antrean)}] (baris Excel {row['no']}) {row['nama']}")
                print(f"        {row['link']}")
                try:
                    status, catatan = proses_link(page, row, dry_run=dry_run, interaktif=interaktif)
                except SSOGagal as e:
                    print(f"\n[SSO ERROR] {e}")
                    break
                except KeyboardInterrupt:
                    raise
                except Exception as e:
                    status, catatan = "GAGAL_APPROVE", f"exception: {str(e)[:200]}"
                    screenshot(page, "EXCEPTION")

                stat[status] = stat.get(status, 0) + 1
                print(f"   => {status}" + (f"  ({catatan})" if catatan else ""))

                if not dry_run and status != "DILEWATI":
                    catat_hasil(row, status, catatan)
                    sukses_xlsx += 1
                    if sukses_xlsx % SIMPAN_XLSX_SETIAP == 0:
                        simpan_xlsx(baris)

                err_berturut = err_berturut + 1 if status in STATUS_ERROR else 0
                if err_berturut >= MAX_ERR_BERTURUT:
                    print(f"\n[STOP-LOSS] {MAX_ERR_BERTURUT} error berturut-turut. Bot dihentikan supaya tidak "
                          "menghabiskan antrean. Cek screenshots/ dan hasil_approve.csv, lalu jalankan lagi.")
                    break

                rata = (time.time() - t0) / n
                print(f"   (sisa {len(antrean) - n} link, perkiraan {fmt_durasi(rata * (len(antrean) - n))})")
                if n < len(antrean):
                    time.sleep(random.uniform(JEDA_ANTAR_LINK_MIN, JEDA_ANTAR_LINK_MAX))
        except KeyboardInterrupt:
            print("\n[INFO] Dihentikan user (Ctrl+C).")
        finally:
            if not dry_run:
                simpan_xlsx(baris)
            print("\n" + "=" * 65)
            print(f"RINGKASAN ({fmt_durasi(time.time() - t0)}):")
            for k, v in sorted(stat.items()):
                print(f"  {k:<16}: {v}")
            if not dry_run:
                print(f"Hasil lengkap  : {os.path.basename(HASIL_CSV)} / {os.path.basename(HASIL_XLSX)}")
            print("=" * 65)


if __name__ == "__main__":
    main()
