import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Optional, List

import streamlit as st
from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter

# ── Sayfa yapılandırması ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Teslim Tesellüm Oluşturucu",
    page_icon="📄",
    layout="centered",
)

# ── Sabit kolon haritası (değiştirmek istersen burası) ──────────────────────
# Excel listenizdeki sütun numaraları (1 = A, 2 = B, ...)
COL_MAP = {
    "TARIH_COL":  1,   # Tarih
    "FATURA_COL": 2,   # Fatura No
    "ISIM_COL":   3,   # Cari Hesap (isim/unvan)
    "ADRES_COL":  4,   # Adres
    "TC1_COL":    5,   # TC Kimlik No (şahıs)
    "TC2_COL":    6,   # Vergi No (şirket) — hangisi doluysa o alınır
    "GR_COL":     8,   # Toplam Miktar (gram)
    "TL_COL":     10,  # Brüt Toplam (TL)
}

# Şablondaki hücre adresleri — akakçe şablonu için
AKAKCE_CELLS = {
    "tarih":  "E6",
    "yer":    "E7",   # sabit "Fatih / İstanbul"
    "fatura": "E8",
    "gr":     "E9",
    "tl":     "E10",
    "isim":   "E15",
    "tc":     "E16",
    "gr2":    "E17",  # gram tekrar (şablonda iki yerde varsa)
    "adres":  "E18",
}

ISLEM_LISTESI      = ["Alış", "Satış"]
BIRLESTIRME_LISTESI = ["Tüm PDF'ler ortak", "TC / Vergi No'ya göre gruplu"]
TEMPLATES_DIR      = Path(__file__).parent / "templates"

# ── Yardımcı fonksiyonlar ───────────────────────────────────────────────────

def normalize_tr(s: str) -> str:
    s = (s or "").lower().strip()
    return (s.replace("ı", "i").replace("ç", "c").replace("ğ", "g")
              .replace("ş", "s").replace("ö", "o").replace("ü", "u"))


def safe_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", "", name)
    return name or "bos_fatura"


def first_non_empty(ws, row: int, cols: List[int]):
    for c in cols:
        v = ws.cell(row, c).value
        if v is not None and str(v).strip() != "":
            return v
    return None


def kisa_firma_adi(name: str) -> str:
    if not name:
        return ""
    s = str(name).strip().upper()
    replacements = {
        "LİMİTED": "LTD.", "LIMITED": "LTD.", "ŞİRKETİ": "ŞTİ.",
        "SIRKETI": "ŞTİ.", "ANONİM": "A.Ş.", "ANONIM": "A.Ş.",
        "SANAYİ": "SAN.", "SANAYI": "SAN.", "TİCARET": "TİC.",
        "TICARET": "TİC.", "PAZARLAMA": "PAZ.", "HİZMETLERİ": "HİZ.",
        "İNŞAAT": "İNŞ.", "NAKLİYE": "NAK.", "LOJİSTİK": "LOJ.",
        "TEKNOLOJİ": "TEK.", "MÜHENDİSLİK": "MÜH.", "DANIŞMANLIK": "DAN.",
        "MUHASEBE": "MUH.", "SİGORTA": "SIG.", "GAYRİMENKUL": "GMK.",
    }
    for old, new in sorted(replacements.items(), key=lambda x: len(x[0]), reverse=True):
        s = re.sub(rf"\b{re.escape(old)}\b", new, s)
    return re.sub(r"\s+", " ", s).strip()


def find_soffice() -> Optional[str]:
    candidates = [
        "soffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.isabs(c):
            if os.path.exists(c):
                return c
        else:
            if shutil.which(c):
                return c
    return None


def convert_xlsx_to_pdf(soffice: str, xlsx_path: Path, pdf_dir: Path) -> Path:
    pdf_dir.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    env["HOME"] = str(pdf_dir)
    cmd = [
        soffice, "--headless", "--nologo", "--nolockcheck",
        "--nodefault", "--nofirststartwizard",
        "--convert-to", "pdf",
        "--outdir", str(pdf_dir),
        str(xlsx_path),
    ]
    result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                            text=True, env=env, timeout=60)
    if result.returncode != 0:
        raise RuntimeError(f"LibreOffice hatası:\n{result.stderr or result.stdout}")
    out_pdf = pdf_dir / (xlsx_path.stem + ".pdf")
    if not out_pdf.exists():
        pdfs = sorted(pdf_dir.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
        if not pdfs:
            raise RuntimeError("PDF bulunamadı — LibreOffice çıktı üretmedi.")
        out_pdf = pdfs[0]
    return out_pdf


def merge_pdf_list(pdf_paths: List[Path]) -> bytes:
    writer = PdfWriter()
    for p in pdf_paths:
        reader = PdfReader(str(p))
        for page in reader.pages:
            writer.add_page(page)
    buf = io.BytesIO()
    writer.write(buf)
    buf.seek(0)
    return buf.read()


# ── Ana iş mantığı ──────────────────────────────────────────────────────────

def run_job(islem_turu, input_bytes, group_mode, progress_bar, status_text) -> bytes:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice bulunamadı. Sunucuda kurulu olmayabilir.")

    tpl_name = f"akakce_{'alis' if islem_turu == 'Alış' else 'satis'}.xlsx"
    tpl = TEMPLATES_DIR / tpl_name
    if not tpl.exists():
        raise FileNotFoundError(f"Şablon bulunamadı: templates/{tpl_name}")

    wb_in = load_workbook(io.BytesIO(input_bytes), data_only=True)
    ws_in = wb_in.active

    TARIH_COL  = COL_MAP["TARIH_COL"]
    FATURA_COL = COL_MAP["FATURA_COL"]
    ISIM_COL   = COL_MAP["ISIM_COL"]
    ADRES_COL  = COL_MAP["ADRES_COL"]
    TC_COLS    = [COL_MAP["TC1_COL"], COL_MAP["TC2_COL"]]
    GR_COL     = COL_MAP["GR_COL"]
    TL_COL     = COL_MAP["TL_COL"]

    # Geçerli satırları topla (fatura no dolu olanlar)
    valid_rows = []
    for r in range(2, ws_in.max_row + 1):
        fatura = ws_in.cell(r, FATURA_COL).value
        if fatura is not None and str(fatura).strip() != "":
            valid_rows.append(r)

    if not valid_rows:
        raise ValueError("İşlenecek satır bulunamadı. Fatura No kolonunu kontrol edin.")

    total = len(valid_rows)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp      = Path(tmpdir)
        excel_dir = tmp / "exceller"
        pdf_dir   = tmp / "pdfler"
        excel_dir.mkdir()
        pdf_dir.mkdir()

        pdf_meta = {}  # fname -> (tarih_key, fatura_str, tc_or_vkn, pdf_path)

        for i, r in enumerate(valid_rows):
            tarih     = ws_in.cell(r, TARIH_COL).value
            fatura    = ws_in.cell(r, FATURA_COL).value
            isim      = ws_in.cell(r, ISIM_COL).value
            adres     = ws_in.cell(r, ADRES_COL).value
            tc_or_vkn = first_non_empty(ws_in, r, TC_COLS)
            gr        = ws_in.cell(r, GR_COL).value
            tl        = ws_in.cell(r, TL_COL).value
            kisa_isim = kisa_firma_adi(isim)

            # Şablonu doldur
            wb_tmp = load_workbook(tpl)
            ws_tmp = wb_tmp.active

            c = AKAKCE_CELLS
            if tarih     is not None: ws_tmp[c["tarih"]].value  = tarih
            ws_tmp[c["yer"]].value = "Fatih / İstanbul"
            if fatura    is not None: ws_tmp[c["fatura"]].value = fatura
            if gr        is not None: ws_tmp[c["gr"]].value     = gr
            if tl        is not None: ws_tmp[c["tl"]].value     = tl
            if kisa_isim:             ws_tmp[c["isim"]].value   = kisa_isim
            if tc_or_vkn is not None: ws_tmp[c["tc"]].value     = tc_or_vkn
            if gr        is not None: ws_tmp[c["gr2"]].value    = gr
            if adres     is not None: ws_tmp[c["adres"]].value  = adres

            ws_tmp[c["tl"]].number_format = '#,##0.00'
            ws_tmp[c["gr"]].number_format = '#,##0.0000'

            fname    = safe_filename(fatura)
            out_xlsx = excel_dir / f"{fname}.xlsx"
            wb_tmp.save(out_xlsx)

            pdf_path  = convert_xlsx_to_pdf(soffice, out_xlsx, pdf_dir)
            tarih_key = tarih.isoformat() if hasattr(tarih, "isoformat") else str(tarih or "")
            pdf_meta[fname] = (tarih_key, str(fatura), tc_or_vkn, pdf_path)

            progress_bar.progress(int((i + 1) / total * 88))
            status_text.text(f"İşleniyor: {i+1}/{total} — {fname}")

        status_text.text("ZIP oluşturuluyor…")
        progress_bar.progress(92)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, (_, _, _, pdf_path) in pdf_meta.items():
                zf.write(pdf_path, f"pdfler/{fname}.pdf")

            if group_mode == "Tüm PDF'ler ortak":
                items_sorted = sorted(pdf_meta.items(), key=lambda x: (x[1][0], x[1][1]))
                merged = merge_pdf_list([v[3] for _, v in items_sorted])
                zf.writestr("tum_pdfler.pdf", merged)
            else:
                groups = {}
                for fname, (tarih_key, fatura_str, tc_or_vkn, pdf_path) in pdf_meta.items():
                    gkey = safe_filename(str(tc_or_vkn)) if tc_or_vkn else "NO_TC_VKN"
                    groups.setdefault(gkey, []).append((tarih_key, fatura_str, pdf_path))
                for gkey, items in groups.items():
                    merged = merge_pdf_list([it[2] for it in sorted(items, key=lambda x: (x[0], x[1]))])
                    zf.writestr(f"gruplu_pdfler/{gkey}.pdf", merged)

        progress_bar.progress(100)
        status_text.text(f"✅ Tamamlandı! {total} belge işlendi.")
        zip_buf.seek(0)
        return zip_buf.read()


# ── Streamlit Arayüzü ───────────────────────────────────────────────────────

def main():
    st.title("📄 Teslim Tesellüm Oluşturucu")
    st.caption("Akakçe — Excel listesinden otomatik teslim-tesellüm belgesi üretir.")

    st.divider()

    col1, col2 = st.columns(2)
    with col1:
        islem_turu = st.selectbox("İşlem Türü", ISLEM_LISTESI)
    with col2:
        group_mode = st.selectbox("PDF Birleştirme", BIRLESTIRME_LISTESI)

    st.divider()

    uploaded = st.file_uploader("📂 Esnek Rapor (.xlsx) yükle")

    file_bytes = None
    if uploaded is not None:
        file_bytes = uploaded.read()
        # Satır sayısını göster
        try:
            wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
            ws = wb.active
            row_count = sum(
                1 for r in range(2, ws.max_row + 1)
                if ws.cell(r, COL_MAP["FATURA_COL"]).value not in (None, "")
            )
            wb.close()
            st.success(f"✔ Dosya okundu — {row_count} kayıt bulundu.")
        except Exception as e:
            st.error(f"Dosya okunamadı: {e}")
            return

    st.divider()

    if st.button("🚀 Çalıştır", type="primary", disabled=file_bytes is None, use_container_width=True):
        progress_bar = st.progress(0)
        status_text  = st.empty()
        try:
            zip_bytes = run_job(islem_turu, file_bytes, group_mode, progress_bar, status_text)
            st.download_button(
                label="⬇️ ZIP Dosyasını İndir",
                data=zip_bytes,
                file_name="teslim_tesellum_ciktilari.zip",
                mime="application/zip",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"❌ Hata: {e}")


if __name__ == "__main__":
    main()
