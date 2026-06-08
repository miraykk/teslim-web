import io
import os
import re
import shutil
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Optional, List, Tuple

import streamlit as st
from openpyxl import load_workbook
from pypdf import PdfReader, PdfWriter

# ── Sayfa yapılandırması ────────────────────────────────────────────────────
st.set_page_config(
    page_title="Teslim Tesellüm Oluşturucu",
    page_icon="📄",
    layout="wide",
)

# ── Sabitler ────────────────────────────────────────────────────────────────
FIRMA_LISTESI = ["akakçe", "cevher", "çıtır", "altınsa"]
ISLEM_LISTESI = ["Alış", "Satış"]
BIRLESTIRME_LISTESI = ["Tüm PDF'ler ortak", "TC / Vergi No'ya göre gruplu"]

DEFAULT_COL_MAP = {
    "TARIH_COL": 1,
    "FATURA_COL": 2,
    "ISIM_COL": 3,
    "ADRES_COL": 4,
    "TC1_COL": 5,
    "TC2_COL": 6,
    "GR_COL": 8,
    "TL_COL": 10,
}

TEMPLATES_DIR = Path(__file__).parent / "templates"

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
    if name is None:
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


def template_path(firma: str, islem: str) -> Path:
    f = normalize_tr(firma)
    t = "alis" if islem.lower().startswith("al") else "satis"
    return TEMPLATES_DIR / f"{f}_{t}.xlsx"


def get_headers(file_bytes: bytes) -> List[str]:
    wb = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active
    headers = []
    for cell in ws[1]:
        v = cell.value
        headers.append("" if v is None else str(v).strip())
    wb.close()
    return headers


def header_to_col_index(headers: List[str], chosen: str) -> Optional[int]:
    chosen = (chosen or "").strip()
    if not chosen or chosen == "Boş bırak":
        return None
    for i, h in enumerate(headers, 1):
        if (h or "").strip() == chosen:
            return i
    return None


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
    env["HOME"] = str(pdf_dir)  # LibreOffice profil sorunu için

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--nodefault",
        "--nofirststartwizard",
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

def run_job(firma, islem_turu, input_bytes, col_map, group_mode, progress_bar, status_text) -> bytes:
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice bulunamadı. Sunucuda kurulu olmayabilir."
        )

    tpl = template_path(firma, islem_turu)
    if not tpl.exists():
        raise FileNotFoundError(
            f"Şablon bulunamadı: templates/{normalize_tr(firma)}_"
            f"{'alis' if islem_turu == 'Alış' else 'satis'}.xlsx\n"
            f"Bu dosyayı templates/ klasörüne eklemeniz gerekiyor."
        )

    wb_in = load_workbook(io.BytesIO(input_bytes), data_only=True)
    ws_in = wb_in.active

    TARIH_COL  = col_map.get("TARIH_COL")
    FATURA_COL = col_map.get("FATURA_COL")
    ISIM_COL   = col_map.get("ISIM_COL")
    ADRES_COL  = col_map.get("ADRES_COL")
    TC1_COL    = col_map.get("TC1_COL")
    TC2_COL    = col_map.get("TC2_COL")
    GR_COL     = col_map.get("GR_COL")
    TL_COL     = col_map.get("TL_COL")
    TC_COLS    = [c for c in [TC1_COL, TC2_COL] if c]

    valid_rows = []
    for r in range(2, ws_in.max_row + 1):
        if FATURA_COL:
            fatura = ws_in.cell(r, FATURA_COL).value
            if fatura is None or str(fatura).strip() == "":
                continue
        valid_rows.append(r)

    if not valid_rows:
        raise ValueError("İşlenecek satır bulunamadı. Fatura No kolonunu kontrol edin.")

    total = len(valid_rows)

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        excel_dir = tmp / "exceller"
        pdf_dir   = tmp / "pdfler"
        excel_dir.mkdir()
        pdf_dir.mkdir()

        pdf_meta = {}  # fname -> (tarih_key, fatura_str, tc_or_vkn, pdf_path)

        for i, r in enumerate(valid_rows):
            tarih     = ws_in.cell(r, TARIH_COL).value  if TARIH_COL  else None
            fatura    = ws_in.cell(r, FATURA_COL).value if FATURA_COL else f"satir_{r}"
            isim      = ws_in.cell(r, ISIM_COL).value   if ISIM_COL   else None
            adres     = ws_in.cell(r, ADRES_COL).value  if ADRES_COL  else None
            tc_or_vkn = first_non_empty(ws_in, r, TC_COLS) if TC_COLS else None
            gr        = ws_in.cell(r, GR_COL).value     if GR_COL     else None
            tl        = ws_in.cell(r, TL_COL).value     if TL_COL     else None
            kisa_isim = kisa_firma_adi(isim)

            # Şablonu doldur
            wb_tmp = load_workbook(tpl)
            ws_tmp = wb_tmp.active

            if normalize_tr(firma) == "akakce":
                if tarih     is not None: ws_tmp["E6"].value  = tarih
                ws_tmp["E7"].value = "Fatih / İstanbul"
                if fatura    is not None: ws_tmp["E8"].value  = fatura
                if gr        is not None: ws_tmp["E9"].value  = gr
                if tl        is not None: ws_tmp["E10"].value = tl
                if kisa_isim:             ws_tmp["E15"].value = kisa_isim
                if tc_or_vkn is not None: ws_tmp["E16"].value = tc_or_vkn
                if gr        is not None: ws_tmp["E17"].value = gr
                if adres     is not None: ws_tmp["E18"].value = adres
                ws_tmp["E10"].number_format = '#,##0.00'
                ws_tmp["E9"].number_format  = '#,##0.0000'
            else:
                if tarih     is not None: ws_tmp["E5"].value  = tarih
                if fatura    is not None: ws_tmp["E6"].value  = fatura
                if gr        is not None: ws_tmp["E7"].value  = gr
                if tl        is not None: ws_tmp["E8"].value  = tl
                if kisa_isim:             ws_tmp["E14"].value = kisa_isim
                if tc_or_vkn is not None: ws_tmp["E15"].value = tc_or_vkn
                ws_tmp["E8"].number_format = '#,##0.00'
                ws_tmp["E7"].number_format = '#,##0.0000'

            fname = safe_filename(fatura)
            out_xlsx = excel_dir / f"{fname}.xlsx"
            wb_tmp.save(out_xlsx)

            # LibreOffice ile PDF'e çevir
            pdf_path = convert_xlsx_to_pdf(soffice, out_xlsx, pdf_dir)

            tarih_key = tarih.isoformat() if hasattr(tarih, "isoformat") else str(tarih or "")
            pdf_meta[fname] = (tarih_key, str(fatura), tc_or_vkn, pdf_path)

            progress_bar.progress(int((i + 1) / total * 88))
            status_text.text(f"İşleniyor: {i+1}/{total} — {fname}")

        status_text.text("ZIP oluşturuluyor…")
        progress_bar.progress(92)

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            # Ayrı PDF'leri ekle
            for fname, (_, _, _, pdf_path) in pdf_meta.items():
                zf.write(pdf_path, f"pdfler/{fname}.pdf")

            if group_mode == "Tüm PDF'ler ortak":
                items_sorted = sorted(
                    pdf_meta.items(),
                    key=lambda x: (x[1][0], x[1][1])
                )
                merged = merge_pdf_list([v[3] for _, v in items_sorted])
                zf.writestr("tum_pdfler.pdf", merged)
            else:
                groups = {}
                for fname, (tarih_key, fatura_str, tc_or_vkn, pdf_path) in pdf_meta.items():
                    gkey = safe_filename(str(tc_or_vkn)) if tc_or_vkn else "NO_TC_VKN"
                    groups.setdefault(gkey, []).append((tarih_key, fatura_str, pdf_path))
                for gkey, items in groups.items():
                    items_s = sorted(items, key=lambda x: (x[0], x[1]))
                    merged = merge_pdf_list([it[2] for it in items_s])
                    zf.writestr(f"gruplu_pdfler/{gkey}.pdf", merged)

        progress_bar.progress(100)
        status_text.text(f"✅ Tamamlandı! {total} belge işlendi.")
        zip_buf.seek(0)
        return zip_buf.read()


# ── Streamlit Arayüzü ───────────────────────────────────────────────────────

def main():
    st.title("📄 Teslim Tesellüm Oluşturucu")
    st.caption("Excel listesinden her satır için teslim-tesellüm belgesi (PDF) üretir.")

    st.divider()

    col1, col2, col3 = st.columns([2, 2, 3])
    with col1:
        firma = st.selectbox("Firma", FIRMA_LISTESI)
    with col2:
        islem_turu = st.selectbox("İşlem Türü", ISLEM_LISTESI)
    with col3:
        group_mode = st.selectbox("PDF Birleştirme", BIRLESTIRME_LISTESI)

    st.divider()

    uploaded = st.file_uploader(
        "📂 Esnek Rapor (.xlsx) yükle",
        help="İlk satır başlık olmalı. Fatura No kolonu boş olan satırlar atlanır.",
    )

    headers = []
    file_bytes = None

    if uploaded is not None:
        file_bytes = uploaded.read()
        try:
            headers = get_headers(file_bytes)
            st.success(f"✔ Dosya okundu — {len(headers)} kolon bulundu.")
        except Exception as e:
            st.error(f"Dosya okunamadı: {e}")
            return

    st.divider()

    with st.expander("⚙️ Kolon Eşleme (Opsiyonel)", expanded=bool(headers)):
        st.caption("Başlık isimleriyle eşleştirin. Boş bırakılanlar varsayılan sıraya göre atanır.")

        header_opts = ["Boş bırak"] + [h for h in headers if h.strip()]

        fields = [
            ("Tarih",            "TARIH_COL",  1),
            ("Fatura No",        "FATURA_COL", 2),
            ("İsim / Unvan",     "ISIM_COL",   3),
            ("Adres",            "ADRES_COL",  4),
            ("TC Kimlik No",     "TC1_COL",    5),
            ("Vergi No (varsa)", "TC2_COL",    6),
            ("Gram",             "GR_COL",     8),
            ("Tutar (TL)",       "TL_COL",     10),
        ]

        col_selections = {}
        cola, colb = st.columns(2)
        for idx, (label, key, default_col) in enumerate(fields):
            container = cola if idx % 2 == 0 else colb
            with container:
                if headers and 1 <= default_col <= len(headers) and headers[default_col - 1].strip():
                    dval = headers[default_col - 1]
                    didx = header_opts.index(dval) if dval in header_opts else 0
                else:
                    didx = 0
                sel = st.selectbox(
                    label, options=header_opts, index=didx,
                    key=f"col_{key}", disabled=not headers,
                )
                col_selections[key] = sel

    # Kolon indekslerini çöz
    col_map = {}
    for key, chosen in col_selections.items():
        col_map[key] = header_to_col_index(headers, chosen)

    if not any(v is not None for v in col_map.values()):
        col_map = dict(DEFAULT_COL_MAP)

    st.divider()

    # Şablon kontrolü
    tpl = template_path(firma, islem_turu)
    if not tpl.exists():
        st.warning(
            f"⚠️ **{firma} / {islem_turu}** şablonu henüz yüklenmemiş. "
            f"`templates/{normalize_tr(firma)}_{'alis' if islem_turu=='Alış' else 'satis'}.xlsx` "
            f"dosyasını repoya ekleyin.",
        )

    if st.button("🚀 Çalıştır", type="primary", disabled=file_bytes is None, use_container_width=True):
        progress_bar = st.progress(0)
        status_text  = st.empty()

        try:
            zip_bytes = run_job(
                firma=firma,
                islem_turu=islem_turu,
                input_bytes=file_bytes,
                col_map=col_map,
                group_mode=group_mode,
                progress_bar=progress_bar,
                status_text=status_text,
            )
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
